# -*- coding: utf-8 -*-
"""小虾记账 后端API服务
封装 accounting_report 引擎，提供REST API供前端调用
"""
import sys
import os
import time
import traceback
import asyncio
import threading
import uuid
from pathlib import Path
from io import BytesIO
from typing import Optional
import shutil
import json
import datetime

# 确保项目根在 sys.path
def _resource_root() -> Path:
    """随程序分发的只读资源目录（打包后为 PyInstaller 解包目录）"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent.parent


RESOURCE_ROOT = _resource_root()

# 数据根目录（输入/输出/归档/config 所在位置，必须可写）
# - 桌面版：Electron 通过 XJZ_DATA_DIR 指定（默认 我的文档\小虾记账）
# - 开发模式：项目根目录
_ENV_DATA_ROOT = os.environ.get("XJZ_DATA_DIR")
if _ENV_DATA_ROOT:
    ROOT = Path(_ENV_DATA_ROOT)
elif getattr(sys, "frozen", False):
    ROOT = Path(os.environ.get("APPDATA") or Path.home()) / "小虾记账"
else:
    ROOT = Path(__file__).resolve().parent.parent

if str(RESOURCE_ROOT) not in sys.path:
    sys.path.insert(0, str(RESOURCE_ROOT))


def _bootstrap_data_dir() -> None:
    """首次启动时初始化数据目录：建目录 + 复制可编辑的规则文件"""
    try:
        for sub in ("输入", "输出", "归档"):
            (ROOT / sub).mkdir(parents=True, exist_ok=True)

        # 规则文件复制到数据目录，允许用户在「高级配置」里改
        bundled_rules = RESOURCE_ROOT / "accounting_report" / "rules"
        runtime_rules = ROOT / "accounting_report" / "rules"
        if bundled_rules.is_dir() and not runtime_rules.is_dir():
            shutil.copytree(bundled_rules, runtime_rules)
    except Exception as exc:  # 初始化失败不应阻止服务启动
        print(f"[bootstrap] 数据目录初始化警告: {exc}")


_bootstrap_data_dir()

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from pydantic import BaseModel
import uvicorn

# 上传文件安全限制
ALLOWED_UPLOAD_EXT = {".xls", ".xlsx", ".pdf", ".jpg", ".jpeg", ".png", ".csv", ".txt"}
MAX_UPLOAD_SIZE = 50 * 1024 * 1024  # 50MB

# ---- App ----
app = FastAPI(
    title="小虾记账 API",
    version="1.0.0",
    docs_url="/api/docs",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---- Models ----


class ComputeRequest(BaseModel):
    year: int
    month: int
    entity_type: Optional[str] = None  # None→读 config.json; "corp" | "individual"


class ComputeResponse(BaseModel):
    success: bool
    month_label: str
    outputs: list[dict]
    errors: list[str]
    warnings: list[str]
    stats: dict


class PeriodInfo(BaseModel):
    period: str
    has_data: bool
    output_count: int


@app.get("/api/health")
async def health():
    return {"status": "ok", "company": _get_company_name()}


@app.get("/api/diag")
async def diag():
    """运行环境自检：确认打包完整性（供售后排障与验收使用）"""
    modules = {}
    for name in ("fitz", "pymupdf", "pdfplumber", "openpyxl", "xlrd", "yaml", "fastapi", "uvicorn"):
        try:
            mod = __import__(name)
            modules[name] = getattr(mod, "__version__", "ok")
        except Exception as exc:  # noqa: BLE001
            modules[name] = f"MISSING: {exc}"
    return {
        "frozen": bool(getattr(sys, "frozen", False)),
        "python": sys.version.split()[0],
        "data_root": str(ROOT),
        "resource_root": str(RESOURCE_ROOT),
        "modules": modules,
    }


@app.get("/api/periods")
async def list_periods():
    """列出所有有数据的账期"""
    periods = []
    input_dir = ROOT / "输入"
    output_dir = ROOT / "输出"

    seen = set()
    # 从输出目录发现
    if output_dir.exists():
        for f in output_dir.iterdir():
            if f.is_dir() and len(f.name) >= 7 and f.name[4] == "-":
                seen.add(f.name)
            elif f.name[:7].replace("-", "").isdigit() and len(f.name) >= 7:
                period = f.name[:7]
                seen.add(period)

    # 从输入目录发现
    if input_dir.exists():
        for f in input_dir.rglob("*"):
            if f.is_file():
                # 尝试从文件名中提取YYYY-MM
                for part in f.name.replace("_", "-").split():
                    if len(part) >= 7 and part[4] == "-" and part[:7].replace("-", "").isdigit():
                        seen.add(part[:7])

    for p in sorted(seen, reverse=True):
        out_files = []
        out_path = output_dir / p
        if out_path.is_dir():
            out_files = list(out_path.iterdir())
        periods.append(PeriodInfo(
            period=p,
            has_data=(input_dir / p).exists() or bool(out_files),
            output_count=len(out_files),
        ))

    if not periods:
        # 返回当前月作为默认
        now = datetime.date.today()
        p = f"{now.year}-{now.month:02d}"
        periods = [PeriodInfo(period=p, has_data=False, output_count=0)]

    return [p.model_dump() for p in periods]


@app.post("/api/upload")
async def upload_file(
    period: str = Form(...),
    module: str = Form(...),
    file: UploadFile = File(...),
):
    """上传原始资料文件到指定账期的输入目录"""
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}，应为 YYYY-MM")

    # 模块目录映射
    module_dirs = {
        "revenue": "收入发票",
        "expenses": "成本费用发票",
        "bank": "银行对账单",
        "payroll": "工资社保",
        "company-info": "企业信息",
        "tax-adjust": "税务调整",
    }

    if module not in module_dirs:
        raise HTTPException(400, f"不支持的模块: {module}，可选: {list(module_dirs)}")
    subdir = module_dirs[module]
    target_dir = ROOT / "输入" / period / subdir
    target_dir.mkdir(parents=True, exist_ok=True)

    # 安全文件名 + 类型/大小校验（服务端强制，防绕过前端限制）
    safe_name = Path(file.filename or "unknown").name
    ext = Path(safe_name).suffix.lower()
    if ext not in ALLOWED_UPLOAD_EXT:
        raise HTTPException(400, f"不支持的文件类型: {ext}，允许: {sorted(ALLOWED_UPLOAD_EXT)}")
    target_path = target_dir / safe_name

    # 去重：重名加序号
    counter = 1
    stem, ext = Path(safe_name).stem, Path(safe_name).suffix
    while target_path.exists():
        target_path = target_dir / f"{stem}({counter}){ext}"
        counter += 1

    # 流式落盘 + 边读边累计长度，超限立即中止（避免超大文件全量读入内存 OOM）
    total = 0
    CHUNK = 1024 * 1024  # 1MB

    def _cleanup_partial() -> None:
        # 清理失败（如沙箱回收站不可用）不得掩盖原始 413/异常语义
        try:
            target_path.unlink(missing_ok=True)
        except OSError:
            pass

    try:
        with open(target_path, "wb") as fh:
            while True:
                chunk = await file.read(CHUNK)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_UPLOAD_SIZE:
                    _cleanup_partial()
                    raise HTTPException(413, f"文件过大: {total / 1024 / 1024:.1f}MB > 50MB 限制")
                fh.write(chunk)
    except HTTPException:
        raise
    except Exception:
        _cleanup_partial()
        raise

    return {
        "success": True,
        "filename": target_path.name,
        "path": str(target_path.relative_to(ROOT)),
        "size": total,
    }


# 后台任务表（进程内）。任务完成后保留结果供前端轮询，超时后自动淘汰，
# 防止常驻内存无限增长；同时限制同一账期并发，避免多线程同时写同一输出目录。
_COMPUTE_TASKS: dict[str, dict] = {}
_TASK_TTL_SECONDS = 30 * 60  # 任务完成/失败后保留 30 分钟


def _cleanup_expired_tasks() -> None:
    """删除已过期（超过 TTL）的任务，保持内存有界。"""
    now = time.time()
    expired = [
        tid for tid, state in _COMPUTE_TASKS.items()
        if state.get("expires_at") is not None and now > state["expires_at"]
    ]
    for tid in expired:
        _COMPUTE_TASKS.pop(tid, None)


def _running_task_for_period(year: int, month: int) -> Optional[str]:
    """返回指定账期正在运行/排队中的任务 id（用于对新任务去重/复用）

    含 pending：任务创建与后台线程启动之间有一小段窗口，此时再点「生成」会
    开出第二个并发任务写同一输出目录；pending 任务同样复用。"""
    for tid, state in _COMPUTE_TASKS.items():
        if state.get("status") in ("running", "pending") \
                and state.get("year") == year and state.get("month") == month:
            return tid
    return None


@app.post("/api/compute/task")
async def start_compute_task(req: ComputeRequest):
    """创建异步计算任务：立即返回 task_id，前端轮询进度（真实进度，非模拟）"""
    _cleanup_expired_tasks()

    # 同一账期已有运行中任务 → 直接复用，避免并发线程写同一输出目录互相覆盖
    existing = _running_task_for_period(req.year, req.month)
    if existing is not None:
        return {"task_id": existing, "status": "running", "reused": True}

    task_id = uuid.uuid4().hex[:12]
    _COMPUTE_TASKS[task_id] = {
        "task_id": task_id,
        "status": "pending",      # pending | running | done | error
        "progress": 0,
        "message": "排队中",
        "result": None,
        "error": None,
        "year": req.year,
        "month": req.month,
        "expires_at": None,       # 完成后设置，用于 TTL 淘汰
    }

    def _run() -> None:
        state = _COMPUTE_TASKS[task_id]
        state["status"] = "running"
        state["message"] = "启动中"
        try:
            from accounting_report.full_pipeline import run_full
            from accounting_report.config import OUTPUT_DIR, INPUT_DIR

            entity_type = req.entity_type or _get_config_entity_type()

            def cb(pct: int, msg: str) -> None:
                state["progress"] = pct
                state["message"] = msg

            result = run_full(
                req.year, req.month, INPUT_DIR, OUTPUT_DIR,
                entity_type=entity_type, progress=cb,
            )
            outputs = _build_output_list(result)
            state["result"] = {
                "outputs": outputs,
                "errors": result.errors,
                "warnings": result.query.warnings if result.query else [],
                "stats": _build_stats(result),
            }
            state["progress"] = 100
            state["status"] = "done" if not result.errors else "done_with_errors"
        except Exception as e:
            state["status"] = "error"
            state["error"] = f"{e}"
            state["message"] = "运算失败"
        finally:
            # 完成后开始 TTL 计时，之后每次查询都会走过期清理
            state["expires_at"] = time.time() + _TASK_TTL_SECONDS

    threading.Thread(target=_run, daemon=True).start()
    return {"task_id": task_id, "status": "pending"}


@app.get("/api/compute/task/{task_id}")
async def get_compute_task(task_id: str):
    """查询任务进度与结果"""
    _cleanup_expired_tasks()
    state = _COMPUTE_TASKS.get(task_id)
    if state is None:
        raise HTTPException(404, f"任务不存在: {task_id}")
    return state


@app.post("/api/compute", response_model=ComputeResponse)
async def compute(req: ComputeRequest):
    """触发完整账务运算"""
    try:
        from accounting_report.full_pipeline import run_full
        from accounting_report.config import OUTPUT_DIR, INPUT_DIR
    except Exception as e:
        raise HTTPException(500, f"加载 accounting_report 模块失败: {e}")

    input_dir = INPUT_DIR
    output_dir = OUTPUT_DIR

    try:
        entity_type = req.entity_type or _get_config_entity_type()
        # 同步引擎放入线程池，避免阻塞事件循环（否则运算期间 health/upload/download 全部卡死）
        result = await asyncio.to_thread(
            run_full, req.year, req.month, input_dir, output_dir, entity_type=entity_type
        )
    except Exception as e:
        tb = traceback.format_exc()
        return ComputeResponse(
            success=False,
            month_label=f"{req.year}-{req.month:02d}",
            outputs=[],
            errors=[f"运算异常: {e}（详见服务端日志）"],
            warnings=[],
            stats={},
        )

    outputs = _build_output_list(result)
    stats = _build_stats(result)

    return ComputeResponse(
        success=len(result.errors) == 0,
        month_label=f"{req.year}-{req.month:02d}",
        outputs=outputs,
        errors=result.errors,
        warnings=result.query.warnings if result.query else [],
        stats=stats,
    )


@app.get("/api/results/{period}")
async def list_results(period: str):
    """列出指定账期的结果文件"""
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}")

    output_dir = ROOT / "输出"
    files: dict[str, dict] = {}  # 按文件名去重（目录内 + 根前缀匹配可能重复列出）

    # 直接目录
    period_dir = output_dir / period
    if period_dir.is_dir():
        for f in period_dir.iterdir():
            if f.is_file():
                files[f.name] = {
                    "name": f.name,
                    "size": f.stat().st_size,
                    "ext": f.suffix.lower(),
                }

    # 也查找前缀匹配的文件（如 "2026-07 会计报表.xlsx"）
    for f in output_dir.iterdir():
        if f.is_file() and f.name.startswith(period):
            files[f.name] = {
                "name": f.name,
                "size": f.stat().st_size,
                "ext": f.suffix.lower(),
            }

    # 稳定排序：报表类优先再按名称（避免同名文件顺序抖动）
    ordered = sorted(files.values(), key=lambda x: x["name"])
    return {"period": period, "files": ordered}


@app.get("/api/results/{period}/{filename:path}")
async def download_result(period: str, filename: str):
    """下载指定结果文件（filename 仅取文件名，防路径穿越）"""
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}")

    filename = Path(filename).name  # 安全：丢弃任何目录部分
    output_dir = ROOT / "输出"

    # 尝试多种路径匹配
    candidates = [
        output_dir / period / filename,
        output_dir / filename,
    ]
    # 前缀匹配（如 period="2026-07" 匹配 "2026-07 会计报表.xlsx"）。
    # 严格约束：文件名必须以 period 开头，且其后是空格/分隔符，防止 "2026-0" 之类的
    # 前缀把 "2026-071" 这种不相干文件也匹配进来。
    for f in output_dir.iterdir():
        if not (f.is_file() and f.name.startswith(period)):
            continue
        rest = f.name[len(period):]
        if rest and rest[0] not in (" ", "-", "_", "（", "("):
            continue
        if filename in f.name:
            candidates.append(f)
            break

    for filepath in candidates:
        if filepath.exists() and filepath.is_file():
            # 根据类型返回
            ext = filepath.suffix.lower()
            if ext in (".xlsx", ".xls"):
                media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == ".pdf":
                media_type = "application/pdf"
            else:
                media_type = "application/octet-stream"

            display_name = filepath.name if filepath.name.startswith(period) else f"[{period}] {filepath.name}"
            return FileResponse(
                filepath,
                filename=display_name,
                media_type=media_type,
            )

    raise HTTPException(404, f"文件未找到: {period}/{filename}")


@app.post("/api/upload-text")
async def upload_text(
    period: str = Form(...),
    module: str = Form(...),
    text: str = Form(""),
):
    """保存文本补充说明"""
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}")

    target_dir = ROOT / "输入" / period
    target_dir.mkdir(parents=True, exist_ok=True)

    safe_module = Path(module or "").name or "补充"
    filepath = target_dir / f"_{safe_module}_补充说明.txt"
    filepath.write_text(text, encoding="utf-8")

    return {"success": True, "path": str(filepath.relative_to(ROOT))}


@app.delete("/api/input/{period}/{module}/{filename}")
async def delete_input_file(period: str, module: str, filename: str):
    """删除输入目录中的指定文件（module 白名单 + filename 去路径，防越权删除）"""
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}")
    module_dirs = {
        "revenue": "收入发票",
        "expenses": "成本费用发票",
        "bank": "银行对账单",
        "payroll": "工资社保",
        "company-info": "企业信息",
        "tax-adjust": "税务调整",
    }
    if module not in module_dirs:
        raise HTTPException(400, f"不支持的模块: {module}")
    safe = Path(filename).name
    target = ROOT / "输入" / period / module_dirs[module] / safe
    # 安全边界：解析后必须仍位于输入目录内
    try:
        target.resolve().relative_to((ROOT / "输入").resolve())
    except ValueError:
        raise HTTPException(400, "非法路径")
    if not target.exists():
        raise HTTPException(404, f"文件不存在: {target.name}")
    target.unlink()
    return {"success": True, "deleted": safe}


# ---- Helpers ----


# ---- 数据导出 ----


@app.get("/api/export/all")
async def export_all_data():
    """
    一键导出所有数据为 ZIP 文件
    包含：输入目录结构 + 输出目录结构 + 规则文件
    """
    import zipfile
    import io

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 输入目录
        input_root = ROOT / "输入"
        if input_root.exists():
            for f in input_root.rglob("*"):
                if f.is_file():
                    arcname = "输入/" + str(f.relative_to(input_root)).replace("\\", "/")
                    zf.write(f, arcname)

        # 输出目录
        output_root = ROOT / "输出"
        if output_root.exists():
            for f in output_root.rglob("*"):
                if f.is_file() and f.suffix.lower() in (".xlsx", ".xls", ".pdf", ".csv", ".txt"):
                    arcname = "输出/" + str(f.relative_to(output_root)).replace("\\", "/")
                    zf.write(f, arcname)

        # 规则文件
        rules_dir = ROOT / "accounting_report" / "rules"
        if rules_dir.exists():
            for f in rules_dir.iterdir():
                if f.is_file() and f.suffix.lower() in (".yaml", ".yml", ".json"):
                    zf.write(f, f"规则/{f.name}")

        # 配置文件
        for cfg_name in ["config.json", "config.py"]:
            cfg_path = ROOT / "backend" / cfg_name
            if not cfg_path.exists():
                cfg_path = ROOT / "accounting_report" / cfg_name
            if cfg_path.exists():
                zf.write(cfg_path, f"配置/{cfg_name}")

    buf.seek(0)
    safe_name = f"lobster-books-backup-{datetime.date.today().isoformat()}.zip"
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={safe_name}"},
    )


@app.get("/api/export/period/{period}")
async def export_period_data(period: str):
    """
    导出单个账期数据为 ZIP
    """
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}")

    import zipfile
    import io

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 该账期输入
        input_period = ROOT / "输入" / period
        if input_period.exists():
            for f in input_period.rglob("*"):
                if f.is_file():
                    arcname = f"{period}/输入/" + str(f.relative_to(input_period)).replace("\\", "/")
                    zf.write(f, arcname)

        # 该账期输出
        output_period = ROOT / "输出" / period
        if output_period.exists():
            for f in output_period.rglob("*"):
                if f.is_file():
                    arcname = f"{period}/输出/" + str(f.relative_to(output_period)).replace("\\", "/")
                    zf.write(f, arcname)

        # 根级别前缀匹配文件
        output_root = ROOT / "输出"
        if output_root.exists():
            for f in output_root.iterdir():
                if f.is_file() and f.name.startswith(period):
                    zf.write(f, f"{period}/输出/{f.name}")

    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=lobster-books-{period}.zip"},
    )


# ---- Helpers ----


def _get_company_name() -> str:
    # 优先读 Web UI 保存的 config.json；没有才回落到 config.py 的默认值
    cfg_path = ROOT / "backend" / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            name = cfg.get("company_name")
            if name:
                return name
        except Exception:
            pass
    try:
        from accounting_report.config import COMPANY_NAME
        return COMPANY_NAME
    except Exception:
        return "未设置"


def _get_config_entity_type() -> str:
    """从 Web UI 保存的 config.json 读取 entity_type"""
    cfg_path = ROOT / "backend" / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return cfg.get("entity_type", "corp")
        except Exception:
            pass
    return "corp"


def _validate_period(period: str) -> bool:
    if len(period) < 7:
        return False
    try:
        y, m = int(period[:4]), int(period[5:7])
        return 2020 <= y <= 2099 and 1 <= m <= 12 and period[4] == "-"
    except (ValueError, IndexError):
        return False


def _build_output_list(result) -> list[dict]:
    """从 FullResult 构建输出列表"""
    outputs = []
    output_map = [
        ("会计报表", getattr(result, "output3_report", None)),
        ("银行对账报告", getattr(result, "output1_reconciliation", None)),
        ("核算汇总表", getattr(result, "output2_summary", None)),
        ("分析简报", getattr(result, "output4_analysis", None)),
        ("增值税申报底稿", getattr(result, "output5_vat", None)),
        ("所得税预缴底稿", getattr(result, "output6_income_tax", None)),
        ("个税扣缴明细表", getattr(result, "output7_payroll_tax", None)),
        ("记账凭证", getattr(result, "output8_vouchers", None)),
        ("试算平衡表", getattr(result, "output9_trial_balance", None)),
    ]

    for label, path in output_map:
        if path and Path(path).exists():
            p = Path(path)
            # 从文件名前 7 位提取账期 YYYY-MM；不满足格式时回退到输出目录名，
            # 避免 `p.name[:7]` 对非 YYYY-MM 开头的文件名给出错误账期。
            period_part = p.name[:7]
            if not _validate_period(period_part):
                period_part = p.parent.name if _validate_period(p.parent.name) else period_part
            outputs.append({
                "label": label,
                "filename": p.name,
                "size": p.stat().st_size,
                "download_url": f"/api/results/{period_part}/{p.name}",
            })

    return outputs


def _build_stats(result) -> dict:
    """构建统计信息"""
    stats = {
        "voucher_count": 0,
        "bank_txn_count": 0,
        "journal_entry_count": 0,
    }
    if result.query:
        if hasattr(result.query, "voucher_entries"):
            stats["voucher_count"] = len(result.query.voucher_entries)
        if hasattr(result.query, "bank_txns"):
            stats["bank_txn_count"] = len(result.query.bank_txns)
        if hasattr(result.query, "journal_entries"):
            stats["journal_entry_count"] = len(result.query.journal_entries)
    return stats


# ---- 规则与配置 ----


# ================================================================
# Rules & Config API
# ================================================================


@app.get("/api/rules/mapping")
async def get_mapping_rules():
    """获取银行摘要 → 科目映射规则"""
    try:
        import yaml
        mapping_path = ROOT / "accounting_report" / "rules" / "mapping.yaml"
        if not mapping_path.exists():
            raise HTTPException(404, "mapping.yaml 未找到")
        with open(mapping_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        rules = data.get("bank_rules", [])
        return {
            "rules": rules,
            "cash_account": data.get("cash_account", "货币资金"),
            "source_file": str(mapping_path.relative_to(ROOT)),
        }
    except Exception as e:
        raise HTTPException(500, f"读取映射规则失败: {e}")


@app.put("/api/rules/mapping")
async def update_mapping_rules(body: dict):
    """更新银行摘要 → 科目映射规则（整表替换）"""
    import yaml
    mapping_path = ROOT / "accounting_report" / "rules" / "mapping.yaml"
    if not mapping_path.exists():
        raise HTTPException(404, "mapping.yaml 未找到")

    rules = body.get("bank_rules", body.get("rules", []))
    cash_account = body.get("cash_account", "货币资金")

    # 重新构建 YAML
    output_lines = [
        "# 银行流水摘要 → 对方科目",
        "#",
        "# 匹配规则：按 match 关键词在「摘要 + 对方户名」里做子串查找，自上而下first-match。",
        "# 匹配不上的流水不会被猜测入账——程序会列出待归类清单并报错。",
        "",
        "bank_rules:",
    ]
    for r in rules:
        output_lines.append("")
        match_val = r.get("match", "")
        account_val = r.get("account", "")
        # 安全写入 YAML：对所有值加双引号，避免冒号/井号/引号破坏语法
        match_safe = match_val.replace("\\", "\\\\").replace('"', '\\"')
        account_safe = account_val.replace("\\", "\\\\").replace('"', '\\"')
        output_lines.append(f'  - match: "{match_safe}"')
        output_lines.append(f'    account: "{account_safe}"')
        note = r.get("note", "")
        if note:
            note_safe = note.replace("\\", "\\\\").replace('"', '\\"')
            output_lines.append(f'    note: "{note_safe}"')

    output_lines.extend([
        "",
        "# 现金及现金等价物科目（对账单的借贷方向据此判断流入/流出）",
        f"cash_account: {cash_account}",
        "",
    ])

    with open(mapping_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

    return {"success": True, "rules_count": len(rules)}


@app.get("/api/rules/accounts")
async def get_accounts():
    """获取科目表"""
    try:
        import yaml
        accounts_path = ROOT / "accounting_report" / "rules" / "accounts.yaml"
        if not accounts_path.exists():
            raise HTTPException(404, "accounts.yaml 未找到")
        with open(accounts_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        return {
            "aliases": data.get("aliases", {}),
            "accounts": data.get("accounts", {}),
        }
    except Exception as e:
        raise HTTPException(500, f"读取科目表失败: {e}")


@app.get("/api/config")
async def get_config():
    """获取企业配置"""
    config_path = ROOT / "backend" / "config.json"
    if config_path.exists():
        import json
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)

    from accounting_report.config import COMPANY_NAME
    return {
        "entity_type": "corp",
        "company_name": COMPANY_NAME,
        "tax_rate": 5.0,
    }


@app.post("/api/config")
async def save_config(body: dict):
    """保存企业配置。

    只写入 backend/config.json（JSON 序列化天然处理引号/换行转义）。
    不再重写 accounting_report/config.py 源码——此前用 re.sub 把 company_name
    原样拼进源码，企业名含 `"` 或换行会破坏语法，甚至构造可执行 Python
    （结合历史 host=0.0.0.0 即 RCE）。config.py 的 COMPANY_NAME 仅作默认值，
    读取配置一律优先走 config.json。
    """
    config_path = ROOT / "backend" / "config.json"
    config_path.parent.mkdir(parents=True, exist_ok=True)
    import json
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False, indent=2)

    return {"success": True}


# ---- 历史账期归档 ----

ARCHIVE_DIR = ROOT / "归档"
MIN_OUTPUT_FILES = 4  # 至少需要 4 个 xlsx 输出文件才算完整


class ArchiveRequest(BaseModel):
    period: str


class ArchiveResponse(BaseModel):
    success: bool
    message: str
    period: str
    archive_path: str = ""


@app.post("/api/archive", response_model=ArchiveResponse)
async def archive_period(req: ArchiveRequest):
    """将指定账期的输入/输出文件移动到归档目录"""
    period = req.period
    if not _validate_period(period):
        raise HTTPException(400, f"账期格式错误: {period}，应为 YYYY-MM")

    input_dir = ROOT / "输入" / period
    output_dir = ROOT / "输出" / period

    # 检查输出是否完整：至少 MIN_OUTPUT_FILES 个 xlsx 文件
    output_xlsx_files: list[Path] = []
    if output_dir.is_dir():
        output_xlsx_files = [f for f in output_dir.iterdir() if f.is_file() and f.suffix.lower() == ".xlsx"]

    # 也查找根级别前缀匹配的文件
    root_output = ROOT / "输出"
    if root_output.is_dir():
        for f in root_output.iterdir():
            if f.is_file() and f.suffix.lower() == ".xlsx" and f.name.startswith(period):
                output_xlsx_files.append(f)

    if len(output_xlsx_files) < MIN_OUTPUT_FILES:
        return ArchiveResponse(
            success=False,
            message=f"输出不完整：仅找到 {len(output_xlsx_files)} 个 xlsx 文件（至少需要 {MIN_OUTPUT_FILES} 个），请先生成完整的账务文件后再归档。",
            period=period,
        )

    # 创建归档目标目录
    archive_period_dir = ARCHIVE_DIR / period
    archive_output_dir = archive_period_dir / "输出"
    archive_output_dir.mkdir(parents=True, exist_ok=True)

    moved_count = 0

    # 移动输入目录
    if input_dir.exists():
        dest_input = archive_period_dir
        for item in input_dir.iterdir():
            dest = dest_input / item.name
            if dest.exists():
                if dest.is_dir():
                    shutil.rmtree(dest)
                else:
                    dest.unlink()
            shutil.move(str(item), str(dest))
            moved_count += 1
        # 删除空的输入账期目录
        try:
            input_dir.rmdir()
        except OSError:
            pass

    # 移动输出目录中的文件
    if output_dir.is_dir():
        for item in output_dir.iterdir():
            dest = archive_output_dir / item.name
            if dest.exists():
                if dest.is_dir():
                    shutil.rmtree(dest)
                else:
                    dest.unlink()
            shutil.move(str(item), str(dest))
            moved_count += 1
        # 删除空的输出账期目录
        try:
            output_dir.rmdir()
        except OSError:
            pass

    # 移动根级别前缀匹配文件
    if root_output.is_dir():
        for f in list(root_output.iterdir()):
            if f.is_file() and f.name.startswith(period):
                dest = archive_output_dir / f.name
                if dest.exists():
                    dest.unlink()
                shutil.move(str(f), str(dest))
                moved_count += 1

    return ArchiveResponse(
        success=True,
        message=f"账期 {period} 归档完成，共移动 {moved_count} 项。",
        period=period,
        archive_path=str(archive_period_dir.relative_to(ROOT)),
    )


@app.get("/api/archive")
async def list_archived():
    """列出已归档的账期列表"""
    if not ARCHIVE_DIR.exists():
        return {"archived": []}

    archived = []
    for d in sorted(ARCHIVE_DIR.iterdir(), reverse=True):
        if d.is_dir() and _validate_period(d.name):
            # 统计输出文件数
            out_dir = d / "输出"
            out_count = 0
            if out_dir.is_dir():
                out_count = len([f for f in out_dir.iterdir() if f.is_file()])
            archived.append({
                "period": d.name,
                "output_count": out_count,
                "archived_at": datetime.datetime.fromtimestamp(d.stat().st_mtime).isoformat(),
            })

    return {"archived": archived}


# ---- Main ----

if __name__ == "__main__":
    # 桌面版：Electron 通过 XJZ_PORT 传入随机可用端口
    # 开发/服务器版：默认 8100，生产不启用 reload（避免 stat reloader 重启）
    _port = int(os.environ.get("XJZ_PORT", "8100"))
    _host = os.environ.get("XJZ_HOST", "127.0.0.1")
    _reload = os.environ.get("MAR_RELOAD") == "1" and not getattr(sys, "frozen", False)

    uvicorn.run(
        "server:app" if _reload else app,
        # 仅监听本机回环地址：nginx 已在本机反代 127.0.0.1:8100 对外提供服务，
        # 避免局域网内任意设备直连 8100 访问无鉴权的 API（下载/删改/配置等）。
        host=_host,
        port=_port,
        reload=_reload,
        log_level=os.environ.get("XJZ_LOG_LEVEL", "info"),
    )
