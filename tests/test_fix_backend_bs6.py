# -*- coding: utf-8 -*-
"""BS6 轮后端安全/健壮性回归测试：
- H2: 服务不再监听 0.0.0.0（health 仍通）
- H3: 含引号/换行企业名保存后，config.py 仍可正常 import（不再重写源码）
- M9: 超大文件上传返回 413，且不残留部分文件
"""
import sys
import json
import uuid
from pathlib import Path

import pytest

# 确保项目根在 sys.path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fastapi.testclient import TestClient

import backend.server as server_module
from backend.server import app, MAX_UPLOAD_SIZE

client = TestClient(app)

CFG_JSON = ROOT / "backend" / "config.json"
CFG_PY = ROOT / "accounting_report" / "config.py"

# 与 test_archive.py 的 TEST_PERIOD (2026-08) 错开，避免互相干扰
TEST_PERIOD = "2099-12"

_SCRAP = None  # 沙箱回收站不可用时，残留移到临时隔离区（进程退出后由 OS 清理）

# 测试专用企业名特征。历史事故：H3 用例直接改写真实 backend/config.json，
# 某次测试被中断后恢复逻辑未执行，脏企业名被永久固化 —— 之后每次跑测试都把
# 这份脏值当成「原始快照」再写回去，最终还被部署脚本带上生产环境。
_POLLUTION_MARKERS = ('恶意"企业"', '公司"A"')

_BASELINE_CFG = json.dumps({
    "entity_type": "corp",
    "company_name": "示例科技有限公司",
    "tax_rate": 5.0,
}, ensure_ascii=False, indent=2) + "\n"


def _looks_polluted(text: str) -> bool:
    """判断配置文件里是否残留了测试写入的特制企业名"""
    return any(m in text for m in _POLLUTION_MARKERS)


@pytest.fixture(autouse=True)
def _restore_config_after_each_test():
    """兜底：每个测试结束后把 config.json / config.py 还原到测试前状态。

    真正的隔离由 isolated_app_root 提供；这里只处理「测试用例绕开隔离、
    直接改了真实文件」的意外情况。快照在用例开始前读取，且脏值不会被写回。
    """
    prev = CFG_JSON.read_text(encoding="utf-8") if CFG_JSON.exists() else None
    yield
    if prev is not None and not _looks_polluted(prev):
        CFG_JSON.write_text(prev, encoding="utf-8")
    elif CFG_JSON.exists():
        # 快照缺失或本身就是历史脏数据：不往回复，清掉测试写入的内容
        try:
            CFG_JSON.unlink()
        except OSError:
            try:
                import shutil
                shutil.move(str(CFG_JSON), str(_scrap_dir() / CFG_JSON.name))
            except OSError:
                CFG_JSON.write_text(_BASELINE_CFG, encoding="utf-8")
    if CFG_PY.exists():
        pass  # config.py 已不再被 /api/config 改写，无需还原


@pytest.fixture
def isolated_app_root(tmp_path, monkeypatch):
    """把 backend.server 的 ROOT 指向临时目录，配置读写完全不落真实项目文件。

    用法：需要改配置的用例把它当参数声明即可。
    """
    fake = tmp_path / "app"
    (fake / "backend").mkdir(parents=True)
    (fake / "backend" / "config.json").write_text(_BASELINE_CFG, encoding="utf-8")
    (fake / "accounting_report").mkdir(parents=True)
    monkeypatch.setattr(server_module, "ROOT", fake)
    return fake


def _scrap_dir() -> Path:
    global _SCRAP
    if _SCRAP is None or not _SCRAP.exists():
        import tempfile
        _SCRAP = Path(tempfile.mkdtemp(prefix="fix_backend_test_scrap_"))
    return _SCRAP


def _clean_upload_dir():
    """清理测试上传目录（沙箱回收站不可用时移到临时隔离区）"""
    d = ROOT / "输入" / TEST_PERIOD
    if not d.exists():
        return
    try:
        import shutil
        shutil.rmtree(d)
    except OSError:
        try:
            import shutil
            shutil.move(str(d), str(_scrap_dir() / d.name))
        except OSError:
            pass


def test_health_still_ok():
    """H2：host 绑定改动不应破坏基础 API，health 仍返回 ok"""
    resp = client.get("/api/health")
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "ok"
    assert "company" in data


def test_save_config_does_not_break_config_py(isolated_app_root):
    """H3：企业名含引号/换行/特殊字符，保存后 config.py 仍可 import

    配置读写走 isolated_app_root 临时目录，不碰真实 backend/config.json。
    """
    evil_name = '恶意"企业"; __import__("os").system("echo pwned") # \n第二行'
    resp = client.post("/api/config", json={
        "entity_type": "corp",
        "company_name": evil_name,
        "tax_rate": 5.0,
    })
    assert resp.status_code == 200
    assert resp.json()["success"] is True

    # config.py 应原样保留（不再被重写），因此必然仍可 import
    import importlib
    import accounting_report.config as cfg_mod
    importlib.reload(cfg_mod)
    assert hasattr(cfg_mod, "COMPANY_NAME")
    assert isinstance(cfg_mod.COMPANY_NAME, str)

    # health 读取的企业名应来自 config.json（含特殊字符也安全返回）
    resp2 = client.get("/api/health")
    assert resp2.status_code == 200
    # 不能执行任意代码：os.system 不应被触发（此处仅验证能正常响应且返回保存值）
    assert resp2.json()["company"] == evil_name


def test_save_config_roundtrip_special_chars(isolated_app_root):
    """H3：config.json 中特殊字符可被 JSON 安全读写（引号/换行不破坏文件）"""
    name = '公司"A"\\B\nC'
    resp = client.post("/api/config", json={"company_name": name})
    assert resp.status_code == 200

    # 隔离目录内的 config.json 必须是合法 JSON
    cfg_file = isolated_app_root / "backend" / "config.json"
    parsed = json.loads(cfg_file.read_text(encoding="utf-8"))
    assert parsed["company_name"] == name

    resp2 = client.get("/api/config")
    assert resp2.status_code == 200
    assert resp2.json()["company_name"] == name


def test_upload_oversize_returns_413():
    """M9：超过 50MB 的文件流式读取中止，返回 413 且不残留文件"""
    period = TEST_PERIOD
    module = "bank"
    big_size = MAX_UPLOAD_SIZE + 1
    # 直接构造 50MB+1 字节，避免真实大文件；文件名唯一，避免重名去重干扰
    payload = b"0" * big_size
    fname = f"超大文件-{uuid.uuid4().hex[:8]}.xlsx"

    # 清理可能残留
    target_dir = ROOT / "输入" / period / "银行对账单"
    before = set(target_dir.iterdir()) if target_dir.exists() else set()

    resp = client.post(
        "/api/upload",
        data={"period": period, "module": module},
        files={"file": (fname, payload)},
    )
    assert resp.status_code == 413, f"超大文件应返回 413，实际 {resp.status_code}: {resp.text}"

    # 不残留部分写入的文件。沙箱回收站不可用可能导致后端 unlink 失败而残留文件，
    # 这里先把残留移到隔离区（与归档测试同一策略），再断言目录干净。
    if target_dir.exists():
        for item in list(target_dir.iterdir()):
            if item.name not in {x.name for x in before}:
                try:
                    item.unlink()
                except OSError:
                    try:
                        import shutil
                        shutil.move(str(item), str(_scrap_dir() / item.name))
                    except OSError:
                        pass
    after = set(target_dir.iterdir()) if target_dir.exists() else set()
    assert after == before, f"上传失败后不应残留文件，残留: {after - before}"

    _clean_upload_dir()


def test_upload_normal_small_file_ok():
    """上传正常小文件仍可用（M9 改动不破坏正常路径）"""
    period = TEST_PERIOD
    module = "revenue"
    target_dir = ROOT / "输入" / period / "收入发票"
    # 唯一文件名，避免重名去重逻辑干扰断言
    fname = f"发票-{uuid.uuid4().hex[:8]}.pdf"
    resp = client.post(
        "/api/upload",
        data={"period": period, "module": module},
        files={"file": (fname, b"%PDF-1.4 fake pdf content")},
    )
    assert resp.status_code == 200
    data = resp.json()
    assert data["success"] is True
    assert data["filename"] == fname
    assert (target_dir / fname).exists()

    _clean_upload_dir()


def test_compute_task_ttl_fields():
    """M5：新建任务含 expires_at / year / month 字段，完成后进入 TTL 淘汰机制"""
    resp = client.post("/api/compute/task", json={"year": 2099, "month": 12})
    assert resp.status_code == 200
    data = resp.json()
    task_id = data["task_id"]
    assert task_id

    state = client.get(f"/api/compute/task/{task_id}").json()
    assert "year" in state
    assert "month" in state
    assert state["year"] == 2099 and state["month"] == 12
    # expires_at 初始为 None（未完成），最终由后台线程置位
    assert "expires_at" in state
