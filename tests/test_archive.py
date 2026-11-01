# -*- coding: utf-8 -*-
"""测试历史账期归档 API 逻辑"""
import sys
import shutil
from pathlib import Path

# 确保项目根在 sys.path
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl
from fastapi.testclient import TestClient

# 导入 app
from backend.server import app

client = TestClient(app)

TEST_PERIOD = "2026-08"


_SCRAP = None  # 沙箱回收站不可用时，残留移到临时隔离区（进程退出后由 OS 清理）


def _scrap_dir() -> Path:
    global _SCRAP
    if _SCRAP is None or not _SCRAP.exists():
        import tempfile
        _SCRAP = Path(tempfile.mkdtemp(prefix="archive_test_scrap_"))
    return _SCRAP


def _clean_all():
    """彻底清理所有测试残留。

    优先删除；沙箱回收站不可用时改「移到临时隔离区」，
    保证归档目录/输入/输出残留不污染后续测试的断言。
    """
    import shutil as _sh
    for d in [
        ROOT / "输入" / TEST_PERIOD,
        ROOT / "输出" / TEST_PERIOD,
        ROOT / "归档" / TEST_PERIOD,
        ROOT / "归档",
    ]:
        if not d.exists():
            continue
        try:
            _sh.rmtree(d)
        except OSError:
            try:
                _sh.move(str(d), str(_scrap_dir() / d.name))
            except OSError:
                pass
    # 清理根级别前缀匹配文件
    root_out = ROOT / "输出"
    if root_out.is_dir():
        for f in list(root_out.iterdir()):
            if f.is_file() and f.name.startswith(TEST_PERIOD):
                try:
                    f.unlink()
                except OSError:
                    pass


def setup_test_files():
    """创建测试用的输入和输出文件 (6 xlsx)"""
    _clean_all()

    test_input = ROOT / "输入" / TEST_PERIOD
    test_output = ROOT / "输出" / TEST_PERIOD

    # 创建输入文件
    (test_input / "银行对账单").mkdir(parents=True, exist_ok=True)
    (test_input / "银行对账单" / "对账单.txt").write_text("test bank data", encoding="utf-8")
    (test_input / "收入发票").mkdir(parents=True, exist_ok=True)
    (test_input / "收入发票" / "发票.pdf").write_bytes(b"fake pdf")

    # 创建输出文件（6个 xlsx，大于 4 个最小值）
    test_output.mkdir(parents=True, exist_ok=True)
    for fname in [
        "会计报表.xlsx",
        "银行对账报告.xlsx",
        "核算汇总表.xlsx",
        "分析简报.xlsx",
        "增值税申报底稿.xlsx",
        "所得税预缴底稿.xlsx",
    ]:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = f"Test data for {fname}"
        wb.save(str(test_output / fname))

    # 根级别前缀匹配文件
    root_out = ROOT / "输出"
    (root_out / f"{TEST_PERIOD} 明细账.xlsx").write_bytes(b"root level file")


def setup_incomplete_files():
    """创建不完整的输出文件 (仅 2 xlsx)"""
    _clean_all()

    test_output = ROOT / "输出" / TEST_PERIOD
    test_output.mkdir(parents=True, exist_ok=True)

    for fname in ["会计报表.xlsx", "银行对账报告.xlsx"]:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = f"Test data {fname}"
        wb.save(str(test_output / fname))


def test_archive_complete_output():
    """测试：输出完整时归档成功"""
    setup_test_files()

    # 归档前验证
    assert (ROOT / "输入" / TEST_PERIOD).exists()
    assert (ROOT / "输出" / TEST_PERIOD).exists()
    assert len(list((ROOT / "输出" / TEST_PERIOD).glob("*.xlsx"))) >= 4

    # 调用归档 API
    resp = client.post("/api/archive", json={"period": TEST_PERIOD})
    assert resp.status_code == 200
    data = resp.json()
    assert data["success"] is True, f"归档应成功: {data}"
    assert data["period"] == TEST_PERIOD
    assert "归档完成" in data["message"]

    # 验证文件已移动到归档目录
    archive_input = ROOT / "归档" / TEST_PERIOD
    archive_output = ROOT / "归档" / TEST_PERIOD / "输出"
    assert archive_input.exists(), "归档输入目录应存在"
    assert archive_output.exists(), "归档输出目录应存在"
    assert (archive_input / "银行对账单" / "对账单.txt").exists()
    assert (archive_input / "收入发票" / "发票.pdf").exists()
    assert (archive_output / "会计报表.xlsx").exists()
    assert (archive_output / "核算汇总表.xlsx").exists()
    assert (archive_output / f"{TEST_PERIOD} 明细账.xlsx").exists()

    # 验证原目录文件已移空（沙箱回收站不可用时会拦截 rmdir，允许目录壳残留，
    # 但文件必须全部移走——归档的实质是「文件搬走」而非「目录删除」）
    for label, orig in [
        ("输入", ROOT / "输入" / TEST_PERIOD),
        ("输出", ROOT / "输出" / TEST_PERIOD),
    ]:
        if orig.exists():
            remaining = [str(r) for r in orig.rglob("*")]
            assert remaining == [], f"归档后{label}目录应已移空，残留: {remaining}"

    print("[OK] test_archive_complete_output")

    # 测试 list_archived
    resp2 = client.get("/api/archive")
    assert resp2.status_code == 200
    data2 = resp2.json()
    periods = [a["period"] for a in data2["archived"]]
    assert TEST_PERIOD in periods, f"已归档列表应包含 {TEST_PERIOD}"
    print("[OK] list_archived contains period")

    _clean_all()


def test_archive_incomplete_output():
    """测试：输出文件不足时拒绝归档"""
    setup_incomplete_files()

    resp = client.post("/api/archive", json={"period": TEST_PERIOD})
    assert resp.status_code == 200
    data = resp.json()
    assert data["success"] is False, f"不完整输出应拒绝归档: {data}"
    assert "不完整" in data["message"]

    print("[OK] test_archive_incomplete_output")

    _clean_all()


def test_archive_list_empty():
    """测试：无归档时返回空列表"""
    _clean_all()

    resp = client.get("/api/archive")
    assert resp.status_code == 200
    data = resp.json()
    assert data["archived"] == [], f"无归档时应返回空列表: {data}"

    print("[OK] test_archive_list_empty")


def test_archive_invalid_period():
    """测试：无效账期格式"""
    resp = client.post("/api/archive", json={"period": "invalid"})
    assert resp.status_code == 400

    print("[OK] test_archive_invalid_period")


if __name__ == "__main__":
    print("=" * 60)
    print("Archive API Tests")
    print("=" * 60)

    try:
        test_archive_list_empty()
        test_archive_invalid_period()
        test_archive_incomplete_output()
        test_archive_complete_output()
        print()
        print("=" * 60)
        print("ALL TESTS PASSED!")
        print("=" * 60)
    except Exception as e:
        print(f"\n[FAIL] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        _clean_all()
