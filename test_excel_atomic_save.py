"""Test: 保存 Excel 不能把原文件写坏，图片必须在 add 时就读进内存。

背景（2026-08-06）：员工的汇总表跑完后损坏，文件里躺着 40 张图但没有
[Content_Types].xml。原因是 openpyxl 按路径持有图片、到 wb.save() 才回磁盘
重读，某张读失败就在流式写 zip 的中途抛异常 —— 而 wb.save() 是直接往目标
文件上写的，原内容已经没了。

Run: python test_excel_atomic_save.py
"""
import shutil
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

from shein_scraper import save_workbook_atomic, _add_picture_to_cell


def _good_workbook(path: Path, marker: str = "原始数据"):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = marker
    for r in range(2, 200):
        ws.cell(r, 1, f"row{r}")
    wb.save(path)
    return path


def _make_webp(path: Path, size=(120, 160), color=(200, 30, 30)):
    PILImage.new("RGB", size, color).save(path, format="WEBP")
    return path


def _is_openable(path: Path) -> bool:
    try:
        load_workbook(path)
        return True
    except Exception:
        return False


# ── save_workbook_atomic ────────────────────────────────────────────────────

def test_atomic_save_writes_new_content_on_success():
    tmp = Path(tempfile.mkdtemp())
    try:
        target = tmp / "表.xlsx"
        _good_workbook(target, "旧内容")
        wb = Workbook()
        wb.active["A1"] = "新内容"
        save_workbook_atomic(wb, target)
        assert load_workbook(target).active["A1"].value == "新内容"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_atomic_save_leaves_original_intact_when_save_raises():
    """核心保证：保存中途炸了，原文件必须一个字节都没变。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        target = tmp / "重要总表.xlsx"
        _good_workbook(target, "员工积累了几个月的数据")
        before = target.read_bytes()

        wb = Workbook()
        wb.active["A1"] = "本轮的新内容"

        # openpyxl 内部炸掉：Workbook.save 抛异常 → save_workbook_atomic
        # 在校验前就失败 → 原表一个字节都没动过（core atomic 保证）。
        with patch("openpyxl.workbook.workbook.Workbook.save",
                   side_effect=RuntimeError("simulated openpyxl mid-write failure")):
            try:
                save_workbook_atomic(wb, target)
                raised = False
            except Exception:
                raised = True

        assert raised, "保存该失败却成功了，测试前提不成立"
        assert target.read_bytes() == before, "原文件被改动了"
        assert _is_openable(target), "原文件损坏了"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_broken_image_is_dropped_instead_of_killing_save():
    """图片读不出来时应该只丢掉那一张，其他内容照常保存 —— 不能因为一张
    烂图把整轮抓取的成果拖下水（v0.3.10/v0.3.11 前的行为）。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        target = tmp / "带烂图的表.xlsx"
        img = _make_webp(tmp / "victim.webp")

        wb = Workbook()
        wb.active["A1"] = "新内容"
        wb.active.add_image(XLImage(str(img)), "H2")
        img.unlink()   # save 时 _data() 会读不到

        save_workbook_atomic(wb, target)   # 不该抛
        assert load_workbook(target).active["A1"].value == "新内容"
        # 那张烂图被 _harden_workbook_images 丢掉了
        z = zipfile.ZipFile(target)
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert media == [], f"烂图应被丢弃却留在了 xlsx 里: {media}"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_atomic_save_leaves_no_temp_files_behind():
    tmp = Path(tempfile.mkdtemp())
    try:
        target = tmp / "表.xlsx"
        _good_workbook(target)
        img = _make_webp(tmp / "victim.webp")
        wb = Workbook()
        wb.active.add_image(XLImage(str(img)), "H2")
        img.unlink()
        try:
            save_workbook_atomic(wb, target)
        except Exception:
            pass
        leftovers = [p.name for p in tmp.iterdir()
                     if p.name != target.name and p.suffix != ".webp"]
        assert leftovers == [], f"留下了临时文件: {leftovers}"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_atomic_save_creates_file_that_did_not_exist():
    tmp = Path(tempfile.mkdtemp())
    try:
        target = tmp / "全新.xlsx"
        wb = Workbook()
        wb.active["A1"] = "x"
        save_workbook_atomic(wb, target)
        assert target.is_file() and _is_openable(target)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ── _add_picture_to_cell 必须在 add 时就把字节读进内存 ────────────────────────

def test_picture_bytes_are_read_eagerly_not_at_save_time():
    """add 完之后把源文件删掉，保存仍然要成功 —— 证明字节已经在内存里。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        img = _make_webp(tmp / "img_001.webp")
        wb = Workbook()
        ws = wb.active
        h = _add_picture_to_cell(ws, 2, 8, img)
        assert h > 0, "图片没嵌进去"
        img.unlink()                       # 源文件消失

        target = tmp / "out.xlsx"
        save_workbook_atomic(wb, target)   # 不该抛
        z = zipfile.ZipFile(target)
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert len(media) == 1, media
        assert "[Content_Types].xml" in z.namelist()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_unreadable_image_is_skipped_not_raised():
    """坏图只该少一张图，不该把整次保存拖下水。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        bad = tmp / "img_001.webp"
        bad.write_bytes("这不是图片".encode("utf-8"))
        wb = Workbook()
        ws = wb.active
        assert _add_picture_to_cell(ws, 2, 8, bad) == 0
        assert len(ws._images) == 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_missing_image_file_is_skipped():
    wb = Workbook()
    ws = wb.active
    assert _add_picture_to_cell(ws, 2, 8, Path("__no_such_image__.webp")) == 0
    assert len(ws._images) == 0


def test_webp_is_converted_so_excel_can_show_it():
    """Excel 不认 webp。嵌进去的必须是 Excel 支持的格式（现在是 jpg）。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        img = _make_webp(tmp / "img_001.webp")
        wb = Workbook()
        _add_picture_to_cell(wb.active, 2, 8, img)
        target = tmp / "out.xlsx"
        save_workbook_atomic(wb, target)
        media = [n for n in zipfile.ZipFile(target).namelist()
                 if n.startswith("xl/media/")]
        assert media and all(
            n.lower().endswith((".jpg", ".jpeg", ".png", ".gif")) for n in media
        ), media
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_embedded_image_is_compressed_not_full_resolution():
    """1050×1050 原图不该原样嵌进去 —— 单张 embed 必须小于源文件。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        big = tmp / "img_big.webp"
        # Shein 常见的 1050×1050 商品图，压 webp 后本身就有一定体积
        PILImage.new("RGB", (1050, 1050), (30, 120, 200)).save(
            big, format="WEBP", quality=95)
        wb = Workbook()
        _add_picture_to_cell(wb.active, 2, 8, big)
        target = tmp / "out.xlsx"
        save_workbook_atomic(wb, target)
        z = zipfile.ZipFile(target)
        media_infos = [i for i in z.infolist() if i.filename.startswith("xl/media/")]
        assert len(media_infos) == 1, [i.filename for i in media_infos]
        embed = media_infos[0]
        # 目标是每张缩到 ~200×200 附近，JPEG q=80，一张远小于 50 KB
        assert embed.file_size < 50_000, (
            f"embedded image {embed.filename} = {embed.file_size} bytes "
            f"— compression didn't kick in"
        )
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ── 2026-09-23: local-tmp + copy-to-destination architecture ────────────────

def test_workbook_is_verified_before_anything_touches_the_destination():
    """openpyxl 绝不能直接往目标目录里流式写。

    目标一般是 Google Drive 的虚拟盘：刚写完的大文件立刻读回来经常还没
    落地，`zipfile.ZipFile()` 直接 BadZipFile —— 文件是好的，是「写完
    马上读回」这一眼不可靠（员工的 v0.3.10/v0.3.11 现场，美国站 2026-08-17
    也踩过同一个坑）。所以自检必须在本地盘上做完，目标盘只接一次已经
    验证过的整文件拷贝。
    """
    import shein_scraper
    tmp = Path(tempfile.mkdtemp())
    try:
        dest_dir = tmp / "云盘"
        dest_dir.mkdir()
        target = dest_dir / "总表.xlsx"

        written_to = []
        orig = shein_scraper._write_workbook_to_temp

        def spy(wb, path):
            written_to.append(Path(path))
            return orig(wb, path)

        shein_scraper._write_workbook_to_temp = spy
        try:
            wb = Workbook()
            wb.active["A1"] = "x"
            save_workbook_atomic(wb, target)
        finally:
            shein_scraper._write_workbook_to_temp = orig

        assert written_to, "根本没调用 _write_workbook_to_temp"
        for p in written_to:
            assert dest_dir not in p.parents, (
                f"openpyxl 直接写进了目标目录: {p}")
        assert _is_openable(target)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_save_survives_a_destination_that_lies_on_read_back():
    """目标盘刚写完的文件读回来是坏的 —— 保存仍然必须成功。

    员工的 v0.3.11 现场：`wb.save(tmp)` 成功，紧接着 `zipfile.ZipFile(tmp)`
    抛 BadZipFile。文件其实是好的，是那一眼读回来不可靠。自检搬到本地盘
    之后这条路就不该再炸。
    """
    import shein_scraper
    tmp = Path(tempfile.mkdtemp())
    try:
        dest_dir = tmp / "云盘"
        dest_dir.mkdir()
        target = dest_dir / "总表.xlsx"

        real_zipfile = shein_scraper.zipfile.ZipFile

        class LyingZipFile(real_zipfile):
            def __init__(self, file, *a, **kw):
                if dest_dir in Path(str(file)).parents:
                    raise zipfile.BadZipFile("File is not a zip file")
                super().__init__(file, *a, **kw)

        shein_scraper.zipfile.ZipFile = LyingZipFile
        try:
            img = _make_webp(tmp / "img_001.webp")
            wb = Workbook()
            wb.active["A1"] = "抓了一整轮的结果"
            _add_picture_to_cell(wb.active, 2, 8, img)
            save_workbook_atomic(wb, target)
        finally:
            shein_scraper.zipfile.ZipFile = real_zipfile

        assert _is_openable(target)
        assert load_workbook(target).active["A1"].value == "抓了一整轮的结果"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_workbook_with_images_can_be_saved_twice():
    """备份优先 + 回写原表的流程会对同一个 wb 存两次。第二次不能因为图片
    被 close 掉而炸（v0.3.10/v0.3.11 员工现场）。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        img = _make_webp(tmp / "img_001.webp")
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Test1"
        ws2 = wb.create_sheet("Test2")
        _add_picture_to_cell(ws1, 2, 8, img)
        target = tmp / "总表.xlsx"

        save_workbook_atomic(wb, target)          # 第一次
        _add_picture_to_cell(ws2, 2, 8, img)
        save_workbook_atomic(wb, target)          # 第二次 —— 以前在这里炸
        save_workbook_atomic(wb, target)          # 第三次也要活着

        z = zipfile.ZipFile(target)
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert len(media) == 2, media
        assert "[Content_Types].xml" in z.namelist()
        assert _is_openable(target)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_reloaded_workbook_images_survive_repeated_saves():
    """load_workbook 读出来的图也是 BytesIO ref —— 同样要能反复存。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        img = _make_webp(tmp / "img_001.webp")
        wb = Workbook()
        _add_picture_to_cell(wb.active, 2, 8, img)
        target = tmp / "已有图的总表.xlsx"
        save_workbook_atomic(wb, target)

        wb2 = load_workbook(target)
        assert len(wb2.active._images) == 1
        save_workbook_atomic(wb2, target)
        save_workbook_atomic(wb2, target)
        media = [n for n in zipfile.ZipFile(target).namelist()
                 if n.startswith("xl/media/")]
        assert len(media) == 1, media
        assert _is_openable(target)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_many_images_all_land():
    """60 张图全部落盘，且清单完整 —— 就是出事的那个规模。"""
    tmp = Path(tempfile.mkdtemp())
    try:
        wb = Workbook()
        ws = wb.active
        for i in range(1, 61):
            p = _make_webp(tmp / f"img_{i:03d}.webp", color=(i * 3 % 256, 40, 90))
            assert _add_picture_to_cell(ws, i + 1, 8, p) > 0
        target = tmp / "out.xlsx"
        save_workbook_atomic(wb, target)
        z = zipfile.ZipFile(target)
        media = [n for n in z.namelist() if n.startswith("xl/media/")]
        assert len(media) == 60, len(media)
        assert "[Content_Types].xml" in z.namelist()
        assert _is_openable(target)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    test_atomic_save_writes_new_content_on_success()
    test_atomic_save_leaves_original_intact_when_save_raises()
    test_broken_image_is_dropped_instead_of_killing_save()
    test_atomic_save_leaves_no_temp_files_behind()
    test_atomic_save_creates_file_that_did_not_exist()
    test_picture_bytes_are_read_eagerly_not_at_save_time()
    test_unreadable_image_is_skipped_not_raised()
    test_missing_image_file_is_skipped()
    test_webp_is_converted_so_excel_can_show_it()
    test_embedded_image_is_compressed_not_full_resolution()
    test_workbook_is_verified_before_anything_touches_the_destination()
    test_save_survives_a_destination_that_lies_on_read_back()
    test_workbook_with_images_can_be_saved_twice()
    test_reloaded_workbook_images_survive_repeated_saves()
    test_many_images_all_land()
    print("ALL PASS")
