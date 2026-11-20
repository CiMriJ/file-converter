"""文件转换核心引擎。

覆盖：Word -> PDF、PDF -> Word、图片 -> PDF、PDF -> 图片、图片格式互转，
以及 PDF 合并 / 拆分 / 旋转 / 加密 / 解密 / 提取文字 / 提取图片 / 压缩。
"""

from __future__ import annotations

import gc
import os
import re
import shutil
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageOps

# ---------------------------------------------------------------- 常量

WORD_EXTS = {".docx", ".doc", ".rtf", ".odt", ".txt"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
PDF_EXTS = {".pdf"}
IMAGE_TARGET_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}

DOCX_SAVE_FORMAT = 16  # wdFormatDocumentDefault
PDF_SAVE_FORMAT = 17  # wdFormatPDF


class ConversionError(RuntimeError):
    """转换过程中的可读错误。"""


# ---------------------------------------------------------------- 小工具


def ext_of(path: str | Path) -> str:
    return Path(path).suffix.lower()


def unique_path(path: str | Path) -> Path:
    """目标文件已存在时自动追加 (1)、(2)……，避免覆盖。"""
    path = Path(path)
    if not path.exists():
        return path
    index = 1
    while True:
        candidate = path.with_name(f"{path.stem}({index}){path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def ensure_dir(path: str | Path) -> Path:
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def _split_groups(text: str) -> list[str]:
    return [part for part in re.split(r"[,，;；\n]+", text.strip()) if part.strip()]


def parse_page_ranges(text: str, total: int) -> list[int]:
    """解析 "1-3,5,8-" 形式的页码范围，返回去重的 1 起页码列表。"""
    text = (text or "").strip()
    if not text or text.lower() in {"all", "全部", "*"}:
        return list(range(1, total + 1))

    pages: list[int] = []
    for part in re.split(r"[,，;；\s]+", text):
        part = part.strip()
        if not part:
            continue
        match = re.fullmatch(r"(\d+)?\s*[-~至到]\s*(\d+)?", part)
        if match:
            start = int(match.group(1) or 1)
            end = int(match.group(2) or total)
            if start > end:
                start, end = end, start
            pages.extend(range(start, end + 1))
        elif part.isdigit():
            pages.append(int(part))
        else:
            raise ConversionError(f"无法识别的页码范围：{part}")

    result: list[int] = []
    for page in pages:
        if 1 <= page <= total and page not in result:
            result.append(page)
    if not result:
        raise ConversionError(f"页码范围超出文件总页数（共 {total} 页）")
    return result


def pdf_page_count(path: str | Path) -> int:
    with fitz.open(path) as doc:
        return doc.page_count


# ---------------------------------------------------------------- Word -> PDF


def word_to_pdf(src: str | Path, dst: str | Path) -> None:
    """调用本机 Microsoft Word 导出 PDF，排版与原文档保持一致。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise ConversionError("缺少 pywin32 组件，无法调用 Word 进行转换") from exc

    src_abs = os.path.abspath(str(src))
    dst_abs = os.path.abspath(str(dst))
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(src_abs, ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs(dst_abs, FileFormat=PDF_SAVE_FORMAT)
        finally:
            doc.Close(0)
            doc = None
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Word 转换失败：{exc}") from exc
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        # 必须先释放 COM 引用，再反初始化，否则退出时可能崩溃
        doc = None
        word = None
        gc.collect()
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------- PDF <-> Word


def pdf_to_docx(src: str | Path, dst: str | Path, pages: str = "") -> list[str]:
    """基于 pdf2docx 按版式重建 Word 文档（页码范围需连续）。"""
    try:
        from pdf2docx import Converter
    except ImportError as exc:  # pragma: no cover
        raise ConversionError("缺少 pdf2docx 组件") from exc

    total = pdf_page_count(src)
    page_list = parse_page_ranges(pages, total)
    start, end = min(page_list), max(page_list)
    try:
        converter = Converter(str(src))
        try:
            converter.convert(str(dst), start=start - 1, end=end)
        finally:
            converter.close()
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"PDF 转 Word 失败：{exc}") from exc
    return [str(dst)]


# ---------------------------------------------------------------- 图片 <-> PDF


def _flatten_to_rgb(image: Image.Image) -> Image.Image:
    """带透明通道的图片合成到白底。"""
    if image.mode == "RGB":
        return image
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        background = Image.new("RGB", rgba.size, "white")
        background.paste(rgba, mask=rgba.split()[-1])
        return background
    return image.convert("RGB")


def images_to_pdf(paths: list[str | Path], dst: str | Path, quality: int = 90) -> list[str]:
    """把一张或多张图片合成一个 PDF（自动按 EXIF 纠正方向）。"""
    if not paths:
        raise ConversionError("没有可用的图片")
    frames: list[Image.Image] = []
    for path in paths:
        with Image.open(path) as image:
            image.load()
            frames.append(_flatten_to_rgb(ImageOps.exif_transpose(image)))
    if not frames:
        raise ConversionError("没有可用的图片")
    dst = Path(dst)
    frames[0].save(
        dst,
        "PDF",
        resolution=150.0,
        save_all=len(frames) > 1,
        append_images=frames[1:],
        quality=int(quality),
    )
    return [str(dst)]


def pdf_to_images(
    src: str | Path,
    out_dir: str | Path,
    fmt: str = "png",
    dpi: int = 150,
    pages: str = "",
    prefix: str | None = None,
    quality: int = 90,
) -> list[str]:
    """把 PDF 的指定页渲染成图片。"""
    fmt = fmt.lower().lstrip(".")
    if fmt == "jpeg":
        fmt = "jpg"
    out_dir = ensure_dir(out_dir)
    prefix = prefix or Path(src).stem
    outputs: list[str] = []
    with fitz.open(src) as doc:
        for page_no in parse_page_ranges(pages, doc.page_count):
            page = doc[page_no - 1]
            pixmap = page.get_pixmap(dpi=int(dpi))
            mode = "RGBA" if pixmap.alpha else "RGB"
            image = Image.frombytes(mode, (pixmap.width, pixmap.height), pixmap.samples)
            if fmt in {"jpg", "bmp"}:
                image = _flatten_to_rgb(image)
            extension = "jpg" if fmt == "jpeg" else fmt
            dst = unique_path(out_dir / f"{prefix}_第{page_no:03d}页.{extension}")
            params: dict = {}
            if fmt in {"jpg", "webp"}:
                params["quality"] = int(quality)
            elif fmt == "png":
                params["optimize"] = True
            image.save(dst, **params)
            outputs.append(str(dst))
    return outputs


def convert_image(
    src: str | Path,
    dst: str | Path,
    quality: int = 90,
    max_width: int = 0,
) -> list[str]:
    """图片格式互转，可选按最长边等比缩放。"""
    target = ext_of(dst).lstrip(".")
    if target == "jpeg":
        target = "jpg"
    supported = {"jpg", "jpeg", "png", "webp", "bmp"}
    if target not in supported:
        raise ConversionError(f"不支持输出为 {target} 格式")
    with Image.open(src) as image:
        image.load()
        image = ImageOps.exif_transpose(image)
        if max_width and image.width > max_width:
            ratio = max_width / float(image.width)
            new_size = (int(max_width), max(1, round(image.height * ratio)))
            image = image.resize(new_size, Image.LANCZOS)
        if target in {"jpg", "bmp"}:
            image = _flatten_to_rgb(image)
        params: dict = {}
        if target in {"jpg", "webp"}:
            params["quality"] = int(quality)
        elif target == "png":
            params["optimize"] = True
        image.save(dst, **params)
    return [str(dst)]


# ---------------------------------------------------------------- 通用入口


def convert_file(
    src: str | Path,
    dst: str | Path,
    *,
    quality: int = 90,
    dpi: int = 150,
    max_width: int = 0,
    pages: str = "",
) -> list[str]:
    """按扩展名自动路由的单文件转换，返回实际产物路径列表。"""
    src_ext = ext_of(src)
    dst_ext = ext_of(dst)

    if dst_ext == ".pdf":
        if src_ext in WORD_EXTS:
            word_to_pdf(src, dst)
            return [str(dst)]
        if src_ext in IMAGE_EXTS:
            return images_to_pdf([src], dst, quality)
        if src_ext == ".pdf":
            shutil.copy2(src, dst)
            return [str(dst)]
        raise ConversionError(f"暂不支持 {src_ext} 转 PDF")

    if dst_ext in {".docx", ".doc"}:
        if src_ext == ".pdf":
            return pdf_to_docx(src, dst, pages)
        raise ConversionError("转 Word 仅支持 PDF 来源")

    if dst_ext in IMAGE_TARGET_EXTS:
        fmt = dst_ext.lstrip(".").replace("jpeg", "jpg")
        if src_ext == ".pdf":
            return pdf_to_images(src, Path(dst).parent, fmt, dpi, pages, Path(dst).stem, quality)
        if src_ext in IMAGE_EXTS:
            return convert_image(src, dst, quality, max_width)
        raise ConversionError(f"暂不支持 {src_ext} 转图片")

    raise ConversionError(f"不支持的目标格式：{dst_ext}")


# ---------------------------------------------------------------- PDF 工具箱


def pdf_merge(paths: list[str | Path], dst: str | Path) -> list[str]:
    if len(paths) < 2:
        raise ConversionError("合并 PDF 至少需要 2 个文件")
    merged = fitz.open()
    try:
        for path in paths:
            with fitz.open(path) as doc:
                merged.insert_pdf(doc)
        merged.save(str(dst), garbage=4, deflate=True, use_objstms=1)
    finally:
        merged.close()
    return [str(dst)]


def pdf_split(
    src: str | Path,
    out_dir: str | Path,
    mode: str = "page",
    chunk: int = 1,
    ranges: str = "",
) -> list[str]:
    """mode: page=每页一个 / chunk=每 N 页一个 / range=按逗号分组拆分。"""
    out_dir = ensure_dir(out_dir)
    stem = Path(src).stem
    outputs: list[str] = []
    groups: list[list[int]] = []

    with fitz.open(src) as doc:
        total = doc.page_count
        if mode == "page":
            groups = [[page] for page in range(1, total + 1)]
        elif mode == "chunk":
            size = max(1, int(chunk))
            groups = [list(range(start, min(start + size, total + 1))) for start in range(1, total + 1, size)]
        else:
            if not ranges.strip():
                raise ConversionError("请输入拆分范围，例如 1-3,4-6")
            for group in _split_groups(ranges):
                groups.append(parse_page_ranges(group, total))

        for index, pages in enumerate(groups, 1):
            part = fitz.open()
            try:
                for page_no in pages:
                    part.insert_pdf(doc, from_page=page_no - 1, to_page=page_no - 1)
                dst = unique_path(out_dir / f"{stem}_第{index}部分.pdf")
                part.save(str(dst), garbage=4, deflate=True, use_objstms=1)
                outputs.append(str(dst))
            finally:
                part.close()
    return outputs


def pdf_rotate(
    src: str | Path,
    dst: str | Path,
    angle: int = 90,
    pages: str = "",
) -> list[str]:
    angle %= 360
    with fitz.open(src) as doc:
        for page_no in parse_page_ranges(pages, doc.page_count):
            page = doc[page_no - 1]
            page.set_rotation((page.rotation + angle) % 360)
        doc.save(str(dst), garbage=4, deflate=True)
    return [str(dst)]


def pdf_encrypt(
    src: str | Path,
    dst: str | Path,
    password: str,
    allow_print: bool = True,
    allow_copy: bool = True,
) -> list[str]:
    if not password:
        raise ConversionError("请输入打开密码")
    permissions = fitz.PDF_PERM_ACCESSIBILITY
    if allow_print:
        permissions |= fitz.PDF_PERM_PRINT
    if allow_copy:
        permissions |= fitz.PDF_PERM_COPY
    with fitz.open(src) as doc:
        doc.save(
            str(dst),
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=password,
            user_pw=password,
            permissions=permissions,
            garbage=4,
            deflate=True,
        )
    return [str(dst)]


def pdf_decrypt(src: str | Path, dst: str | Path, password: str) -> list[str]:
    try:
        doc = fitz.open(src)
    except Exception as exc:
        raise ConversionError(f"无法打开 PDF：{exc}") from exc
    try:
        if doc.needs_pass and not doc.authenticate(password or ""):
            raise ConversionError("密码错误，无法解密")
        doc.save(str(dst), encryption=fitz.PDF_ENCRYPT_NONE, garbage=4, deflate=True)
    finally:
        doc.close()
    return [str(dst)]


def pdf_extract_text(
    src: str | Path,
    dst: str | Path,
    pages: str = "",
    keep_layout: bool = False,
) -> list[str]:
    with fitz.open(src) as doc:
        page_numbers = parse_page_ranges(pages, doc.page_count)
        chunks: list[str] = []
        for page_no in page_numbers:
            page = doc[page_no - 1]
            text = page.get_text("text", sort=bool(keep_layout)) or ""
            if len(page_numbers) > 1:
                chunks.append(f"===== 第 {page_no} 页 =====\n{text}")
            else:
                chunks.append(text)
    Path(dst).write_text("\n\n".join(chunks), encoding="utf-8")
    return [str(dst)]


def pdf_extract_images(
    src: str | Path,
    out_dir: str | Path,
    min_size: int = 100,
) -> list[str]:
    out_dir = ensure_dir(out_dir)
    stem = Path(src).stem
    outputs: list[str] = []
    seen: set[int] = set()
    with fitz.open(src) as doc:
        for page_index in range(doc.page_count):
            page = doc[page_index]
            for info in page.get_images(full=True):
                xref = info[0]
                if xref in seen:
                    continue
                seen.add(xref)
                try:
                    raw = doc.extract_image(xref)
                except Exception:
                    continue
                if raw["width"] < min_size and raw["height"] < min_size:
                    continue
                dst = unique_path(out_dir / f"{stem}_p{page_index + 1}_{xref}.{raw['ext']}")
                dst.write_bytes(raw["image"])
                outputs.append(str(dst))
    if not outputs:
        raise ConversionError("没有找到符合尺寸要求的图片")
    return outputs


def pdf_compress(src: str | Path, dst: str | Path, level: str = "standard") -> list[str]:
    presets = {
        "light": {"dpi_threshold": 300, "dpi_target": 220, "quality": 88},
        "standard": {"dpi_threshold": 200, "dpi_target": 140, "quality": 75},
        "strong": {"dpi_threshold": 150, "dpi_target": 96, "quality": 60},
    }
    settings = presets.get(level, presets["standard"])
    with fitz.open(src) as doc:
        try:
            doc.subset_fonts()
        except Exception:
            pass
        try:
            doc.rewrite_images(**settings)
        except Exception:
            pass
        doc.save(
            str(dst),
            garbage=4,
            clean=True,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            use_objstms=1,
        )
    return [str(dst)]
