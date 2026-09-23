"""把采购单附件读成「可交给规则或模型看」的文本/图片。

支持四类附件（实测语料里的全部形态）：

===========  ==========================================  ==========================
后缀          读法                                        交给谁
===========  ==========================================  ==========================
xlsx/xlsm    openpyxl（只读、取计算值）                    规则优先，必要时给模型
xls          xlrd（老式 OLE 二进制）                       同上
pdf          pypdf 抽文本层；无文本层则标为需视觉           模型
png/jpg/…    直接给图片                                     模型（视觉）
===========  ==========================================  ==========================

输出统一成 `Attachment`：`kind` + `sheets`（表格）+ `text`（PDF 文本）+ `images`。
`sheets_to_text()` 是给模型看的紧凑表示（带行号，空格压缩，跳过空行）。
"""
from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

TABLE_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
MEDIA_ROOT = Path(__file__).resolve().parents[2] / ".cache" / "erp" / "attachments" / "media"
SPARSE_CELLS = 40   # 文字单元格少于这个数、又有嵌入图时，认为"数据写在图片里"


def extract_embedded_media(path: Path) -> list[Path]:
    """导出 xlsx 里嵌入的图片。

    供应商常把**扫描/拍照的入库单**直接贴在 Excel 里，单元格里只剩几个数字
    （预付款/尾款）。只读单元格就会整单取不到金额（实测 `26MT-05Q215`），
    所以这类附件要把图片交给视觉模型读。
    """
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return []
    try:
        with zipfile.ZipFile(path) as archive:
            names = [
                name
                for name in archive.namelist()
                if name.startswith("xl/media/") and not name.endswith("/")
            ]
            if not names:
                return []
            target_dir = MEDIA_ROOT / path.parent.name / path.stem
            out: list[Path] = []
            for name in names:
                target = target_dir / Path(name).name
                info = archive.getinfo(name)
                if not target.exists() or target.stat().st_size != info.file_size:
                    target_dir.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(archive.read(name))
                out.append(target)
            return out
    except Exception:  # noqa: BLE001
        return []


@dataclass
class Sheet:
    name: str
    rows: list[list[str]]
    n_rows: int = 0
    n_cols: int = 0


@dataclass
class Attachment:
    path: Path
    kind: str                      # table | pdf | image | unsupported
    sheets: list[Sheet] = field(default_factory=list)
    text: str = ""
    images: list[Path] = field(default_factory=list)
    embedded: list[Path] = field(default_factory=list)   # xlsx 里嵌的图（多为扫描件/签字章）
    note: str = ""

    @property
    def digest(self) -> str:
        data = self.path.read_bytes()
        return hashlib.sha256(data).hexdigest()[:16]


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        # 6.0 -> 6，避免模型把数量读成小数
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def _read_xlsx(path: Path) -> list[Sheet]:
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets: list[Sheet] = []
    try:
        for worksheet in workbook.worksheets:
            rows = [
                [_cell_text(value) for value in row]
                for row in worksheet.iter_rows(values_only=True)
            ]
            width = max((len(row) for row in rows), default=0)
            for row in rows:
                row.extend([""] * (width - len(row)))
            sheets.append(
                Sheet(name=worksheet.title, rows=rows, n_rows=len(rows), n_cols=width)
            )
    finally:
        workbook.close()
    return sheets


def _read_xls(path: Path) -> list[Sheet]:
    import xlrd

    book = xlrd.open_workbook(str(path))
    sheets: list[Sheet] = []
    for sheet in book.sheets():
        rows: list[list[str]] = []
        for index in range(sheet.nrows):
            rows.append([_cell_text(value) for value in sheet.row_values(index)])
        width = max((len(row) for row in rows), default=0)
        for row in rows:
            row.extend([""] * (width - len(row)))
        sheets.append(Sheet(name=sheet.name, rows=rows, n_rows=sheet.nrows, n_cols=width))
    return sheets


def _read_pdf(path: Path) -> tuple[str, str]:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    chunks: list[str] = []
    for page in reader.pages:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:  # noqa: BLE001
            chunks.append("")
    text = "\n".join(chunks).strip()
    note = "" if len(text) >= 30 else "PDF 无文本层（需视觉识别）"
    return text, note


def read_attachment(path: Path | str) -> Attachment:
    """按后缀分派读取。异常不抛出，而是记在 `note` 里。"""
    target = Path(path)
    suffix = target.suffix.lower()
    if suffix in TABLE_SUFFIXES:
        try:
            sheets = _read_xlsx(target) if suffix != ".xls" else _read_xls(target)
            cells = sum(
                1
                for sheet in sheets
                for row in sheet.rows
                for cell in row
                if cell and str(cell).strip()
            )
            embedded = extract_embedded_media(target) if cells < SPARSE_CELLS else []
            return Attachment(
                path=target, kind="table", sheets=sheets, embedded=embedded
            )
        except Exception as exc:  # noqa: BLE001
            return Attachment(
                path=target, kind="unsupported", note=f"表格读取失败: {type(exc).__name__}: {exc}"
            )
    if suffix == ".pdf":
        try:
            text, note = _read_pdf(target)
            return Attachment(path=target, kind="pdf", text=text, note=note)
        except Exception as exc:  # noqa: BLE001
            return Attachment(
                path=target, kind="unsupported", note=f"PDF 读取失败: {type(exc).__name__}: {exc}"
            )
    if suffix in IMAGE_SUFFIXES:
        return Attachment(path=target, kind="image", images=[target])
    return Attachment(path=target, kind="unsupported", note=f"未知后缀 {suffix}")


def used_width(rows: list[list[str]]) -> int:
    """表里**真的有内容**的列数（忽略右侧空白列）。"""
    width = 0
    for row in rows:
        for index in range(len(row) - 1, -1, -1):
            if row[index]:
                width = max(width, index + 1)
                break
    return width


def sheets_to_text(sheets: list[Sheet], *, max_rows: int = 300, max_cols: int = 60) -> str:
    """把表格渲染成带行号的紧凑文本，供模型阅读。

    2026-09-24 修：旧版硬截断在 24 列，把宽表右侧的实发区块直接丢掉了
    （`25MT-03P005-BCD` 的 C/D 批在第 25~36 列，模型因此漏读两批）。
    现在按「实际有内容的列」渲染，上限 60 列；真的还有列没展示时会明确写出。
    """
    lines: list[str] = []
    for sheet in sheets[:3]:
        width = used_width(sheet.rows)
        shown_cols = min(width, max_cols)
        head = f"[sheet] {sheet.name}  行数={sheet.n_rows} 列数={width}"
        if width > shown_cols:
            head += f"（只展示第 1~{shown_cols} 列）"
        lines.append(head)
        for index, row in enumerate(sheet.rows[:max_rows], 1):
            cells = row[:shown_cols]
            if not any(cells):
                continue
            body = " | ".join(cells).rstrip(" |")
            lines.append(f"r{index}: {body}")
        if width > shown_cols:
            lines.append(f"...（该表第 {shown_cols + 1}~{width} 列未展示）")
        if sheet.n_rows > max_rows:
            lines.append(f"...（该表还有 {sheet.n_rows - max_rows} 行未展示）")
    return "\n".join(lines)
