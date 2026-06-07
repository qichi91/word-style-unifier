import copy
import io
import time
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable

import jaconv
from docx import Document
from docx.oxml.ns import qn
from fastapi import UploadFile


MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024
PROCESS_TIMEOUT_SECONDS = 5 * 60
DEFAULT_FONTS = [
    "MS Gothic",
    "MS Mincho",
    "Yu Gothic",
    "Meiryo",
    "Arial",
]


@dataclass
class ResultBlob:
    filename: str
    data: bytes
    created_at: float


DOWNLOAD_STORE: dict[str, ResultBlob] = {}


def cleanup_expired_downloads(ttl_seconds: int = 30 * 60) -> None:
    now = time.time()
    expired = [k for k, v in DOWNLOAD_STORE.items() if (now - v.created_at) > ttl_seconds]
    for key in expired:
        DOWNLOAD_STORE.pop(key, None)


def sanitize_relative_path(raw_name: str, fallback_name: str) -> Path:
    candidate = raw_name.strip().replace("\\", "/")
    if not candidate:
        candidate = fallback_name
    posix = PurePosixPath(candidate)
    safe_parts = [p for p in posix.parts if p not in ("", ".", "..")]
    if not safe_parts:
        safe_parts = [Path(fallback_name).name or "upload.docx"]
    return Path(*safe_parts)


def uniquify_path(path: Path, used: set[str]) -> Path:
    path_str = str(path)
    if path_str not in used:
        used.add(path_str)
        return path

    suffix_index = 1
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    while True:
        candidate = parent / f"{stem}_{suffix_index}{suffix}"
        candidate_str = str(candidate)
        if candidate_str not in used:
            used.add(candidate_str)
            return candidate
        suffix_index += 1


async def save_upload(upload: UploadFile, destination: Path) -> int:
    destination.parent.mkdir(parents=True, exist_ok=True)
    total = 0
    with destination.open("wb") as f:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_FILE_SIZE_BYTES:
                raise ValueError(f"size limit exceeded: {MAX_FILE_SIZE_BYTES} bytes")
            f.write(chunk)
    await upload.close()
    return total


def is_halfwidth_ascii_alnum_symbol(ch: str) -> bool:
    code = ord(ch)
    return 0x20 <= code <= 0x7E


def apply_font(run, font_name: str) -> None:
    run.font.name = font_name
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def copy_run_properties(src_run, dst_run) -> None:
    src_rpr = src_run._element.rPr
    if src_rpr is None:
        return

    dst_rpr = dst_run._element.get_or_add_rPr()
    for child in list(dst_rpr):
        dst_rpr.remove(child)
    for key, value in src_rpr.attrib.items():
        dst_rpr.set(key, value)
    for child in src_rpr:
        dst_rpr.append(copy.deepcopy(child))


def split_segments(text: str) -> list[tuple[str, bool]]:
    if not text:
        return []

    converted = jaconv.h2z(text, kana=True, ascii=False, digit=False)
    segments: list[tuple[str, bool]] = []
    buff = [converted[0]]
    current_ascii = is_halfwidth_ascii_alnum_symbol(converted[0])
    for ch in converted[1:]:
        is_ascii = is_halfwidth_ascii_alnum_symbol(ch)
        if is_ascii == current_ascii:
            buff.append(ch)
            continue
        segments.append(("".join(buff), current_ascii))
        buff = [ch]
        current_ascii = is_ascii
    segments.append(("".join(buff), current_ascii))
    return segments


def normalize_paragraph(paragraph, font_a: str, font_b: str) -> None:
    runs = list(paragraph.runs)
    for run in runs:
        original_text = run.text
        segments = split_segments(original_text)
        if not segments:
            continue

        first_text, first_is_ascii = segments[0]
        run.text = first_text
        apply_font(run, font_a if first_is_ascii else font_b)

        previous = run
        for text, is_ascii in segments[1:]:
            new_run = paragraph.add_run(text)
            copy_run_properties(run, new_run)
            apply_font(new_run, font_a if is_ascii else font_b)
            previous._element.addnext(new_run._element)
            previous = new_run


def iter_paragraphs_from_container(container) -> Iterable:
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs_from_container(cell)


def normalize_document(input_path: Path, output_path: Path, font_a: str, font_b: str) -> None:
    document = Document(str(input_path))

    for paragraph in iter_paragraphs_from_container(document):
        normalize_paragraph(paragraph, font_a, font_b)

    for section in document.sections:
        for paragraph in iter_paragraphs_from_container(section.header):
            normalize_paragraph(paragraph, font_a, font_b)
        for paragraph in iter_paragraphs_from_container(section.footer):
            normalize_paragraph(paragraph, font_a, font_b)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))


def build_zip_bytes(output_root: Path, failed_entries: list[tuple[str, str]]) -> bytes:
    buffer = io.BytesIO()

    import zipfile

    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(output_root.rglob("*")):
            if file_path.is_file():
                zf.write(file_path, arcname=str(file_path.relative_to(output_root)))

        if failed_entries:
            lines = ["path\treason"]
            lines.extend(f"{path}\t{reason}" for path, reason in failed_entries)
            zf.writestr("faield.txt", "\n".join(lines) + "\n")

    buffer.seek(0)
    return buffer.getvalue()
