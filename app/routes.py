import io
import secrets
import tempfile
import time
from pathlib import Path

from fastapi import APIRouter, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates

from app.services import (
    DEFAULT_FONTS,
    DOWNLOAD_STORE,
    ResultBlob,
    PROCESS_TIMEOUT_SECONDS,
    build_zip_bytes,
    cleanup_expired_downloads,
    normalize_document,
    sanitize_relative_path,
    save_upload,
    uniquify_path,
)


router = APIRouter()
templates = Jinja2Templates(directory="templates")


@router.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "request": request,
            "default_fonts": DEFAULT_FONTS,
            "default_font_a": "Arial",
            "default_font_b": "Yu Gothic",
        },
    )


@router.post("/convert", response_class=HTMLResponse)
async def convert(
    request: Request,
    font_a: str = Form("Arial"),
    font_b: str = Form("Yu Gothic"),
    single_file: UploadFile | None = File(default=None),
    folder_files: list[UploadFile] = File(default=[]),
):
    cleanup_expired_downloads()

    uploads: list[UploadFile] = []
    if single_file is not None and single_file.filename:
        uploads.append(single_file)
    uploads.extend([u for u in folder_files if u.filename])

    if not uploads:
        return templates.TemplateResponse(
            request=request,
            name="result.html",
            context={
                "request": request,
                "status": "error",
                "message": "アップロードされたファイルがありません。",
                "download_token": None,
                "converted_count": 0,
                "failed_count": 0,
                "skipped_count": 0,
            },
            status_code=400,
        )

    start = time.time()
    failed_entries: list[tuple[str, str]] = []
    skipped_entries: list[tuple[str, str]] = []
    converted_count = 0

    with tempfile.TemporaryDirectory(prefix="wsu_") as tmp:
        tmp_root = Path(tmp)
        input_root = tmp_root / "input"
        output_root = tmp_root / "output"

        used_paths: set[str] = set()
        saved: list[tuple[Path, Path]] = []

        for upload in uploads:
            raw_name = upload.filename or "upload.docx"
            rel_path = sanitize_relative_path(raw_name, Path(raw_name).name or "upload.docx")
            rel_path = uniquify_path(rel_path, used_paths)
            dest = input_root / rel_path
            try:
                await save_upload(upload, dest)
                saved.append((rel_path, dest))
            except ValueError as e:
                failed_entries.append((str(rel_path), str(e)))
            except Exception as e:
                failed_entries.append((str(rel_path), f"upload error: {e}"))

        for index_num, (rel_path, src_path) in enumerate(saved):
            elapsed = time.time() - start
            if elapsed > PROCESS_TIMEOUT_SECONDS:
                # 途中まで処理済みのものを残さず、残りはタイムアウト扱いにする。
                failed_entries.append((str(rel_path), "timeout before conversion"))
                for remain_rel, _ in saved[index_num + 1 :]:
                    failed_entries.append((str(remain_rel), "timeout before conversion"))
                break

            if src_path.suffix.lower() != ".docx":
                skipped_entries.append((str(rel_path), "skipped: not a .docx file"))
                continue

            dst_path = output_root / rel_path
            try:
                normalize_document(src_path, dst_path, font_a=font_a, font_b=font_b)
                converted_count += 1
            except Exception as e:
                failed_entries.append((str(rel_path), f"convert error: {e}"))

        failed_report = failed_entries + skipped_entries
        zip_bytes = build_zip_bytes(output_root, failed_report)

    token = secrets.token_urlsafe(24)
    # ZIPはメモリ上に保持し、token付きのダウンロードURLで返す。
    DOWNLOAD_STORE[token] = ResultBlob(
        filename="converted_results.zip",
        data=zip_bytes,
        created_at=time.time(),
    )

    status = "success" if converted_count > 0 else "error"
    message = "変換処理が完了しました。"
    if converted_count == 0:
        message = "変換対象の.docxが無いか、すべて失敗しました。"

    return templates.TemplateResponse(
        request=request,
        name="result.html",
        context={
            "request": request,
            "status": status,
            "message": message,
            "download_token": token,
            "converted_count": converted_count,
            "failed_count": len(failed_entries),
            "skipped_count": len(skipped_entries),
        },
    )


@router.get("/download/{token}")
def download(token: str):
    cleanup_expired_downloads()
    blob = DOWNLOAD_STORE.get(token)
    if blob is None:
        return StreamingResponse(
            io.BytesIO("ダウンロード情報が見つかりません。再実行してください。".encode("utf-8")),
            media_type="text/plain; charset=utf-8",
            status_code=404,
        )

    return StreamingResponse(
        io.BytesIO(blob.data),
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{blob.filename}"',
            "Content-Length": str(len(blob.data)),
        },
    )
