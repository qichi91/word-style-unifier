import asyncio
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from app.routes import router
from app.services import (
	DOWNLOAD_CLEANUP_INTERVAL_SECONDS,
	DOWNLOAD_TTL_SECONDS,
	cleanup_expired_downloads,
)


async def _cleanup_download_store_periodically() -> None:
	"""一定間隔でダウンロード用一時データを掃除する。"""

	while True:
		cleanup_expired_downloads(ttl_seconds=DOWNLOAD_TTL_SECONDS)
		await asyncio.sleep(DOWNLOAD_CLEANUP_INTERVAL_SECONDS)


@asynccontextmanager
async def lifespan(_: FastAPI):
	"""アプリ起動中だけ定期クリーンアップタスクを動かす。"""

	cleanup_task = asyncio.create_task(_cleanup_download_store_periodically())
	try:
		yield
	finally:
		cleanup_task.cancel()
		try:
			await cleanup_task
		except asyncio.CancelledError:
			pass


app = FastAPI(title="Word Style Unifier", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="static"), name="static")
app.include_router(router)
