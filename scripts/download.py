"""
scripts/download.py -- 使用 XHS-Downloader 下载收藏笔记的详情与媒体文件

读取 datas/collect_list.json（由 fetch_list.py 生成），逐条调用
XHS-Downloader Python API 抓取笔记详情并下载图片/视频，同时在每条笔记的
下载目录中写入 info.json 和 detail.txt。

每处理 EXPORT_BATCH_SIZE 条笔记后，自动调用 export_from_db 导出 Excel/JSON
摘要，保证数据实时可查阅。

本脚本无需 Cookie — XHS-Downloader 通过解析页面内嵌的
window.__INITIAL_STATE__ 获取数据。
"""

import argparse
import asyncio
import json
import sys
from pathlib import Path

from loguru import logger

# ---------------------------------------------------------------------------
# 项目路径
# ---------------------------------------------------------------------------

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 将 XHS-Downloader 加入 sys.path（位于项目的同级目录）
_XHS_DOWNLOADER_ROOT = _PROJECT_ROOT.parent / "XHS-Downloader"
sys.path.insert(0, str(_XHS_DOWNLOADER_ROOT))

from source.application.app import XHS  # noqa: E402

from scripts.export import export_from_db  # noqa: E402

from apis.xhs_pc_apis import XHS_Apis  # noqa: E402
from xhs_utils.common_util import load_env  # noqa: E402
from xhs_utils.data_util import handle_note_info, download_note  # noqa: E402

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

COLLECT_LIST_PATH = _PROJECT_ROOT / "datas" / "collect_list.json"

WORK_PATH = _PROJECT_ROOT / "datas" / "xhs_download"
FOLDER_NAME = "Download"
DOWNLOAD_ROOT = WORK_PATH / FOLDER_NAME

EXPLORE_DATA_DB = DOWNLOAD_ROOT / "ExploreData.db"

EXPORT_EXCEL_PATH = _PROJECT_ROOT / "datas" / "excel_datas" / "我的收藏.xlsx"
EXPORT_JSON_PATH = _PROJECT_ROOT / "datas" / "excel_datas" / "我的收藏.json"

EXPORT_BATCH_SIZE = 10
MAX_CONSECUTIVE_FAILURES = 5
FAILURE_COOLDOWN_SECONDS = 60

# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------


def _load_collect_list() -> list[dict]:
    """读取 collect_list.json，返回笔记列表。"""
    if not COLLECT_LIST_PATH.exists():
        logger.error(f"收藏列表文件不存在: {COLLECT_LIST_PATH}")
        return []
    try:
        with open(COLLECT_LIST_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            logger.error("collect_list.json 格式异常，期望 list")
            return []
        logger.info(f"已加载收藏列表: {len(data)} 条")
        return data
    except (json.JSONDecodeError, OSError) as exc:
        logger.error(f"读取收藏列表失败: {exc}")
        return []


def _scan_completed_ids() -> set[str]:
    """扫描 Download 目录，找到已有 info.json 的笔记 ID（真正完成的）。"""
    completed = set()
    if not DOWNLOAD_ROOT.exists():
        return completed
    for info_path in DOWNLOAD_ROOT.rglob("info.json"):
        parts = info_path.parent.name.rsplit("_", 1)
        if len(parts) == 2:
            completed.add(parts[1])
    logger.info(f"本地已完成: {len(completed)} 条（有 info.json）")
    return completed


def _build_note_url(note_id: str, xsec_token: str) -> str:
    """构建笔记的完整 URL。"""
    return (
        f"https://www.xiaohongshu.com/discovery/item/{note_id}"
        f"?xsec_token={xsec_token}&xsec_source=pc_share"
    )


# ---------------------------------------------------------------------------
# 数据转换
# ---------------------------------------------------------------------------


def _convert_to_info_json(result: dict, note_url: str) -> dict:
    """将 XHS-Downloader 的 extract 结果转换为统一的 info.json 格式。"""
    raw_type = result.get("作品类型", "")
    note_type = "图集" if raw_type in ("图文", "图集") else "视频"

    tags_str = result.get("作品标签", "")
    tags = [t.strip() for t in tags_str.split() if t.strip()] if tags_str else []

    download_urls = result.get("下载地址", [])
    if note_type == "图集":
        image_list = [u for u in download_urls if u]
        video_addr = None
    else:
        image_list = []
        video_addr = download_urls[0] if download_urls else None

    return {
        "note_id": result.get("作品ID", ""),
        "note_url": result.get("作品链接", note_url),
        "note_type": note_type,
        "user_id": result.get("作者ID", ""),
        "home_url": result.get("作者链接", ""),
        "nickname": result.get("作者昵称", ""),
        "avatar": "",
        "title": result.get("作品标题", ""),
        "desc": result.get("作品描述", ""),
        "liked_count": str(result.get("点赞数量", "")),
        "collected_count": str(result.get("收藏数量", "")),
        "comment_count": str(result.get("评论数量", "")),
        "share_count": str(result.get("分享数量", "")),
        "video_cover": None,
        "video_addr": video_addr,
        "image_list": image_list,
        "tags": tags,
        "upload_time": result.get("发布时间", ""),
        "ip_location": "",
    }


def _write_info_json(note_info: dict, note_dir: Path) -> None:
    """将 info.json 写入笔记下载目录。"""
    note_dir.mkdir(parents=True, exist_ok=True)
    with open(note_dir / "info.json", "w", encoding="utf-8") as f:
        json.dump(note_info, f, ensure_ascii=False, indent=2)


def _find_note_dir(note_id: str) -> Path | None:
    """
    在下载根目录中搜索以 _{note_id} 结尾的目录。

    XHS-Downloader 的 folder_mode + author_archive 会生成类似
    Download/{author}/{title}_{note_id}/ 的目录结构。
    """
    if not DOWNLOAD_ROOT.exists():
        return None
    for candidate in DOWNLOAD_ROOT.rglob(f"*_{note_id}"):
        if candidate.is_dir():
            return candidate
    return None


def _fallback_spider_xhs(note_id: str, xsec_token: str, cookies_str: str) -> bool:
    """XHS-Downloader 失败后，用 Spider_XHS（Cookie）作为 fallback。

    下载到 DOWNLOAD_ROOT 下，目录格式 {nickname}_{user_id}/{title}_{note_id}/，
    与 _scan_completed_ids 兼容。
    """
    note_url = (
        f"https://www.xiaohongshu.com/explore/{note_id}"
        f"?xsec_token={xsec_token}&xsec_source=pc_user"
    )
    xhs_apis = XHS_Apis()
    success, msg, res_json = xhs_apis.get_note_info(note_url, cookies_str)
    if not success or not res_json:
        logger.warning(f"  -> Spider_XHS fallback 也失败: {msg}")
        return False

    items = (res_json.get("data") or {}).get("items")
    if not items:
        logger.warning(f"  -> Spider_XHS fallback: items 为空 ({msg})")
        return False

    note_info = items[0]
    note_info["url"] = note_url
    note_info = handle_note_info(note_info)
    download_note(note_info, str(DOWNLOAD_ROOT), "all")
    logger.info(f"  -> Spider_XHS fallback 成功: {note_info['title'][:30]}")
    return True


def _run_export() -> None:
    """调用 export 模块导出数据库到 Excel/JSON。"""
    try:
        export_from_db(
            db_path=EXPLORE_DATA_DB,
            excel_path=EXPORT_EXCEL_PATH,
            json_path=EXPORT_JSON_PATH,
        )
    except Exception as exc:
        logger.warning(f"导出失败（不影响下载流程）: {exc}")


# ---------------------------------------------------------------------------
# 核心下载逻辑
# ---------------------------------------------------------------------------


async def _download_all(notes: list[dict], cookies_str: str | None = None) -> None:
    """
    使用 XHS-Downloader 逐条下载笔记。
    XHS-Downloader 失败时自动 fallback 到 Spider_XHS（需要 Cookie）。

    连续失败 MAX_CONSECUTIVE_FAILURES 次后休眠 FAILURE_COOLDOWN_SECONDS 秒，
    然后重置计数器继续处理。
    """
    total = len(notes)
    logger.info(f"开始下载: 共 {total} 条待处理")

    async with XHS(
        work_path=str(WORK_PATH),
        folder_name=FOLDER_NAME,
        name_format="作品标题 作品ID",
        author_archive=True,
        folder_mode=True,
        record_data=True,
        download_record=False,
        image_format="JPEG",
        image_download=True,
        video_download=True,
    ) as xhs:
        consecutive_failures = 0
        processed = 0

        for idx, note in enumerate(notes, start=1):
            note_id = note.get("note_id") or note.get("id", "")
            xsec_token = note.get("xsec_token", "")

            if not note_id:
                logger.warning(f"[{idx}/{total}] 笔记缺少 note_id，跳过")
                continue

            url = _build_note_url(note_id, xsec_token)
            logger.info(f"[{idx}/{total}] 处理笔记: {note_id}")

            try:
                result = await xhs.extract(url=url, download=True)

                # extract 返回列表；result[0] 为空 dict 或含数据的 dict
                if not result or not result[0] or not result[0].get("作品ID"):
                    logger.warning(f"[{idx}/{total}] 笔记 {note_id} XHS-Downloader 失败，尝试 Spider_XHS fallback")
                    if cookies_str and _fallback_spider_xhs(note_id, xsec_token, cookies_str):
                        consecutive_failures = 0
                        processed += 1
                    else:
                        consecutive_failures += 1
                else:
                    data = result[0]
                    note_info = _convert_to_info_json(data, url)

                    # 查找下载目录并写入 info.json / detail.txt
                    note_dir = _find_note_dir(note_id)
                    if note_dir:
                        _write_info_json(note_info, note_dir)
                        logger.info(
                            f"[{idx}/{total}] 笔记 {note_id} 完成，"
                            f"已写入 {note_dir}"
                        )
                    else:
                        logger.warning(
                            f"[{idx}/{total}] 笔记 {note_id} 下载目录未找到，"
                            f"跳过写入 info.json"
                        )

                    consecutive_failures = 0
                    processed += 1

            except Exception as exc:
                logger.error(f"[{idx}/{total}] 笔记 {note_id} 处理异常: {exc}，尝试 Spider_XHS fallback")
                if cookies_str and _fallback_spider_xhs(note_id, xsec_token, cookies_str):
                    consecutive_failures = 0
                    processed += 1
                else:
                    consecutive_failures += 1

            # 连续失败冷却
            if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                logger.warning(
                    f"连续失败 {consecutive_failures} 次，"
                    f"休眠 {FAILURE_COOLDOWN_SECONDS} 秒后继续..."
                )
                await asyncio.sleep(FAILURE_COOLDOWN_SECONDS)
                consecutive_failures = 0

            # 批量导出
            if processed > 0 and processed % EXPORT_BATCH_SIZE == 0:
                logger.info(f"已处理 {processed} 条，触发导出...")
                _run_export()

        # 最终导出
        if processed > 0:
            logger.info(f"下载完成，共成功 {processed}/{total} 条，执行最终导出...")
            _run_export()
        else:
            logger.info("本次无新笔记被成功处理")


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------


def main(limit: int | None = None) -> None:
    """
    下载收藏笔记的详情与媒体文件。

    参数
    ----
    limit : 最多处理的未完成笔记数（None 表示全部处理）
    """
    # 1. 加载收藏列表
    all_notes = _load_collect_list()
    if not all_notes:
        return

    # 2. 扫描本地已完成的笔记（有 info.json 的），过滤掉
    completed_ids = _scan_completed_ids()
    pending_notes = [
        n
        for n in all_notes
        if (n.get("note_id") or n.get("id", "")) not in completed_ids
    ]
    logger.info(
        f"收藏列表 {len(all_notes)} 条，"
        f"已完成 {len(completed_ids)} 条，"
        f"待处理 {len(pending_notes)} 条"
    )

    if not pending_notes:
        logger.info("所有笔记已下载完毕，无需处理")
        return

    # 3. 应用 limit
    if limit is not None and limit > 0:
        pending_notes = pending_notes[:limit]
        logger.info(f"应用 --limit {limit}，本次处理 {len(pending_notes)} 条")

    # 4. 加载 Cookie 用于 fallback
    cookies_str = load_env()
    if not cookies_str:
        logger.warning("未找到 Cookie，Spider_XHS fallback 不可用")

    # 5. 执行下载
    asyncio.run(_download_all(pending_notes, cookies_str))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="下载小红书收藏笔记的详情与媒体文件"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="最多处理的未完成笔记数",
    )
    args = parser.parse_args()
    main(limit=args.limit)
