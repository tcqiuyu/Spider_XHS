"""
scripts/fetch_list.py — 获取小红书收藏列表并导出摘要

从 Xiaohongshu 用户收藏 API 拉取收藏笔记列表，写入 datas/collect_list.json，
并同步导出 Excel/JSON 摘要供人工查阅。

行为：
  - 若 collect_list.json 不存在 → 全量拉取所有页（每页最多 30 条），直到
    has_more 为 False。
  - 若 collect_list.json 已存在 → 增量检测：逐页拉取，遇到第一个已缓存
    note_id 时停止，将新笔记追加到列表头部。

输出：
  datas/collect_list.json            — 原始 API 响应笔记列表（全量缓存）
  datas/excel_datas/收藏列表.xlsx    — 摘要 Excel（笔记ID/标题/类型/作者/点赞数/链接）
  datas/excel_datas/收藏列表.json    — 摘要 JSON（同上字段）

仅此脚本需要 Cookie（写在 .env 的 COOKIES 变量中）。
"""

import json
import os
import time
from pathlib import Path

import openpyxl
from loguru import logger

from apis.xhs_pc_apis import XHS_Apis
from xhs_utils.common_util import load_env

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

USER_ID = "5d500c440000000011015466"

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

COLLECT_LIST_PATH = _PROJECT_ROOT / "datas" / "collect_list.json"
EXCEL_PATH = _PROJECT_ROOT / "datas" / "excel_datas" / "收藏列表.xlsx"
JSON_PATH = _PROJECT_ROOT / "datas" / "excel_datas" / "收藏列表.json"

# ---------------------------------------------------------------------------
# 列表缓存 I/O
# ---------------------------------------------------------------------------


def _load_cached_list() -> list[dict]:
    """读取本地缓存的收藏列表；文件不存在时返回空列表。"""
    if not COLLECT_LIST_PATH.exists():
        return []
    try:
        with open(COLLECT_LIST_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            logger.warning("collect_list.json 格式异常，重置为空列表")
            return []
        logger.info(f"已加载缓存列表: {len(data)} 条")
        return data
    except (json.JSONDecodeError, OSError) as exc:
        logger.error(f"读取缓存列表失败: {exc}")
        return []


def _save_cached_list(notes: list[dict]) -> None:
    """将笔记列表写入本地缓存文件。"""
    COLLECT_LIST_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(COLLECT_LIST_PATH, "w", encoding="utf-8") as f:
        json.dump(notes, f, ensure_ascii=False, indent=2)
    logger.info(f"缓存列表已保存: {COLLECT_LIST_PATH}（{len(notes)} 条）")


# ---------------------------------------------------------------------------
# API 拉取逻辑
# ---------------------------------------------------------------------------


def _fetch_full_list(xhs_apis: XHS_Apis, cookies_str: str) -> list[dict]:
    """全量拉取收藏列表（首次运行或缓存不存在时使用）。"""
    cursor = ""
    all_notes: list[dict] = []
    page = 0

    while True:
        page += 1
        success, msg, res_json = xhs_apis.get_user_collect_note_info(
            USER_ID, cursor, cookies_str, xsec_token="", xsec_source="pc_user"
        )
        if not success:
            logger.error(f"获取收藏列表第 {page} 页失败: {msg}")
            break

        notes: list[dict] = res_json["data"]["notes"]
        all_notes.extend(notes)
        logger.info(f"第 {page} 页: {len(notes)} 条，累计 {len(all_notes)} 条")

        has_more: bool = res_json["data"].get("has_more", False)
        if not has_more or "cursor" not in res_json["data"]:
            break

        cursor = str(res_json["data"]["cursor"])
        time.sleep(1)

    return all_notes


def _fetch_incremental(
    xhs_apis: XHS_Apis, cookies_str: str, known_ids: set[str]
) -> list[dict]:
    """
    增量拉取：遇到第一个已知 note_id 时停止。

    返回仅包含新笔记（不含缓存中已有的笔记）。
    """
    cursor = ""
    new_notes: list[dict] = []
    page = 0

    while True:
        page += 1
        success, msg, res_json = xhs_apis.get_user_collect_note_info(
            USER_ID, cursor, cookies_str, xsec_token="", xsec_source="pc_user"
        )
        if not success:
            logger.error(f"增量拉取第 {page} 页失败: {msg}")
            break

        notes: list[dict] = res_json["data"]["notes"]
        hit_known = False
        for note in notes:
            nid: str = note.get("note_id") or note.get("id", "")
            if nid in known_ids:
                hit_known = True
                break
            new_notes.append(note)

        if hit_known:
            logger.info(f"增量检测：第 {page} 页命中已知笔记，停止拉取")
            break

        has_more: bool = res_json["data"].get("has_more", False)
        if not has_more or "cursor" not in res_json["data"]:
            break

        cursor = str(res_json["data"]["cursor"])
        time.sleep(1)

    return new_notes


# ---------------------------------------------------------------------------
# 摘要导出
# ---------------------------------------------------------------------------

_EXCEL_HEADERS = ["笔记ID", "标题", "类型", "作者", "作者ID", "点赞数", "链接"]
_ROW_KEYS = ["note_id", "title", "type", "author", "author_id", "liked_count", "note_url"]


def _build_summary_rows(notes: list[dict]) -> list[dict]:
    """将原始 API 笔记列表转换为摘要行列表。"""
    rows: list[dict] = []
    for note in notes:
        user: dict = note.get("user") or {}
        interact: dict = note.get("interact_info") or {}
        note_id: str = note.get("note_id", "")
        rows.append(
            {
                "note_id": note_id,
                "title": note.get("display_title", ""),
                "type": "视频" if note.get("type") == "video" else "图集",
                "author": user.get("nickname", ""),
                "author_id": user.get("user_id", ""),
                "liked_count": interact.get("liked_count", ""),
                "note_url": f"https://www.xiaohongshu.com/explore/{note_id}",
            }
        )
    return rows


def _export_summary(notes: list[dict]) -> None:
    """将摘要写出到 Excel 和 JSON 文件。"""
    rows = _build_summary_rows(notes)

    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)

    # ── Excel ──────────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "收藏列表"
    ws.append(_EXCEL_HEADERS)
    for row in rows:
        ws.append([str(row.get(k, "")) for k in _ROW_KEYS])
    wb.save(EXCEL_PATH)
    logger.info(f"Excel 已导出: {EXCEL_PATH}（{len(rows)} 条）")

    # ── JSON ───────────────────────────────────────────────────────────────────
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logger.info(f"JSON 已导出: {JSON_PATH}（{len(rows)} 条）")


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------


def main(refresh: bool = False) -> None:
    """获取收藏列表，更新缓存，导出摘要。

    参数
    ----
    refresh : 强制重新全量拉取（刷新 xsec_token），忽略本地缓存。
    """
    cookies_str = load_env()
    if not cookies_str:
        logger.error("未找到 COOKIES 环境变量，请在 .env 中配置后重试")
        return

    xhs_apis = XHS_Apis()
    cached_list = [] if refresh else _load_cached_list()

    if refresh:
        logger.info("强制刷新模式，重新全量拉取收藏列表...")

    if not cached_list:
        all_notes = _fetch_full_list(xhs_apis, cookies_str)
        if not all_notes:
            logger.warning("未获取到任何笔记，请检查 Cookie 是否有效")
            return
        _save_cached_list(all_notes)
        logger.info(f"全量拉取完成，共 {len(all_notes)} 条")
    else:
        known_ids: set[str] = {
            n.get("note_id") or n.get("id", "") for n in cached_list
        }
        logger.info(f"增量检测中（已缓存 {len(cached_list)} 条）...")
        new_notes = _fetch_incremental(xhs_apis, cookies_str, known_ids)
        if new_notes:
            all_notes = new_notes + cached_list
            _save_cached_list(all_notes)
            logger.info(f"发现 {len(new_notes)} 条新收藏，列表更新至 {len(all_notes)} 条")
        else:
            all_notes = cached_list
            logger.info("没有新收藏")

    _export_summary(all_notes)
    logger.info("fetch_list 完成")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="获取小红书收藏列表")
    parser.add_argument("--refresh", action="store_true", help="强制重新全量拉取（刷新 xsec_token）")
    args = parser.parse_args()
    main(refresh=args.refresh)
