"""
爬取我的小红书收藏笔记，保存完整内容（Excel + JSON + 媒体文件）供分析使用。
支持断点续传、增量检测、本地去重。

用法：
  python crawl_my_collections.py          # 全量/增量爬取
  python crawl_my_collections.py --limit 5 # 只爬前 5 条（测试用）
"""
import argparse
import json
import os
import time
from pathlib import Path

import openpyxl
from loguru import logger

from apis.xhs_pc_apis import XHS_Apis
from xhs_utils.common_util import init
from xhs_utils.data_util import handle_note_info, download_note, save_to_xlsx

USER_ID = "5d500c440000000011015466"

REQUEST_DELAY = 2
PROGRESS_FILE = "datas/collection_progress.json"
EXPORT_BATCH_SIZE = 10

# ---------------------------------------------------------------------------
# Progress: 仅存 collect_list
# ---------------------------------------------------------------------------

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 兼容旧格式：只保留 collect_list
        collect_list = data.get("collect_list", [])
        logger.info(f"加载进度: 缓存列表 {len(collect_list)} 条")
        return {"collect_list": collect_list}
    return {"collect_list": []}


def save_progress(progress):
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False)


# ---------------------------------------------------------------------------
# 收藏列表获取
# ---------------------------------------------------------------------------

def fetch_full_list(xhs_apis, cookies_str):
    """首次运行：拉取完整收藏列表。"""
    cursor = ''
    all_notes = []
    page = 0
    while True:
        page += 1
        success, msg, res_json = xhs_apis.get_user_collect_note_info(
            USER_ID, cursor, cookies_str, xsec_token='', xsec_source='pc_user'
        )
        if not success:
            logger.error(f"获取收藏列表第 {page} 页失败: {msg}")
            break
        notes = res_json["data"]["notes"]
        all_notes.extend(notes)
        logger.info(f"收藏列表第 {page} 页: {len(notes)} 条，累计 {len(all_notes)} 条")
        if 'cursor' not in res_json["data"] or not res_json["data"].get("has_more"):
            break
        cursor = str(res_json["data"]["cursor"])
        time.sleep(1)
    return all_notes


def fetch_new_collections(xhs_apis, cookies_str, known_ids):
    """增量检测：拉新收藏，遇到已知 note_id 立即停止。"""
    cursor = ''
    new_notes = []
    page = 0
    while True:
        page += 1
        success, msg, res_json = xhs_apis.get_user_collect_note_info(
            USER_ID, cursor, cookies_str, xsec_token='', xsec_source='pc_user'
        )
        if not success:
            logger.error(f"获取新收藏第 {page} 页失败: {msg}")
            break
        notes = res_json["data"]["notes"]
        hit_known = False
        for n in notes:
            nid = n.get("note_id") or n.get("id", "")
            if nid in known_ids:
                hit_known = True
                break
            new_notes.append(n)
        if hit_known or 'cursor' not in res_json["data"] or not res_json["data"].get("has_more"):
            break
        cursor = str(res_json["data"]["cursor"])
        time.sleep(1)
    return new_notes


# ---------------------------------------------------------------------------
# 收藏列表摘要导出
# ---------------------------------------------------------------------------

def export_collect_list(cached_list, base_path):
    rows = []
    for n in cached_list:
        user = n.get("user") or {}
        interact = n.get("interact_info") or {}
        rows.append({
            "note_id": n.get("note_id", ""),
            "title": n.get("display_title", ""),
            "type": "视频" if n.get("type") == "video" else "图集",
            "author": user.get("nickname", ""),
            "author_id": user.get("user_id", ""),
            "liked_count": interact.get("liked_count", ""),
            "note_url": f"https://www.xiaohongshu.com/explore/{n.get('note_id', '')}",
        })
    excel_path = os.path.join(base_path["excel"], "收藏列表.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["笔记ID", "标题", "类型", "作者", "作者ID", "点赞数", "链接"]
    keys = ["note_id", "title", "type", "author", "author_id", "liked_count", "note_url"]
    ws.append(headers)
    for row in rows:
        ws.append([str(row.get(k, "")) for k in keys])
    wb.save(excel_path)
    json_path = os.path.join(base_path["excel"], "收藏列表.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logger.info(f"收藏列表已导出: {excel_path}（{len(rows)} 条）")


# ---------------------------------------------------------------------------
# 本地状态扫描
# ---------------------------------------------------------------------------

def scan_local_state(media_path):
    """扫描 media_datas/，返回 (has_details, media_complete)。

    has_details: { note_id: note_info } — 有 info.json 的笔记
    media_complete: { note_id } — 媒体文件完整的笔记
    """
    has_details = {}
    media_complete = set()
    media_dir = Path(media_path)
    if not media_dir.exists():
        return has_details, media_complete

    for user_dir in media_dir.iterdir():
        if not user_dir.is_dir():
            continue
        for note_dir in user_dir.iterdir():
            if not note_dir.is_dir():
                continue
            parts = note_dir.name.rsplit('_', 1)
            if len(parts) != 2:
                continue
            note_id = parts[1]

            info_path = note_dir / 'info.json'
            if not info_path.exists():
                continue
            try:
                with open(info_path, 'r', encoding='utf-8') as f:
                    note_info = json.loads(f.read().strip())
                has_details[note_id] = note_info
            except (json.JSONDecodeError, OSError):
                continue

            note_type = note_info.get('note_type', '')
            if note_type == '图集':
                expected = len(note_info.get('image_list', []))
                actual = len(list(note_dir.glob('image_*.jpg')))
                if expected > 0 and actual >= expected:
                    media_complete.add(note_id)
            elif note_type == '视频':
                if (note_dir / 'video.mp4').exists():
                    media_complete.add(note_id)

    return has_details, media_complete


# ---------------------------------------------------------------------------
# 详情聚合导出
# ---------------------------------------------------------------------------

def export_details(all_details, base_path):
    if not all_details:
        return
    excel_path = os.path.join(base_path["excel"], "我的收藏.xlsx")
    save_to_xlsx(all_details, excel_path)
    json_path = os.path.join(base_path["excel"], "我的收藏.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_details, f, ensure_ascii=False, indent=2)
    logger.info(f"详情已导出: {excel_path}（{len(all_details)} 条）")


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def crawl_collections(limit=None):
    cookies_str, base_path = init()
    xhs_apis = XHS_Apis()
    progress = load_progress()
    cached_list = progress.get("collect_list", [])

    if limit:
        logger.info(f"测试模式: 只爬取前 {limit} 条")

    # ---- 获取/更新收藏列表 ----
    if cached_list:
        cached_ids = {n.get("note_id") or n.get("id", "") for n in cached_list}
        logger.info("检测新收藏...")
        new_notes = fetch_new_collections(xhs_apis, cookies_str, cached_ids)
        if new_notes:
            cached_list = new_notes + cached_list
            progress["collect_list"] = cached_list
            save_progress(progress)
            logger.info(f"发现 {len(new_notes)} 条新收藏，列表更新至 {len(cached_list)} 条")
            export_collect_list(cached_list, base_path)
        else:
            logger.info("没有新收藏")
    else:
        logger.info("首次运行，获取完整收藏列表...")
        cached_list = fetch_full_list(xhs_apis, cookies_str)
        if cached_list:
            progress["collect_list"] = cached_list
            save_progress(progress)
            logger.info(f"收藏列表已缓存: {len(cached_list)} 条")
            export_collect_list(cached_list, base_path)

    # ---- 扫描本地状态 ----
    logger.info("扫描本地文件...")
    has_details, media_complete = scan_local_state(base_path["media"])
    complete_ids = set(has_details.keys()) & media_complete
    logger.info(f"本地状态: {len(has_details)} 条有详情，{len(media_complete)} 条媒体完整，{len(complete_ids)} 条完全完成")

    # ---- 构建待处理列表 ----
    collect_list = [n for n in cached_list if (n.get("note_id") or n.get("id", "")) not in complete_ids]
    if limit:
        collect_list = collect_list[:limit]
    if not collect_list:
        logger.info("没有需要处理的笔记")
        if has_details:
            export_details(list(has_details.values()), base_path)
        return

    total = len(collect_list)
    logger.info(f"待处理: {total} 条")

    # ---- 逐条处理 ----
    all_details = dict(has_details)
    failed = []
    consecutive_fails = 0
    new_success = 0

    for i, note in enumerate(collect_list):
        note_id = note.get("note_id") or note.get("id", "")
        xsec_token = note.get("xsec_token", "")
        note_url = f"https://www.xiaohongshu.com/explore/{note_id}?xsec_token={xsec_token}&xsec_source=pc_user"

        logger.info(f"[{i + 1}/{total}] 笔记 {note_id}")
        note_ok = False
        try:
            if note_id in has_details:
                note_info = has_details[note_id]
                logger.info(f"  -> 详情已有，仅补下载媒体: {note_info['title'][:30]}")
            else:
                success, msg, res_json = xhs_apis.get_note_info(note_url, cookies_str)
                if not success or not res_json:
                    logger.warning(f"  -> API 失败: {msg}")
                    failed.append({"note_id": note_id, "msg": str(msg)})
                    raise ValueError(msg)
                items = (res_json.get("data") or {}).get("items")
                if not items:
                    logger.warning(f"  -> 笔记已删除或不可见")
                    failed.append({"note_id": note_id, "msg": "no items"})
                    raise ValueError("no items")
                note_info = items[0]
                note_info["url"] = note_url
                note_info = handle_note_info(note_info)

            download_note(note_info, base_path["media"], "all")
            all_details[note_id] = note_info
            note_ok = True
            new_success += 1
            consecutive_fails = 0
            logger.info(f"  -> 完成: {note_info['title'][:30]}")
        except Exception as e:
            if not failed or failed[-1].get("note_id") != note_id:
                failed.append({"note_id": note_id, "msg": str(e)})

        if not note_ok:
            consecutive_fails += 1
            if consecutive_fails >= 5:
                logger.error(f"连续失败 {consecutive_fails} 次，暂停 60 秒...")
                time.sleep(60)
                consecutive_fails = 0

        if (i + 1) % EXPORT_BATCH_SIZE == 0:
            export_details(list(all_details.values()), base_path)

        time.sleep(REQUEST_DELAY)

    # ---- 最终导出 ----
    export_details(list(all_details.values()), base_path)
    logger.info(f"本次: 新增 {new_success}/{total}，失败 {len(failed)} | 总计: {len(all_details)} 条详情")
    if failed:
        logger.warning(f"失败列表: {json.dumps(failed, ensure_ascii=False)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="爬取小红书收藏笔记")
    parser.add_argument("--limit", type=int, default=None, help="只爬取前 N 条（测试用）")
    args = parser.parse_args()
    crawl_collections(limit=args.limit)
