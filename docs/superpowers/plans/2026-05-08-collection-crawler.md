# Collection Crawler Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a three-script pipeline to crawl Xiaohongshu bookmarked notes — fetch list (Spider_XHS), download details+media (XHS-Downloader), export to Excel/JSON — plus a one-time migration of existing data.

**Architecture:** Three independent scripts (`scripts/fetch_list.py`, `scripts/download.py`, `scripts/export.py`) communicate via files: `collect_list.json` → XHS-Downloader SQLite DBs → Excel/JSON. A wrapper `scripts/run_all.py` chains them. A one-time `scripts/migrate_old_data.py` brings 591 existing notes into the target format.

**Tech Stack:** Python 3.10+, Spider_XHS APIs, XHS-Downloader Python API (`source.application.app.XHS`), aiosqlite, openpyxl, loguru.

**Spec:** `docs/superpowers/specs/2026-05-08-collection-crawler-design.md`

**XHS-Downloader location:** `../XHS-Downloader` (sibling directory to Spider_XHS)

---

## File Structure

```
scripts/
├── fetch_list.py          # Step 1: collection list fetching
├── download.py            # Step 2: detail + media download
├── export.py              # Step 3: aggregate export
├── run_all.py             # Wrapper
└── migrate_old_data.py    # One-time migration
```

All scripts run from the Spider_XHS project root: `python -m scripts.fetch_list`, etc.

Data outputs:
- `datas/collect_list.json` — cached collection list
- `datas/excel_datas/收藏列表.xlsx` + `.json` — list summary
- `datas/xhs_download/ExploreID.db` — downloaded IDs (at `work_path` root)
- `datas/xhs_download/Download/ExploreData.db` — metadata (at `folder_name` level)
- `datas/xhs_download/Download/{author}/{title}_{note_id}/` — media + info.json
- `datas/excel_datas/我的收藏.xlsx` + `.json` — analysis export

---

### Task 1: Create scripts directory and `__init__.py`

**Files:**
- Create: `scripts/__init__.py`

- [ ] **Step 1: Create the package**

```bash
mkdir -p scripts
touch scripts/__init__.py
```

- [ ] **Step 2: Commit**

```bash
git add scripts/__init__.py
git commit -m "chore: create scripts package"
```

---

### Task 2: Implement `scripts/export.py`

Starting with export because it has no dependencies on the other scripts, and download.py will import from it.

**Files:**
- Create: `scripts/export.py`

- [ ] **Step 1: Write export.py**

```python
"""
Step 3: Export all note metadata from ExploreData.db to Excel/JSON.
Stateless and idempotent — can be run anytime.

Usage: python -m scripts.export
"""
import json
import os
import sqlite3

import openpyxl
from loguru import logger

DB_PATH = "datas/xhs_download/Download/ExploreData.db"
EXCEL_PATH = "datas/excel_datas/我的收藏.xlsx"
JSON_PATH = "datas/excel_datas/我的收藏.json"

COLUMNS = [
    "采集时间", "作品ID", "作品类型", "作品标题", "作品描述", "作品标签",
    "发布时间", "最后更新时间", "收藏数量", "评论数量", "分享数量", "点赞数量",
    "作者昵称", "作者ID", "作者链接", "作品链接", "下载地址", "动图地址",
]

EXCEL_HEADERS = [
    "作品ID", "标题", "类型", "作者", "描述", "标签",
    "发布时间", "收藏数", "评论数", "分享数", "点赞数", "链接",
]

EXCEL_KEYS = [
    "作品ID", "作品标题", "作品类型", "作者昵称", "作品描述", "作品标签",
    "发布时间", "收藏数量", "评论数量", "分享数量", "点赞数量", "作品链接",
]


def export_from_db(db_path=DB_PATH, excel_path=EXCEL_PATH, json_path=JSON_PATH):
    """Read ExploreData.db and write Excel + JSON."""
    if not os.path.exists(db_path):
        logger.warning(f"数据库不存在: {db_path}")
        return 0

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM explore_data").fetchall()
    conn.close()

    if not rows:
        logger.info("数据库为空，跳过导出")
        return 0

    records = [dict(r) for r in rows]

    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(EXCEL_HEADERS)
    for rec in records:
        ws.append([str(rec.get(k, "")) for k in EXCEL_KEYS])
    wb.save(excel_path)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    logger.info(f"已导出 {len(records)} 条 → {excel_path}")
    return len(records)


def main():
    export_from_db()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify it handles missing DB gracefully**

```bash
python -m scripts.export
```

Expected: logs "数据库不存在" (since ExploreData.db doesn't exist yet), no crash.

- [ ] **Step 3: Commit**

```bash
git add scripts/export.py
git commit -m "feat: add export.py — aggregate ExploreData.db to Excel/JSON"
```

---

### Task 3: Implement `scripts/fetch_list.py`

**Files:**
- Create: `scripts/fetch_list.py`

- [ ] **Step 1: Write fetch_list.py**

```python
"""
Step 1: Fetch collection list using Spider_XHS API (requires Cookie).
Supports incremental detection — only fetches new bookmarks.

Usage: python -m scripts.fetch_list
"""
import json
import os
import time

import openpyxl
from loguru import logger

from apis.xhs_pc_apis import XHS_Apis
from xhs_utils.common_util import load_env

USER_ID = "5d500c440000000011015466"
COLLECT_LIST_PATH = "datas/collect_list.json"
EXCEL_PATH = "datas/excel_datas/收藏列表.xlsx"
JSON_PATH = "datas/excel_datas/收藏列表.json"


def load_collect_list():
    if os.path.exists(COLLECT_LIST_PATH):
        with open(COLLECT_LIST_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_collect_list(data):
    os.makedirs(os.path.dirname(COLLECT_LIST_PATH), exist_ok=True)
    with open(COLLECT_LIST_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def fetch_full_list(xhs_apis, cookies_str):
    """Fetch all pages of the collection list."""
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
    """Fetch only new bookmarks. Stop at first known note_id."""
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


def export_summary(cached_list):
    """Export collection list summary to Excel + JSON."""
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
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ["笔记ID", "标题", "类型", "作者", "作者ID", "点赞数", "链接"]
    keys = ["note_id", "title", "type", "author", "author_id", "liked_count", "note_url"]
    ws.append(headers)
    for row in rows:
        ws.append([str(row.get(k, "")) for k in keys])
    wb.save(EXCEL_PATH)
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logger.info(f"收藏列表已导出: {EXCEL_PATH}（{len(rows)} 条）")


def main():
    cookies_str = load_env()
    if not cookies_str:
        logger.error("未找到 Cookie，请在 .env 中设置 COOKIES")
        return
    xhs_apis = XHS_Apis()
    cached_list = load_collect_list()

    if cached_list:
        cached_ids = {n.get("note_id") or n.get("id", "") for n in cached_list}
        logger.info(f"已有缓存 {len(cached_list)} 条，检测新收藏...")
        new_notes = fetch_new_collections(xhs_apis, cookies_str, cached_ids)
        if new_notes:
            cached_list = new_notes + cached_list
            save_collect_list(cached_list)
            logger.info(f"发现 {len(new_notes)} 条新收藏，总计 {len(cached_list)} 条")
        else:
            logger.info("没有新收藏")
    else:
        logger.info("首次运行，获取完整收藏列表...")
        cached_list = fetch_full_list(xhs_apis, cookies_str)
        if cached_list:
            save_collect_list(cached_list)
            logger.info(f"收藏列表已缓存: {len(cached_list)} 条")

    if cached_list:
        export_summary(cached_list)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test with existing cached data**

Copy current `datas/collection_progress.json`'s collect_list to the new location:

```bash
python -c "
import json
with open('datas/collection_progress.json') as f:
    d = json.load(f)
with open('datas/collect_list.json', 'w') as f:
    json.dump(d.get('collect_list', []), f, ensure_ascii=False)
print(f'Migrated {len(d.get(\"collect_list\", []))} items')
"
```

Then run:

```bash
python -m scripts.fetch_list
```

Expected: "已有缓存 1699 条，检测新收藏..." → either "没有新收藏" or finds new ones. Exports `收藏列表.xlsx/json`.

- [ ] **Step 3: Commit**

```bash
git add scripts/fetch_list.py
git commit -m "feat: add fetch_list.py — collection list fetching with incremental detection"
```

---

### Task 4: Implement `scripts/download.py`

**Files:**
- Create: `scripts/download.py`

- [ ] **Step 1: Write download.py**

```python
"""
Step 2: Download note details + media using XHS-Downloader (no Cookie needed).
Reads collect_list.json, skips already-downloaded notes via ExploreID.db.

Usage:
  python -m scripts.download              # process all
  python -m scripts.download --limit 5    # process first 5 unfinished
"""
import argparse
import asyncio
import json
import os
import sqlite3
import sys
import time
from pathlib import Path

from loguru import logger

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "XHS-Downloader"))
from source.application.app import XHS

from scripts.export import export_from_db

COLLECT_LIST_PATH = "datas/collect_list.json"
XHS_WORK_PATH = "datas/xhs_download"
XHS_FOLDER_NAME = "Download"
EXPLORE_ID_DB = os.path.join(XHS_WORK_PATH, "ExploreID.db")
EXPLORE_DATA_DB = os.path.join(XHS_WORK_PATH, XHS_FOLDER_NAME, "ExploreData.db")
EXPORT_BATCH_SIZE = 10


def load_collect_list():
    if not os.path.exists(COLLECT_LIST_PATH):
        logger.error(f"收藏列表不存在: {COLLECT_LIST_PATH}，请先运行 fetch_list.py")
        return []
    with open(COLLECT_LIST_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_downloaded_ids():
    """Read ExploreID.db to get already-downloaded note IDs."""
    if not os.path.exists(EXPLORE_ID_DB):
        return set()
    conn = sqlite3.connect(EXPLORE_ID_DB)
    ids = {r[0] for r in conn.execute("SELECT ID FROM explore_id").fetchall()}
    conn.close()
    return ids


def find_note_dir(note_id):
    """Find the download directory for a note by searching for *_{note_id}."""
    download_root = Path(XHS_WORK_PATH) / XHS_FOLDER_NAME
    if not download_root.exists():
        return None
    matches = list(download_root.rglob(f"*_{note_id}"))
    dirs = [m for m in matches if m.is_dir()]
    return dirs[0] if dirs else None


def convert_to_info_json(result, note_url):
    """Convert XHS-Downloader result dict to info.json format."""
    note_type = "图集" if result.get("作品类型") in ("图文", "图集") else "视频"
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


def write_info_json(note_info, note_dir):
    """Write info.json and detail.txt into the note's download directory."""
    info_path = Path(note_dir) / "info.json"
    with open(info_path, "w", encoding="utf-8") as f:
        json.dump(note_info, f, ensure_ascii=False, indent=2)

    detail_path = Path(note_dir) / "detail.txt"
    with open(detail_path, "w", encoding="utf-8") as f:
        for key, label in [
            ("note_id", "笔记ID"),
            ("note_url", "笔记链接"),
            ("note_type", "笔记类型"),
            ("nickname", "作者"),
            ("title", "标题"),
            ("desc", "描述"),
            ("liked_count", "点赞数"),
            ("collected_count", "收藏数"),
            ("comment_count", "评论数"),
            ("share_count", "分享数"),
            ("tags", "标签"),
            ("upload_time", "发布时间"),
        ]:
            f.write(f"{label}: {note_info.get(key, '')}\n")


async def process_notes(collect_list, downloaded_ids, limit=None):
    """Main download loop using XHS-Downloader."""
    to_process = []
    for note in collect_list:
        note_id = note.get("note_id") or note.get("id", "")
        if note_id not in downloaded_ids:
            to_process.append(note)

    if limit:
        to_process = to_process[:limit]

    if not to_process:
        logger.info("没有需要下载的笔记")
        return 0, 0

    total = len(to_process)
    logger.info(f"待下载: {total} 条（已跳过 {len(collect_list) - total} 条已完成）")

    async with XHS(
        work_path=XHS_WORK_PATH,
        folder_name=XHS_FOLDER_NAME,
        name_format="作品标题 作品ID",
        author_archive=True,
        folder_mode=True,
        record_data=True,
        download_record=True,
        image_format="JPEG",
        image_download=True,
        video_download=True,
    ) as xhs:
        success_count = 0
        fail_count = 0
        consecutive_fails = 0

        for i, note in enumerate(to_process):
            note_id = note.get("note_id") or note.get("id", "")
            xsec_token = note.get("xsec_token", "")
            note_url = f"https://www.xiaohongshu.com/discovery/item/{note_id}?xsec_token={xsec_token}&xsec_source=pc_share"

            logger.info(f"[{i + 1}/{total}] 笔记 {note_id}")
            try:
                result = await xhs.extract(url=note_url, download=True)
                if result and result[0]:
                    note_info = convert_to_info_json(result[0], note_url)
                    note_dir = find_note_dir(note_id)
                    if note_dir:
                        write_info_json(note_info, note_dir)
                    success_count += 1
                    consecutive_fails = 0
                    logger.info(f"  -> 完成: {note_info['title'][:30]}")
                else:
                    fail_count += 1
                    consecutive_fails += 1
                    logger.warning(f"  -> 返回空结果")
            except Exception as e:
                fail_count += 1
                consecutive_fails += 1
                logger.warning(f"  -> 异常: {e}")

            if consecutive_fails >= 5:
                logger.error(f"连续失败 {consecutive_fails} 次，暂停 60 秒...")
                await asyncio.sleep(60)
                consecutive_fails = 0

            if (i + 1) % EXPORT_BATCH_SIZE == 0:
                export_from_db(EXPLORE_DATA_DB)

    return success_count, fail_count


def main(limit=None):
    collect_list = load_collect_list()
    if not collect_list:
        return
    downloaded_ids = load_downloaded_ids()
    logger.info(f"收藏列表: {len(collect_list)} 条，已下载: {len(downloaded_ids)} 条")

    success, fail = asyncio.run(process_notes(collect_list, downloaded_ids, limit))

    export_from_db(EXPLORE_DATA_DB)
    logger.info(f"完成！成功 {success}，失败 {fail}，总计 {len(load_downloaded_ids())} 条已下载")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="下载收藏笔记详情和媒体")
    parser.add_argument("--limit", type=int, default=None, help="只处理前 N 条未完成的")
    args = parser.parse_args()
    main(limit=args.limit)
```

- [ ] **Step 2: Test with --limit 2**

```bash
python -m scripts.download --limit 2
```

Expected:
- Downloads 2 notes via XHS-Downloader
- Creates `datas/xhs_download/Download/{author}/{title}_{note_id}/` directories
- Each directory contains media files + `info.json` + `detail.txt`
- `ExploreID.db` and `ExploreData.db` are created/updated
- `我的收藏.xlsx/json` are exported

- [ ] **Step 3: Test breakpoint — run again**

```bash
python -m scripts.download --limit 2
```

Expected: "已跳过 2 条已完成" (or more if migration ran), processes next 2.

- [ ] **Step 4: Verify directory structure**

```bash
find datas/xhs_download/Download -type f | head -20
```

Expected: files organized as `{author}/{title}_{note_id}/{filename}.jpeg` with `info.json` and `detail.txt` alongside media.

- [ ] **Step 5: Commit**

```bash
git add scripts/download.py
git commit -m "feat: add download.py — XHS-Downloader based detail+media downloading"
```

---

### Task 5: Implement `scripts/run_all.py`

**Files:**
- Create: `scripts/run_all.py`

- [ ] **Step 1: Write run_all.py**

```python
"""
Wrapper: runs fetch_list → download → export in sequence.

Usage:
  python -m scripts.run_all              # full pipeline
  python -m scripts.run_all --limit 10   # limit download step
"""
import argparse

from loguru import logger

from scripts.fetch_list import main as fetch_list_main
from scripts.download import main as download_main
from scripts.export import main as export_main


def main():
    parser = argparse.ArgumentParser(description="收藏笔记爬取完整流程")
    parser.add_argument("--limit", type=int, default=None, help="下载步骤只处理前 N 条")
    args = parser.parse_args()

    logger.info("=== 步骤 1: 获取收藏列表 ===")
    fetch_list_main()

    logger.info("=== 步骤 2: 下载详情和媒体 ===")
    download_main(limit=args.limit)

    logger.info("=== 步骤 3: 导出汇总 ===")
    export_main()

    logger.info("=== 全部完成 ===")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test**

```bash
python -m scripts.run_all --limit 1
```

Expected: runs all three steps in sequence, processes 1 note.

- [ ] **Step 3: Commit**

```bash
git add scripts/run_all.py
git commit -m "feat: add run_all.py — wrapper for full pipeline"
```

---

### Task 6: Implement `scripts/migrate_old_data.py`

**Files:**
- Create: `scripts/migrate_old_data.py`

- [ ] **Step 1: Write migrate_old_data.py**

```python
"""
One-time migration: insert 591 Spider_XHS info.json records into
ExploreID.db and ExploreData.db so that download.py and export.py
only deal with one data source.

Usage: python -m scripts.migrate_old_data
"""
import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from loguru import logger

MEDIA_DATAS_PATH = "datas/media_datas"
XHS_WORK_PATH = "datas/xhs_download"
EXPLORE_ID_DB = os.path.join(XHS_WORK_PATH, "ExploreID.db")
EXPLORE_DATA_DB = os.path.join(XHS_WORK_PATH, "Download", "ExploreData.db")

DATA_COLUMNS = [
    "采集时间", "作品ID", "作品类型", "作品标题", "作品描述", "作品标签",
    "发布时间", "最后更新时间", "收藏数量", "评论数量", "分享数量", "点赞数量",
    "作者昵称", "作者ID", "作者链接", "作品链接", "下载地址", "动图地址",
]


def scan_info_jsons():
    """Scan all info.json files in Spider_XHS media_datas/."""
    results = []
    media_dir = Path(MEDIA_DATAS_PATH)
    if not media_dir.exists():
        return results
    for user_dir in media_dir.iterdir():
        if not user_dir.is_dir():
            continue
        for note_dir in user_dir.iterdir():
            if not note_dir.is_dir():
                continue
            info_path = note_dir / "info.json"
            if not info_path.exists():
                continue
            try:
                with open(info_path, "r", encoding="utf-8") as f:
                    data = json.loads(f.read().strip())
                results.append(data)
            except (json.JSONDecodeError, OSError) as e:
                logger.warning(f"跳过 {info_path}: {e}")
    return results


def convert_to_explore_data(note_info):
    """Convert Spider_XHS info.json to ExploreData.db row."""
    tags = note_info.get("tags", [])
    tags_str = " ".join(tags) if isinstance(tags, list) else str(tags)

    image_list = note_info.get("image_list", [])
    video_addr = note_info.get("video_addr")
    if note_info.get("note_type") == "视频" and video_addr:
        download_urls = json.dumps([video_addr], ensure_ascii=False)
    else:
        download_urls = json.dumps(image_list, ensure_ascii=False)

    note_type_map = {"图集": "图文", "视频": "视频"}
    note_type = note_type_map.get(note_info.get("note_type", ""), note_info.get("note_type", ""))

    return {
        "采集时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "作品ID": note_info.get("note_id", ""),
        "作品类型": note_type,
        "作品标题": note_info.get("title", ""),
        "作品描述": note_info.get("desc", ""),
        "作品标签": tags_str,
        "发布时间": note_info.get("upload_time", ""),
        "最后更新时间": note_info.get("upload_time", ""),
        "收藏数量": str(note_info.get("collected_count", "")),
        "评论数量": str(note_info.get("comment_count", "")),
        "分享数量": str(note_info.get("share_count", "")),
        "点赞数量": str(note_info.get("liked_count", "")),
        "作者昵称": note_info.get("nickname", ""),
        "作者ID": note_info.get("user_id", ""),
        "作者链接": note_info.get("home_url", ""),
        "作品链接": note_info.get("note_url", ""),
        "下载地址": download_urls,
        "动图地址": "",
    }


def main():
    info_list = scan_info_jsons()
    if not info_list:
        logger.info("没有找到 info.json 文件")
        return

    logger.info(f"找到 {len(info_list)} 条 info.json")

    os.makedirs(os.path.dirname(EXPLORE_ID_DB), exist_ok=True)
    os.makedirs(os.path.dirname(EXPLORE_DATA_DB), exist_ok=True)

    id_conn = sqlite3.connect(EXPLORE_ID_DB)
    id_conn.execute("CREATE TABLE IF NOT EXISTS explore_id (ID TEXT PRIMARY KEY)")

    col_defs = ", ".join(f"{name} {typ}" for name, typ in zip(
        DATA_COLUMNS,
        ["TEXT"] * len(DATA_COLUMNS),
    ))
    data_conn = sqlite3.connect(EXPLORE_DATA_DB)
    data_conn.execute(f"CREATE TABLE IF NOT EXISTS explore_data ({col_defs})")

    # Fix PRIMARY KEY: 作品ID is the second column
    # The schema from XHS-Downloader uses "作品ID TEXT PRIMARY KEY"
    # Re-create with correct schema if needed
    data_conn.close()
    data_conn = sqlite3.connect(EXPLORE_DATA_DB)
    data_conn.execute("""CREATE TABLE IF NOT EXISTS explore_data (
        采集时间 TEXT,
        作品ID TEXT PRIMARY KEY,
        作品类型 TEXT,
        作品标题 TEXT,
        作品描述 TEXT,
        作品标签 TEXT,
        发布时间 TEXT,
        最后更新时间 TEXT,
        收藏数量 TEXT,
        评论数量 TEXT,
        分享数量 TEXT,
        点赞数量 TEXT,
        作者昵称 TEXT,
        作者ID TEXT,
        作者链接 TEXT,
        作品链接 TEXT,
        下载地址 TEXT,
        动图地址 TEXT
    )""")

    migrated = 0
    for note_info in info_list:
        note_id = note_info.get("note_id", "")
        if not note_id:
            continue

        id_conn.execute("REPLACE INTO explore_id VALUES (?)", (note_id,))
        row = convert_to_explore_data(note_info)
        placeholders = ", ".join("?" for _ in DATA_COLUMNS)
        values = tuple(row[col] for col in DATA_COLUMNS)
        data_conn.execute(
            f"REPLACE INTO explore_data ({', '.join(DATA_COLUMNS)}) VALUES ({placeholders})",
            values,
        )
        migrated += 1

    id_conn.commit()
    id_conn.close()
    data_conn.commit()
    data_conn.close()

    logger.info(f"迁移完成: {migrated} 条 → ExploreID.db + ExploreData.db")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run migration**

```bash
python -m scripts.migrate_old_data
```

Expected: "找到 591 条 info.json" → "迁移完成: 591 条"

- [ ] **Step 3: Verify migration results**

```bash
python -c "
import sqlite3
id_conn = sqlite3.connect('datas/xhs_download/ExploreID.db')
print(f'ExploreID.db: {id_conn.execute(\"SELECT COUNT(*) FROM explore_id\").fetchone()[0]} 条')
id_conn.close()
data_conn = sqlite3.connect('datas/xhs_download/Download/ExploreData.db')
print(f'ExploreData.db: {data_conn.execute(\"SELECT COUNT(*) FROM explore_data\").fetchone()[0]} 条')
data_conn.close()
"
```

Expected: both show 591 (or 591 + 161 if ExploreID.db already had XHS-Downloader entries).

- [ ] **Step 4: Test export reads migrated data**

```bash
python -m scripts.export
```

Expected: "已导出 591 条" (or more) to Excel/JSON.

- [ ] **Step 5: Commit**

```bash
git add scripts/migrate_old_data.py
git commit -m "feat: add migrate_old_data.py — one-time Spider_XHS to XHS-Downloader migration"
```

---

### Task 7: Clean up old `crawl_my_collections.py`

**Files:**
- Modify: `crawl_my_collections.py`

- [ ] **Step 1: Replace with redirect notice**

```python
"""
DEPRECATED: This script has been replaced by the scripts/ pipeline.

Usage:
  python -m scripts.fetch_list        # Step 1: fetch collection list
  python -m scripts.download           # Step 2: download details + media
  python -m scripts.export             # Step 3: export to Excel/JSON
  python -m scripts.run_all            # Run all steps
  python -m scripts.migrate_old_data   # One-time: migrate old data

See docs/superpowers/specs/2026-05-08-collection-crawler-design.md for details.
"""
print(__doc__)
```

- [ ] **Step 2: Commit**

```bash
git add crawl_my_collections.py
git commit -m "refactor: deprecate crawl_my_collections.py in favor of scripts/ pipeline"
```

---

### Task 8: End-to-end verification

- [ ] **Step 1: Run migration**

```bash
python -m scripts.migrate_old_data
```

Verify: ExploreID.db has 591+ entries, ExploreData.db has 591+ entries.

- [ ] **Step 2: Run full pipeline with limit**

```bash
python -m scripts.run_all --limit 3
```

Verify:
- fetch_list detects cached list, checks for new bookmarks
- download processes 3 new notes
- Each note has `{author}/{title}_{note_id}/` with media + info.json + detail.txt
- Excel/JSON updated with all data (591 migrated + 3 new)

- [ ] **Step 3: Run again to verify breakpoint**

```bash
python -m scripts.run_all --limit 3
```

Verify: skips the 3 just downloaded, processes next 3.

- [ ] **Step 4: Run export standalone**

```bash
python -m scripts.export
```

Verify: Excel/JSON reflect total count (591 + 6).

- [ ] **Step 5: Push to fork**

```bash
git push origin master
```
