# Collection Crawler Design Spec

## Overview

Crawl all bookmarked (collected) notes from Xiaohongshu, saving metadata and media files locally for analysis. The system splits into three independent scripts plus a one-time migration, using Spider_XHS for collection list fetching (requires Cookie) and XHS-Downloader for detail retrieval and media downloading (no Cookie, avoids account-level rate limiting).

## Architecture

```
scripts/
├── fetch_list.py          # Step 1: Fetch collection list (Cookie required)
├── download.py            # Step 2: Download details + media (no Cookie)
├── export.py              # Step 3: Aggregate export to Excel/JSON (offline)
├── run_all.py             # Wrapper: runs all three steps
└── migrate_old_data.py    # One-time: migrate Spider_XHS data to target format
```

### Data Flow

```
fetch_list.py                  download.py                     export.py
┌──────────────┐          ┌───────────────────┐          ┌──────────────┐
│ Spider_XHS   │          │ XHS-Downloader    │          │ Pure offline │
│ Cookie req'd │          │ No Cookie         │          │              │
│              │          │                   │          │              │
│ Output:      │──────────│ Input:            │──────────│ Input:       │
│ collect_list │          │ collect_list.json  │          │ ExploreData  │
│ .json        │          │                   │          │ .db          │
│              │          │ Output:           │          │              │
│ 收藏列表     │          │ ExploreID.db      │          │ Output:      │
│ .xlsx/.json  │          │ ExploreData.db    │          │ 我的收藏     │
│              │          │ Download/{author}/ │          │ .xlsx/.json  │
│              │          │  {title}_{id}/    │          │              │
│              │          │   *.jpeg/*.mp4    │          │              │
│              │          │   info.json       │          │              │
│              │          │                   │          │              │
│              │          │ After each note:  │          │              │
│              │          │ calls export()    │──────────│              │
└──────────────┘          └───────────────────┘          └──────────────┘
```

## File Locations

All paths relative to Spider_XHS project root.

| File | Purpose | Created by |
|------|---------|-----------|
| `datas/collect_list.json` | Cached collection list (note_id + xsec_token + summary) | fetch_list.py |
| `datas/excel_datas/收藏列表.xlsx` | Collection list summary for quick viewing | fetch_list.py |
| `datas/excel_datas/收藏列表.json` | Same as above, JSON format | fetch_list.py |
| `datas/xhs_download/ExploreID.db` | Downloaded note IDs (breakpoint tracking) | download.py / XHS-Downloader |
| `datas/xhs_download/ExploreData.db` | Note metadata (all fields) | download.py / XHS-Downloader |
| `datas/xhs_download/Download/{author}/{title}_{note_id}/` | Per-note directory with media + info.json | download.py |
| `datas/excel_datas/我的收藏.xlsx` | Aggregated analysis export | export.py |
| `datas/excel_datas/我的收藏.json` | Same as above, JSON format | export.py |

## Script Details

### fetch_list.py

**Responsibility**: Fetch collection list using Spider_XHS API. Only step that requires Cookie.

**Input**: `.env` file with `COOKIES` value.

**Output**: `datas/collect_list.json` — raw API response for each note in the collection.

**Logic**:
1. If `collect_list.json` does not exist: fetch all pages from API, save full list.
2. If `collect_list.json` exists: fetch incrementally from API (stop at first known note_id), prepend new notes to list.
3. Export `收藏列表.xlsx/json` (summary with note_id, title, author, liked_count, type).

**Breakpoint**: `collect_list.json` itself. Incremental detection avoids re-fetching.

**Parameters**: None.

### download.py

**Responsibility**: For each note in `collect_list.json`, use XHS-Downloader Python API to fetch details and download media. No Cookie required.

**Input**: `datas/collect_list.json`

**Output**:
- `ExploreID.db`: tracks downloaded note IDs (XHS-Downloader built-in)
- `ExploreData.db`: stores note metadata (XHS-Downloader built-in, `record_data=True`)
- `Download/{author}/{title}_{note_id}/`: per-note directory with media files
- `Download/{author}/{title}_{note_id}/info.json`: note metadata written by download.py after XHS-Downloader completes each note

**XHS-Downloader Configuration**:
```python
XHS(
    work_path="datas/xhs_download",
    folder_name="Download",
    name_format="作品标题 作品ID",    # directory ends with _{note_id}
    author_archive=True,              # {author}/ subdirectory
    folder_mode=True,                 # {title}_{note_id}/ subdirectory
    record_data=True,                 # save metadata to ExploreData.db
    download_record=True,             # save IDs to ExploreID.db
    image_format="JPEG",
    image_download=True,
    video_download=True,
)
```

**Logic for each note**:
1. Check if note_id exists in ExploreID.db → skip if yes.
2. Build URL: `https://www.xiaohongshu.com/discovery/item/{note_id}?xsec_token={xsec_token}&xsec_source=pc_share`
3. Call `xhs.extract(url=url, download=True)` — XHS-Downloader fetches HTML, extracts metadata, downloads media.
4. Locate the download directory by searching for `*_{note_id}` under Download/.
5. Write `info.json` (converted to unified format) into that directory.
6. Call `export.export_from_db()` to update Excel/JSON in real-time.

**Breakpoint**: `ExploreID.db`. XHS-Downloader automatically skips notes already in this database.

**Parameters**:
- `--limit N`: process at most N unfinished notes.

**Error handling**:
- 5 consecutive failures → pause 60 seconds, then continue.
- Failed notes are not added to ExploreID.db, so they will be retried on next run.

### export.py

**Responsibility**: Read all metadata from `ExploreData.db`, output Excel/JSON. Pure offline, no API calls.

**Input**: `datas/xhs_download/ExploreData.db`

**Output**: `datas/excel_datas/我的收藏.xlsx` and `我的收藏.json`

**Logic**:
1. `SELECT * FROM explore_data`
2. Write to Excel (using openpyxl) and JSON.

**Stateless**: no breakpoint needed. Idempotent — can be run anytime.

**Core function `export_from_db()`**: called by download.py after each note for real-time updates. Same logic, just triggered incrementally.

### run_all.py

**Responsibility**: Run all three steps in sequence.

```python
fetch_list.main()
download.main(limit=args.limit)
export.main()
```

**Parameters**: `--limit N` (passed through to download.py).

## info.json Format

Written by download.py into each note's download directory. Unified format used for both old Spider_XHS data and new XHS-Downloader data:

```json
{
  "note_id": "68cca73a000000001302b10d",
  "note_url": "https://www.xiaohongshu.com/explore/68cca73a000000001302b10d",
  "note_type": "图集",
  "user_id": "5c89acc00000000016031277",
  "home_url": "https://www.xiaohongshu.com/user/profile/5c89acc00000000016031277",
  "nickname": "JackTu",
  "avatar": "",
  "title": "Google Stich 3.0：UI设计新范式",
  "desc": "...",
  "liked_count": "103",
  "collected_count": "45",
  "comment_count": "12",
  "share_count": "8",
  "video_cover": null,
  "video_addr": null,
  "image_list": ["https://..."],
  "tags": ["设计", "UI"],
  "upload_time": "2026-01-15_10.30.00",
  "ip_location": ""
}
```

Note: `ip_location` is always empty for XHS-Downloader sourced data (not available from HTML extraction).

## One-Time Migration: migrate_old_data.py

**Purpose**: Migrate 591 notes previously downloaded by Spider_XHS into the target format, so download.py and export.py only deal with one data source.

**Input**: `Spider_XHS/datas/media_datas/*/info.json` (591 files)

**Actions**:
1. Read each info.json.
2. Insert note_id into `ExploreID.db` → download.py will skip these.
3. Convert fields and insert into `ExploreData.db` → export.py can read them.

**Field mapping** (Spider_XHS → ExploreData.db):

| info.json field | ExploreData.db column |
|---|---|
| (current time) | 采集时间 |
| note_id | 作品ID |
| note_type | 作品类型 |
| title | 作品标题 |
| desc | 作品描述 |
| tags (joined) | 作品标签 |
| upload_time | 发布时间 |
| upload_time | 最后更新时间 |
| collected_count | 收藏数量 |
| comment_count | 评论数量 |
| share_count | 分享数量 |
| liked_count | 点赞数量 |
| nickname | 作者昵称 |
| user_id | 作者ID |
| home_url | 作者链接 |
| note_url | 作品链接 |
| image_list / video_addr (json) | 下载地址 |
| "" | 动图地址 |

**Media files**: Stay in `Spider_XHS/datas/media_datas/`. Not moved. They're already downloaded and don't need to be in the new directory structure for analysis purposes.

**Run once, then never again.**

## ExploreData.db Schema

```sql
CREATE TABLE IF NOT EXISTS explore_data (
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
);
```

## Breakpoint Summary

| Script | Breakpoint mechanism | Resume behavior |
|--------|---------------------|-----------------|
| fetch_list.py | `collect_list.json` exists | Incremental: only fetch new notes |
| download.py | `ExploreID.db` records | Auto-skip downloaded notes |
| export.py | None (stateless) | Idempotent, always regenerates |
| migrate_old_data.py | `ExploreID.db` (INSERT OR REPLACE) | Idempotent, safe to re-run |
