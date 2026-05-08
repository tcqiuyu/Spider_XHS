"""
scripts/migrate_old_data.py — 一次性迁移脚本

将 Spider_XHS 旧格式的 info.json 文件迁移到 XHS-Downloader 的 SQLite 数据库：
- ExploreID.db  （表 explore_id，列 ID TEXT PRIMARY KEY）
- ExploreData.db（表 explore_data，详见 schema）

扫描 datas/media_datas/ 下所有 info.json 文件，使用 REPLACE INTO 写入，
可安全重复执行（幂等）。
"""

import json
import logging
import os
import sqlite3
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# 项目路径
# ---------------------------------------------------------------------------

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

_MEDIA_DATAS_DIR = _PROJECT_ROOT / "datas" / "media_datas"
_EXPLORE_ID_DB = _PROJECT_ROOT / "datas" / "xhs_download" / "ExploreID.db"
_EXPLORE_DATA_DB = (
    _PROJECT_ROOT / "datas" / "xhs_download" / "Download" / "ExploreData.db"
)

# ---------------------------------------------------------------------------
# 日志配置
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Schema SQL
# ---------------------------------------------------------------------------

_CREATE_EXPLORE_ID_SQL = """
CREATE TABLE IF NOT EXISTS explore_id (
    ID TEXT PRIMARY KEY
);
"""

_CREATE_EXPLORE_DATA_SQL = """
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
"""


# ---------------------------------------------------------------------------
# 辅助函数
# ---------------------------------------------------------------------------


def _ensure_db(db_path: Path, create_sql: str) -> sqlite3.Connection:
    """确保数据库目录存在，创建连接并初始化表结构。"""
    os.makedirs(db_path.parent, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute(create_sql)
    conn.commit()
    return conn


def _map_note_type(note_type: str) -> str:
    """将旧格式的 note_type 映射到 ExploreData 格式。"""
    return "图文" if note_type == "图集" else note_type


def _build_download_addr(info: dict) -> str:
    """根据笔记类型构建下载地址 JSON 字符串。"""
    note_type = info.get("note_type", "")
    if note_type == "图集":
        image_list = info.get("image_list", [])
        return json.dumps(image_list, ensure_ascii=False)
    video_addr = info.get("video_addr", "")
    if video_addr:
        return json.dumps([video_addr], ensure_ascii=False)
    return "[]"


def _build_tags(tags_raw) -> str:
    """将 tags 字段统一转换为空格分隔的字符串。"""
    if isinstance(tags_raw, list):
        return " ".join(tags_raw)
    return str(tags_raw) if tags_raw is not None else ""


def _insert_explore_id(conn: sqlite3.Connection, note_id: str) -> None:
    conn.execute("REPLACE INTO explore_id (ID) VALUES (?)", (note_id,))


def _insert_explore_data(
    conn: sqlite3.Connection, info: dict, collected_at: str
) -> None:
    note_id = info.get("note_id", "")
    upload_time = info.get("upload_time", "")
    tags_raw = info.get("tags", [])

    row = (
        collected_at,                                   # 采集时间
        note_id,                                         # 作品ID
        _map_note_type(info.get("note_type", "")),       # 作品类型
        info.get("title", ""),                           # 作品标题
        info.get("desc", ""),                            # 作品描述
        _build_tags(tags_raw),                           # 作品标签
        upload_time,                                     # 发布时间
        upload_time,                                     # 最后更新时间
        str(info.get("collected_count", "")),            # 收藏数量
        str(info.get("comment_count", "")),              # 评论数量
        str(info.get("share_count", "")),                # 分享数量
        str(info.get("liked_count", "")),                # 点赞数量
        info.get("nickname", ""),                        # 作者昵称
        info.get("user_id", ""),                         # 作者ID
        info.get("home_url", ""),                        # 作者链接
        info.get("note_url", ""),                        # 作品链接
        _build_download_addr(info),                      # 下载地址
        "",                                              # 动图地址
    )

    conn.execute(
        """
        REPLACE INTO explore_data (
            采集时间, 作品ID, 作品类型, 作品标题, 作品描述, 作品标签,
            发布时间, 最后更新时间, 收藏数量, 评论数量, 分享数量, 点赞数量,
            作者昵称, 作者ID, 作者链接, 作品链接, 下载地址, 动图地址
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        row,
    )


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def main() -> None:
    logger.info("开始扫描 %s", _MEDIA_DATAS_DIR)

    id_conn = _ensure_db(_EXPLORE_ID_DB, _CREATE_EXPLORE_ID_SQL)
    data_conn = _ensure_db(_EXPLORE_DATA_DB, _CREATE_EXPLORE_DATA_SQL)

    collected_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    migrated = 0

    try:
        info_files = sorted(_MEDIA_DATAS_DIR.rglob("info.json"))
        logger.info("找到 %d 个 info.json 文件", len(info_files))

        for info_path in info_files:
            try:
                with open(info_path, encoding="utf-8") as f:
                    info = json.load(f)
            except (json.JSONDecodeError, OSError) as exc:
                logger.warning("跳过 %s：%s", info_path, exc)
                continue

            note_id = info.get("note_id", "").strip()
            if not note_id:
                logger.warning("跳过 %s：缺少 note_id", info_path)
                continue

            _insert_explore_id(id_conn, note_id)
            _insert_explore_data(data_conn, info, collected_at)
            migrated += 1

        id_conn.commit()
        data_conn.commit()

    finally:
        id_conn.close()
        data_conn.close()

    logger.info("迁移完成: %d 条 → ExploreID.db + ExploreData.db", migrated)


if __name__ == "__main__":
    main()
