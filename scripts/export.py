"""
scripts/export.py — 从 ExploreData.db 导出 Excel + JSON

读取 SQLite 数据库中的所有笔记元数据，输出：
  - Excel: 选定列，供人工分析
  - JSON: 全量记录，供程序消费

可作为独立脚本运行，也可被 download.py 导入调用：
  from scripts.export import export_from_db

默认路径（独立运行时）：
  DB:    datas/xhs_download/Download/ExploreData.db
  Excel: datas/excel_datas/我的收藏.xlsx
  JSON:  datas/excel_datas/我的收藏.json
"""

import json
import re
import sqlite3
from pathlib import Path

import openpyxl
from loguru import logger

# ---------------------------------------------------------------------------
# 列映射：Excel 表头 → 数据库字段名
# ---------------------------------------------------------------------------

_EXCEL_COLUMNS = [
    ("作品ID",   "作品ID"),
    ("标题",     "作品标题"),
    ("类型",     "作品类型"),
    ("作者",     "作者昵称"),
    ("描述",     "作品描述"),
    ("标签",     "作品标签"),
    ("发布时间", "发布时间"),
    ("收藏数",   "收藏数量"),
    ("评论数",   "评论数量"),
    ("分享数",   "分享数量"),
    ("点赞数",   "点赞数量"),
    ("链接",     "作品链接"),
]

_ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _sanitize(value: str) -> str:
    """去除 Excel 不允许的控制字符。"""
    return _ILLEGAL_CHARS_RE.sub("", str(value)) if value is not None else ""


# ---------------------------------------------------------------------------
# 核心导出函数
# ---------------------------------------------------------------------------

def export_from_db(
    db_path: str | Path,
    excel_path: str | Path,
    json_path: str | Path,
) -> int:
    """
    从 SQLite 数据库读取所有笔记元数据，写出 Excel 和 JSON 文件。

    参数
    ----
    db_path   : ExploreData.db 的路径
    excel_path: 输出 Excel 文件路径（.xlsx）
    json_path : 输出 JSON 文件路径（.json）

    返回
    ----
    写出的记录数；若数据库文件不存在则返回 0（不抛异常）。
    """
    db_path = Path(db_path)
    excel_path = Path(excel_path)
    json_path = Path(json_path)

    # ── 数据库不存在 ──────────────────────────────────────────────────────────
    if not db_path.exists():
        logger.warning(f"数据库文件不存在，跳过导出: {db_path}")
        return 0

    # ── 读取数据 ───────────────────────────────────────────────────────────────
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM explore_data")
        rows = [dict(row) for row in cur.fetchall()]
    except sqlite3.Error as exc:
        logger.error(f"读取数据库失败: {exc}")
        return 0
    finally:
        conn.close()

    count = len(rows)

    if count == 0:
        logger.info(f"数据库为空，暂无记录可导出: {db_path}")
        return 0

    # ── 确保输出目录存在 ───────────────────────────────────────────────────────
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.parent.mkdir(parents=True, exist_ok=True)

    # ── 写 Excel ───────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "我的收藏"

    headers = [col[0] for col in _EXCEL_COLUMNS]
    db_fields = [col[1] for col in _EXCEL_COLUMNS]
    ws.append(headers)

    for row in rows:
        ws.append([_sanitize(row.get(field, "")) for field in db_fields])

    wb.save(excel_path)
    logger.info(f"Excel 已导出: {excel_path}（{count} 条）")

    # ── 写 JSON ────────────────────────────────────────────────────────────────
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logger.info(f"JSON 已导出: {json_path}（{count} 条）")

    return count


# ---------------------------------------------------------------------------
# 独立运行入口
# ---------------------------------------------------------------------------

_PROJECT_ROOT = Path(__file__).resolve().parent.parent

_DEFAULT_DB_PATH    = _PROJECT_ROOT / "datas" / "xhs_download" / "Download" / "ExploreData.db"
_DEFAULT_EXCEL_PATH = _PROJECT_ROOT / "datas" / "excel_datas" / "我的收藏.xlsx"
_DEFAULT_JSON_PATH  = _PROJECT_ROOT / "datas" / "excel_datas" / "我的收藏.json"


def main() -> None:
    count = export_from_db(_DEFAULT_DB_PATH, _DEFAULT_EXCEL_PATH, _DEFAULT_JSON_PATH)
    if count:
        logger.info(f"导出完成，共 {count} 条记录")
    else:
        logger.info("导出完成（无数据或数据库不存在）")


if __name__ == "__main__":
    main()
