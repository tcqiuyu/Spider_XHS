"""
scripts/run_all.py — 完整数据采集管道

链接三个步骤的完整处理流程：
  1. fetch_list    — 获取小红书收藏列表
  2. download      — 下载笔记详情与媒体文件
  3. export        — 从数据库导出 Excel + JSON

用法：
  python -m scripts.run_all              # 完整管道
  python -m scripts.run_all --limit 10   # 限制 download 步骤处理数量
"""

import argparse
import sys
from pathlib import Path

from loguru import logger

# 添加项目根目录到 sys.path
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))

from scripts.fetch_list import main as fetch_list_main
from scripts.download import main as download_main
from scripts.export import main as export_main


def main() -> None:
    """执行完整的数据采集管道。"""
    parser = argparse.ArgumentParser(
        description="执行完整的小红书数据采集管道"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="download 步骤的处理数量限制（默认不限制）",
    )
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="强制重新拉取收藏列表（刷新 xsec_token）",
    )
    args = parser.parse_args()

    logger.info("=" * 50)
    logger.info("=== 步骤 1: 获取小红书收藏列表 ===")
    logger.info("=" * 50)
    try:
        fetch_list_main(refresh=args.refresh)
        logger.info("✓ 步骤 1 完成")
    except Exception as e:
        logger.error(f"✗ 步骤 1 失败: {e}")
        raise

    logger.info("")
    logger.info("=" * 50)
    logger.info("=== 步骤 2: 下载笔记详情与媒体文件 ===")
    logger.info("=" * 50)
    try:
        download_main(limit=args.limit)
        logger.info("✓ 步骤 2 完成")
    except Exception as e:
        logger.error(f"✗ 步骤 2 失败: {e}")
        raise

    logger.info("")
    logger.info("=" * 50)
    logger.info("=== 步骤 3: 导出数据为 Excel + JSON ===")
    logger.info("=" * 50)
    try:
        export_main()
        logger.info("✓ 步骤 3 完成")
    except Exception as e:
        logger.error(f"✗ 步骤 3 失败: {e}")
        raise

    logger.info("")
    logger.info("=" * 50)
    logger.info("✓ 完整管道执行成功！")
    logger.info("=" * 50)


if __name__ == "__main__":
    main()
