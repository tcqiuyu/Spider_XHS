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
