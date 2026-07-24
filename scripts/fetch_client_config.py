#!/usr/bin/env python3
"""
CLI script to search and fetch client blueprint configurations from the SQLite database.
Default output directory: scripts/
"""

import sys
import json
import argparse
import logging
from pathlib import Path

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.database.session import SessionLocal
from core.database.repositories.blueprint_repository import BlueprintRepository
from core.database.models import Blueprint

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "database" / "temp_uploads"


def list_blueprints(db_repo: BlueprintRepository):
    """Fetch and display all blueprints in the database."""
    blueprints = db_repo.get_all_blueprints()
    if not blueprints:
        print("No blueprints found in database.")
        return

    print(f"\nFound {len(blueprints)} blueprint(s) in database:\n")
    print(f"{'Customer Code':<30} {'Locale':<10} {'Description'}")
    print("-" * 75)
    for bp in blueprints:
        desc = bp.description or "N/A"
        print(f"{bp.customer_code:<30} {bp.locale:<10} {desc}")
    print()


def search_blueprints(db_repo: BlueprintRepository, query: str):
    """Search blueprints matching client name query (case-insensitive)."""
    all_bps = db_repo.get_all_blueprints()
    query_lower = query.lower()
    matches = [
        bp for bp in all_bps
        if query_lower in bp.customer_code.lower() or (bp.description and query_lower in bp.description.lower())
    ]

    if not matches:
        print(f"\nNo blueprints found matching query: '{query}'")
        return

    print(f"\nFound {len(matches)} matching blueprint(s) for '{query}':\n")
    print(f"{'Customer Code':<30} {'Locale':<10} {'Description'}")
    print("-" * 75)
    for bp in matches:
        desc = bp.description or "N/A"
        print(f"{bp.customer_code:<30} {bp.locale:<10} {desc}")
    print()


def fetch_config(
    db_repo: BlueprintRepository,
    customer_code: str,
    locale: str = None,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    fetch_template: bool = False
) -> Path:
    """Fetch blueprint config JSON for a client and save to target directory."""
    variants = db_repo.get_customer_variants(customer_code)

    if not variants:
        # Try case-insensitive search
        all_bps = db_repo.get_all_blueprints()
        matched = [bp for bp in all_bps if bp.customer_code.lower() == customer_code.lower()]
        if matched:
            variants = matched
            customer_code = matched[0].customer_code

    if not variants:
        logger.error(f"No blueprint found for customer code: '{customer_code}'")
        sys.exit(1)

    target_bp: Optional[Blueprint] = None
    if locale:
        target_bp = next((bp for bp in variants if bp.locale.upper() == locale.upper()), None)
        if not target_bp:
            available_locales = [bp.locale for bp in variants]
            logger.error(f"Locale '{locale}' not found for '{customer_code}'. Available locales: {available_locales}")
            sys.exit(1)
    else:
        # Default to KH if present, otherwise first available
        target_bp = next((bp for bp in variants if bp.locale.upper() == "KH"), variants[0])

    config_data = target_bp.config_json
    if not config_data:
        logger.error(f"Blueprint for '{target_bp.customer_code}' ({target_bp.locale}) has no config_json content.")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    out_filename = f"{target_bp.customer_code}_{target_bp.locale}_config.json"
    out_path = output_dir / out_filename

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(config_data, f, indent=2, ensure_ascii=False)

    logger.info(f"✅ Successfully fetched config for '{target_bp.customer_code}' [{target_bp.locale}] -> {out_path}")

    if fetch_template and target_bp.template_binary and target_bp.template_binary.xlsx_blob:
        xlsx_filename = f"{target_bp.customer_code}_{target_bp.locale}_template.xlsx"
        xlsx_path = output_dir / xlsx_filename
        with open(xlsx_path, "wb") as f:
            f.write(target_bp.template_binary.xlsx_blob)
        logger.info(f"  ✅ Saved binary template Excel file -> {xlsx_path}")

    return out_path


def main():
    parser = argparse.ArgumentParser(
        description="Search and fetch client blueprint configurations from the database."
    )
    parser.add_argument(
        "client",
        nargs="?",
        default=None,
        help="Client name / customer code (e.g. JF, CLW, MOTO)"
    )
    parser.add_argument(
        "-s", "--search",
        type=str,
        help="Search blueprints by keyword matching customer code or description"
    )
    parser.add_argument(
        "-l", "--list",
        action="store_true",
        help="List all blueprints available in the database"
    )
    parser.add_argument(
        "--locale",
        type=str,
        default=None,
        help="Locale code (e.g. KH, VN). Defaults to KH or first available"
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Output directory to save fetched config file (default: {DEFAULT_OUTPUT_DIR})"
    )
    parser.add_argument(
        "-t", "--template",
        action="store_true",
        help="Also fetch the binary template .xlsx file if available"
    )

    args = parser.parse_args()

    db = SessionLocal()
    try:
        db_repo = BlueprintRepository(db)

        if args.list:
            list_blueprints(db_repo)
            return

        if args.search:
            search_blueprints(db_repo, args.search)
            return

        if not args.client:
            parser.print_help()
            print("\nTip: Use --list or --search <keyword> to view available clients.")
            return

        out_dir = Path(args.output).resolve()
        fetch_config(
            db_repo=db_repo,
            customer_code=args.client,
            locale=args.locale,
            output_dir=out_dir,
            fetch_template=args.template
        )
    finally:
        db.close()


if __name__ == "__main__":
    main()
