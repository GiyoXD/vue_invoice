"""
Standalone Debug Runner for Blueprint Generator.
Run this script directly in Python or set breakpoints in VS Code to debug blueprint creation.

Usage:
    python scratch/debug_blueprint_runner.py [path/to/template.xlsx]
"""
import sys
import os
import logging
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.blueprint_generator.generator import BlueprintGenerator, BlueprintGenerationOptions
from core.logger_config import setup_logging
from core.system_config import sys_config


def main():
    setup_logging(log_dir=sys_config.run_log_dir, level=logging.DEBUG)

    sample_dir = PROJECT_ROOT / "database" / "temp_uploads"
    output_dir = sample_dir / "runtime_blueprints"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Change this sample name to debug specific template
    target_sample = "CT&INV&PL JKVN26006 FCA.xlsx"

    if len(sys.argv) > 1:
        arg_path = Path(sys.argv[1])
        if arg_path.exists():
            template_path = arg_path
        elif (sample_dir / arg_path.name).exists():
            template_path = sample_dir / arg_path.name
        else:
            print(f"[ERROR] Template file not found: {sys.argv[1]}")
            return
    elif target_sample and (sample_dir / target_sample).exists():
        template_path = sample_dir / target_sample
    else:
        xlsx_files = [f for f in sample_dir.glob("*.xlsx") if not f.name.startswith("~$")]
        if not xlsx_files:
            print(f"[ERROR] No valid .xlsx files found in {sample_dir}")
            return
        template_path = xlsx_files[0]

    print(f"[RUNNING] Blueprint generator on: {template_path}")
    print(f"[OUTPUT] Runtime blueprint dir: {output_dir}")

    options = BlueprintGenerationOptions(
        output_dir=str(output_dir),
        dry_run=False
    )

    try:
        generator = BlueprintGenerator()
        result_path = generator.generate(
            template_path=str(template_path),
            options=options
        )
        print(f"[SUCCESS] Blueprint generated at: {result_path}")
    except Exception as e:
        print(f"[FAILED] Blueprint generator error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
