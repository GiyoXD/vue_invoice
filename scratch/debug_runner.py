"""
Standalone Debug Runner for Invoice Generator.
Run this script directly in Python or set breakpoints in VS Code to debug layout building.

Usage:
    python scratch/debug_runner.py [path/to/json_file.json]
"""
import sys
import os
import json
from pathlib import Path


# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.invoice_generator.generate_invoice import run_invoice_generation
from core.invoice_generator.models.request import (
    InvoiceGenerationRequest,
    InvoicePathConfig,
    GenerationOptions,
    ExplicitOverrides
)



def main():
    # Find sample json file if not provided as argument
    sample_dir = PROJECT_ROOT / "database" / "temp_uploads" / "processed"

    if len(sys.argv) > 1:
        arg_path = Path(sys.argv[1])
        if arg_path.exists():
            json_path = arg_path
        elif (sample_dir / arg_path.name).exists():
            json_path = sample_dir / arg_path.name
        else:
            print(f"Error: JSON file '{sys.argv[1]}' not found directly or in {sample_dir}")
            return
    else:
        target_file = sample_dir / "test1111111111.json"
        if target_file.exists():
            json_path = target_file
        else:
            json_files = [f for f in sample_dir.glob("*.json") if "metadata" in f.read_text(encoding='utf-8', errors='ignore')]
            if not json_files:
                print(f"No valid JSON files found in {sample_dir}")
                return
            json_path = json_files[0]


    output_path = PROJECT_ROOT / "scratch" / "output_test.xlsx"

    print(f"[RUNNING] Invoice generator on: {json_path}")
    print(f"[OUTPUT] Excel path: {output_path}")

    with open(json_path, 'r', encoding='utf-8') as f:
        invoice_data = json.load(f)

    req = InvoiceGenerationRequest(
        paths=InvoicePathConfig(
            input_data_path=Path(json_path),
            output_path=Path(output_path)
        ),
        overrides=ExplicitOverrides(
            input_data_dict=invoice_data
        ),
        options=GenerationOptions(
            daf_mode=False,
            custom_mode=False
        )
    )



    from core.invoice_generator.mappers import resolve_summary_payload

    print("\n--- Testing resolve_summary_payload ---")
    
    summary_payload = resolve_summary_payload(invoice_data=invoice_data)
    print(f"Resolved Payload Structure:\n  grand_total: {summary_payload['grand_total']}\n  weight_summary: {summary_payload['weight_summary']}\n  leather_summary: {summary_payload['leather_summary']}")
    print("---------------------------------------\n")


    try:
        res = run_invoice_generation(req)
        print("[SUCCESS] Invoice generation completed successfully!")

        if output_path.exists():
            print(f"[OK] File size: {output_path.stat().st_size} bytes")
    except Exception as e:

        print(f"[FAILED] Invoice generation error: {e}")
        import traceback
        traceback.print_exc()



if __name__ == "__main__":
    main()
