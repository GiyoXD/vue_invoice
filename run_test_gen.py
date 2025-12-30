import sys
from pathlib import Path
from core.orchestrator import Orchestrator
from core.system_config import sys_config

def run_test():
    try:
        orch = Orchestrator()
        
        # 1. Define Paths
        input_excel = Path(r"c:\Users\JPZ031127\Desktop\vue_invoice_project\database\blueprints\test_subjects_invoice\JF25062.xlsx")
        config_path = Path("core/invoice_generator/JF_bundle_config.json")
        template_path = sys_config.templates_dir / "JF.xlsx"
        
        print(f"--- Starting Test ---")
        print(f"Input: {input_excel}")
        
        # 2. Process Excel -> JSON
        print("\nStep 1: Processing Excel to JSON...")
        json_path, identifier = orch.process_excel_to_json(
            excel_path=input_excel,
            output_dir=sys_config.temp_uploads_dir
        )
        print(f"JSON Generated: {json_path}")
        
        # 3. Generate Invoice
        print("\nStep 2: Generating Invoice...")
        output_path = sys_config.output_dir / f"{identifier}_Result.xlsx"
        
        # Construct flags to pass explicit config/template
        flags = ["--config", str(config_path), "--template", str(template_path)]
        
        result = orch.generate_invoice(
            json_path=json_path,
            output_path=output_path,
            template_dir=sys_config.templates_dir,
            config_dir=sys_config.registry_dir,
            flags=flags
        )
        print(f"\nSUCCESS! Invoice generated at: {result}")
        
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_test()
