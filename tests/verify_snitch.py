import logging
import sys
import os
from pathlib import Path

# Add project root to path
sys.path.append(os.getcwd())

from core.utils.snitch import start_trace
from core.blueprint_generator.blueprint_generator import BlueprintGenerator
from core.orchestrator import Orchestrator

# Setup logging to stdout
logging.basicConfig(level=logging.INFO, stream=sys.stdout, force=True)

def verify_snitch():
    print("--- STARTING SNITCH VERIFICATION ---")
    
    # 1. Start Trace
    tid = start_trace("TEST-TRACE-001")
    print(f"Trace ID Initialized: {tid}")
    
    # 2. Test Blueprint Generator (expect failure but LOGGED failure)
    bg = BlueprintGenerator(output_base_dir="output/test_snitch")
    try:
        print("Calling BlueprintGenerator.generate...")
        bg.generate("NON_EXISTENT_FILE.xlsx", dry_run=True)
    except Exception as e:
        print(f"Caught expected error: {e}")

    # 3. Test Orchestrator (Data Parser)
    # We pass a dummy path, expect CRASH log
    orch = Orchestrator()
    try:
        print("Calling Orchestrator.process_excel_to_json...")
        orch.process_excel_to_json(Path("fake.xlsx"), Path("output"))
    except Exception as e:
         print(f"Caught expected error: {e}")

if __name__ == "__main__":
    verify_snitch()
