import sqlite3
import json
import logging
import traceback
import sys
from pathlib import Path
from core.logger_config import setup_logging
from core.system_config import sys_config

# Reconfigure stdout for utf-8
sys.stdout.reconfigure(encoding='utf-8')

# Setup logging
setup_logging(log_dir=sys_config.run_log_dir, level=logging.DEBUG)

from core.invoice_generator.generate_invoice import run_invoice_generation, InvoiceGenerationRequest, InvoicePathConfig, ExplicitOverrides, GenerationOptions

# 1. Fetch blueprint from DB
conn = sqlite3.connect('database/invoice_registry.db')
cursor = conn.cursor()
row = cursor.execute("SELECT config_json, template_json FROM blueprints WHERE customer_code='JF' AND locale='KH'").fetchone()
config_json = row[0]
template_json = row[1]

config_data = json.loads(config_json)
template_data = json.loads(template_json)

# Save to temp files to simulate regular config loading
temp_config_path = Path("database/temp_uploads/runtime_blueprints/JF_KH_config.json")
temp_config_path.parent.mkdir(parents=True, exist_ok=True)
with open(temp_config_path, "w", encoding="utf-8") as f:
    json.dump(config_data, f, indent=2)

temp_template_path = Path("database/temp_uploads/runtime_blueprints/JF_KH_template.json")
with open(temp_template_path, "w", encoding="utf-8") as f:
    json.dump(template_data, f, indent=2)

# Load invoice data
with open("database/temp_uploads/processed/JF26024.json", "r", encoding="utf-8") as f:
    invoice_data = json.load(f)

# Run generation
output_path = Path("database/generated_invoices/reproduce_output.xlsx")
req = InvoiceGenerationRequest(
    paths=InvoicePathConfig(
        input_data_path=Path("database/temp_uploads/processed/JF26024.json"),
        output_path=output_path,
        template_dir=Path("database/temp_uploads/runtime_blueprints"),
        config_dir=Path("database/temp_uploads/runtime_blueprints")
    ),
    overrides=ExplicitOverrides(
        explicit_config_path=temp_config_path,
        explicit_template_path=None,
        input_data_dict=invoice_data
    ),
    options=GenerationOptions(
        daf_mode=False,
        custom_mode=False,
        enable_auto_fit=False,
        return_bytes=False
    )
)

print("Running invoice generation...")
log_file_path = Path("run_log/current_session.log")
start_offset = 0
if log_file_path.exists():
    start_offset = log_file_path.stat().st_size

try:
    res = run_invoice_generation(req)
    print("Generation completed. Output saved to:", res)
except Exception as e:
    print("Generation failed with exception:")
    traceback.print_exc()

print("\n--- Current Session Log (Appended) ---")
if log_file_path.exists():
    with open(log_file_path, "r", encoding="utf-8", errors='ignore') as f:
        f.seek(start_offset)
        for line in f:
            print(line.strip())
