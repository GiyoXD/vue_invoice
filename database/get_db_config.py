import sqlite3
import json
import argparse
from pathlib import Path

def list_blueprints(db_path: str, search_query: str = None):
    """
    Lists available blueprints in the database. Optionally filters by a search query.
    """
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        if search_query:
            query = """
                SELECT id, customer_code, locale, description 
                FROM blueprints 
                WHERE customer_code LIKE ? OR locale LIKE ? OR description LIKE ?
            """
            term = f"%{search_query}%"
            cursor.execute(query, (term, term, term))
        else:
            query = "SELECT id, customer_code, locale, description FROM blueprints"
            cursor.execute(query)
            
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            print("No blueprints found.")
            return
            
        print(f"\n{'ID':<5} | {'Customer':<12} | {'Locale':<6} | {'Description'}")
        print("-" * 80)
        for row in rows:
            desc = row[3][:50] + "..." if row[3] and len(row[3]) > 50 else (row[3] or "")
            print(f"{row[0]:<5} | {row[1]:<12} | {row[2]:<6} | {desc}")
        print()
    except Exception as e:
        print(f"Error reading database: {e}")

def fetch_config(db_path: str, customer: str, locale: str, output_dir: str):
    """
    Fetches the configuration JSON and template JSON from the blueprints table for a given customer code and locale,
    and writes them to files.
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    query = "SELECT config_json, template_json FROM blueprints WHERE customer_code = ? AND locale = ?"
    cursor.execute(query, (customer, locale))
    row = cursor.fetchone()
    
    if not row:
        # Check if customer exists but with different locales
        cursor.execute("SELECT locale FROM blueprints WHERE customer_code = ?", (customer,))
        locales = [r[0] for r in cursor.fetchall()]
        
        if locales:
            print(f"\nError: No configuration found for customer '{customer}' with locale '{locale}'.")
            print(f"Available locales for '{customer}': {', '.join(locales)}")
        else:
            # Check if there are similar customer codes (case-insensitive substring search)
            cursor.execute("SELECT DISTINCT customer_code FROM blueprints WHERE customer_code LIKE ?", (f"%{customer}%",))
            similar = [r[0] for r in cursor.fetchall()]
            print(f"\nError: Customer '{customer}' not found.")
            if similar:
                print(f"Did you mean: {', '.join(similar)}?")
            else:
                print("Run with --list to see all available configurations.")
        print()
        conn.close()
        return False
        
    try:
        config_data = json.loads(row[0])
    except Exception as e:
        print(f"Error: Failed to parse configuration JSON from database: {e}")
        config_data = None
        
    try:
        template_data = json.loads(row[1])
    except Exception as e:
        print(f"Error: Failed to parse template JSON from database: {e}")
        template_data = None
        
    out_dir_path = Path(output_dir)
    out_dir_path.mkdir(parents=True, exist_ok=True)
    
    success = False
    
    if config_data is not None:
        output_filename = f"{customer}_{locale}_config.json"
        output_filepath = out_dir_path / output_filename
        try:
            with open(output_filepath, "w", encoding="utf-8") as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            print(f"Success: Configuration exported successfully to {output_filepath}")
            success = True
        except Exception as e:
            print(f"Error: Failed to write configuration to file: {e}")
            
    if template_data is not None:
        output_filename = f"{customer}_{locale}_template.json"
        output_filepath = out_dir_path / output_filename
        try:
            with open(output_filepath, "w", encoding="utf-8") as f:
                json.dump(template_data, f, indent=2, ensure_ascii=False)
            print(f"Success: Template exported successfully to {output_filepath}")
            success = True
        except Exception as e:
            print(f"Error: Failed to write template to file: {e}")
            
    conn.close()
    return success

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch or list JSON configurations from blueprints database.")
    parser.add_argument("--db", default="database/invoice_registry.db", help="Path to sqlite3 database")
    parser.add_argument("--customer", help="Customer code (e.g. JF)")
    parser.add_argument("--locale", default="KH", help="Locale (e.g. KH, VN)")
    parser.add_argument("--outdir", default="database", help="Output directory to save the JSON config")
    parser.add_argument("--list", action="store_true", help="List all available configurations in the database")
    parser.add_argument("--search", help="Search available configurations by keyword")
    
    args = parser.parse_args()
    
    if args.list:
        list_blueprints(args.db)
    elif args.search:
        list_blueprints(args.db, args.search)
    elif args.customer:
        fetch_config(args.db, args.customer, args.locale, args.outdir)
    else:
        # Default behavior: if no specific action, show usage or default to --list
        print("\nNo action specified. Showing all available blueprints:")
        list_blueprints(args.db)
        print("To fetch a configuration, use: python database/get_db_config.py --customer <CODE> [--locale <LOCALE>]")

