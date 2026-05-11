import json
import csv
import os
from datetime import datetime

# ✅ Define the folder containing all JSON files here
JSON_DIR = "/Users/lorimartella/scheduled_jobs/agvend/skus"
OUTPUT_DIR = "/Users/lorimartella/scheduled_jobs/agvend/output"

def list_json_files(directory):
    """Return a list of all JSON files in the directory."""
    files = [f for f in os.listdir(directory) if f.endswith(".json")]
    if not files:
        print(f"No JSON files found in {directory}")
        exit(1)
    return sorted(files)

def select_file(files, prompt_text):
    """Prompt user to select a JSON file by number."""
    print(f"\nSelect {prompt_text} file:")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    while True:
        choice = input("Enter number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]
        print("Invalid choice, try again.")

def load_json_data(file_path):
    """Load JSON data and return dict keyed by 'key'."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {item["key"]: item for item in data["data"]}
    except json.JSONDecodeError as e:
        print(f"Error loading JSON file {file_path}: {e}")
        exit(1)

def compare_json_files(old_data, new_data):
    """Compare two sets of product data and find changes."""
    changes = []

    old_keys = set(old_data.keys())
    new_keys = set(new_data.keys())

    # Collect all superseded_by values from new data
    superseded_by_values = {rec.get("superseded_by") for rec in new_data.values() if rec.get("superseded_by")}

    # --- Detect new records ---
    for key in new_keys - old_keys:
        new_record = new_data[key]
        product_key = new_record.get("product_key")
        # Only add if this key does NOT appear as a superseded_by value
        if key not in superseded_by_values:
            changes.append({
                "product_key": product_key,
                "key": key,
                "change_type": "new",
                "old_sku_key": "",   # leave empty for new
                "new_sku_key": "",   # leave empty for new
                "note": ""
            })

    # --- Detect removed records ---
    for key in old_keys - new_keys:
        old_record = old_data[key]
        product_key = old_record.get("product_key")
        changes.append({
            "product_key": product_key,
            "key": key,
            "change_type": "removed",
            "old_sku_key": "",    # leave empty for removed
            "new_sku_key": "",    # changed from key to empty string
            "note": ""
        })

    # --- Detect updates (same product_key, different key) ---
    old_by_product = {}
    for rec in old_data.values():
        old_by_product.setdefault(rec["product_key"], []).append(rec)

    new_by_product = {}
    for rec in new_data.values():
        new_by_product.setdefault(rec["product_key"], []).append(rec)

    for product_key, old_recs in old_by_product.items():
        new_recs = new_by_product.get(product_key, [])
        old_keys_set = {r["key"] for r in old_recs}
        new_keys_set = {r["key"] for r in new_recs}
        if old_keys_set != new_keys_set:
            removed_keys = old_keys_set - new_keys_set
            added_keys = new_keys_set - old_keys_set
            for old_k in removed_keys:
                for new_k in added_keys:
                    changes.append({
                        "product_key": product_key,
                        "key": new_k,
                        "change_type": "updated",
                        "old_sku_key": old_k,
                        "new_sku_key": new_k,
                        "note": ""
                    })

    # --- Detect superseded records ---
    for rec in new_data.values():
        superseded_by = rec.get("superseded_by")
        if superseded_by:
            product_key = rec.get("product_key")
            changes.append({
                "product_key": product_key,
                "key": rec["key"],
                "change_type": "change",
                "old_sku_key": rec["key"],
                "new_sku_key": superseded_by,
                "note": "superseded"
            })

    return changes

def get_dynamic_output_csv(base_name="compare_sku_results"):
    """Generate a date-stamped CSV filename with incremental counter."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"{base_name}_{today_str}.csv"
    output_path = os.path.join(OUTPUT_DIR, filename)
    counter = 1
    while os.path.exists(output_path):
        output_path = os.path.join(OUTPUT_DIR, f"{base_name}_{today_str}_{counter}.csv")
        counter += 1
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return output_path

def write_changes_to_csv(changes, output_path):
    """Write comparison results to CSV."""
    fieldnames = ["product_key", "key", "change_type", "old_sku_key", "new_sku_key", "note"]
    with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(changes)
    print(f"\n✅ Changes written to: {output_path}\n")

def main():
    print(f"JSON directory: {JSON_DIR}\n")
    files = list_json_files(JSON_DIR)

    old_filename = select_file(files, "OLD (previous day)")
    new_filename = select_file(files, "NEW (current day)")

    old_path = os.path.join(JSON_DIR, old_filename)
    new_path = os.path.join(JSON_DIR, new_filename)

    old_data = load_json_data(old_path)
    new_data = load_json_data(new_path)

    print(f"\nComparing:\n  OLD: {old_filename}\n  NEW: {new_filename}\n")

    changes = compare_json_files(old_data, new_data)

    if not changes:
        print("No changes detected.\n")
    else:
        output_csv = get_dynamic_output_csv()
        write_changes_to_csv(changes, output_csv)

if __name__ == "__main__":
    main()
