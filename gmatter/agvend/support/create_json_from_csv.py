import csv
import json
from datetime import datetime
from collections import defaultdict

INPUT_FILE = r"/Users/lorimartella/Documents/gmatter/agvend/support/redstar_helena_only.csv" #provide path to source file
OUTPUT_FILE = r"/Users/lorimartella/Documents/gmatter/agvend/support/python/output/reconstructed_helena.json" #provide path to output file


def csv_to_json(input_file, output_file):
    # Read all rows from CSV
    with open(input_file, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if not rows:
        print("No data found in CSV.")
        return

    # Strip BOM from column names if present
    rows = [{k.lstrip('\ufeff'): v for k, v in row.items()} for row in rows]

    print("Detected columns:", rows[0].keys())

    # Group by participant_key -> entity_id -> transactions
    participants = defaultdict(lambda: defaultdict(list))

    for row in rows:
        participant_key = row["participant_key"]
        entity_id = row["entity_id"]

        transaction = {
            "transaction_id": row["transaction_id"],
            "transaction_type": row["transaction_type"],
            "sku_key": row["sku_key"],
            "uom_key": row["uom_key"],
            "quantity": float(row["quantity"]),
            "amount": float(row["amount"]),
            "invoice_date": datetime.strptime(row["invoice_date"], "%m/%d/%y").strftime("%Y-%m-%d"),  # <-- updated
            "properties": [],
        }

        # Only include purchaser if data is present
        if row.get("purchaser_id"):
            transaction["purchaser"] = {"id": row["purchaser_id"]}

        # Only include seller if data is present
        if row.get("seller_key"):
            transaction["seller"] = {"key": row["seller_key"]}

        participants[participant_key][entity_id].append(transaction)

    # Build the final JSON structure
    results = []
    for participant_key, entities in participants.items():
        obj = {
            "participant_key": participant_key,
            "request_id": "",
            "properties": [],
            "entities": [
                {
                    "id": entity_id,
                    "properties": [],
                    "transactions": transactions
                }
                for entity_id, transactions in entities.items()
            ],
            "relationships": []
        }
        results.append(obj)

    # If only one participant, output a single object
    output = results[0] if len(results) == 1 else results

    with open(output_file, "w") as f:
        json.dump(output, f, indent=4)

    print(f"Done — {len(rows)} rows converted, {len(results)} participant(s) written.")

csv_to_json(INPUT_FILE, OUTPUT_FILE)
