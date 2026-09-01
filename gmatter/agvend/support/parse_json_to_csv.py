import json
import csv
from datetime import datetime

##put json in csv with all transactions

def flatten_transactions(data):
    rows = []
    participant_key = data.get("participant_key", "")
    
    for entity in data.get("entities", []):
        entity_id = entity.get("id", "")
        
        for txn in entity.get("transactions", []):
            row = {
                "participant_key": participant_key,
                "entity_id": entity_id,
                "transaction_id": txn.get("transaction_id", ""),
                "transaction_type": txn.get("transaction_type", ""),
                "sku_key": txn.get("sku_key", ""),
                "uom_key": txn.get("uom_key", ""),
                "quantity": txn.get("quantity", ""),
                "amount": txn.get("amount", ""),
                "invoice_date": txn.get("invoice_date", ""),
                "purchaser_id": txn.get("purchaser", {}).get("id", ""),
                "seller_key": txn.get("seller", {}).get("key", ""),  # <-- add this line
}
            rows.append(row)
    return rows

with open("/Users/lorimartella/Downloads/bayer_2026_invoices_and_purchase_invoices__request_agtegra_july_30.json", "r") as f:
    data = json.load(f)

rows = flatten_transactions(data)

from datetime import datetime

if rows:
    filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(filename, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)

    print(f"Done — {len(rows)} rows written to {filename}.")
