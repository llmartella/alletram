import json
import csv

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

with open("", "r") as f:
    data = json.load(f)

rows = flatten_transactions(data)

if rows:
    with open("/Users/lorimartella/Documents/gmatter/agvend/support/python/untitled folder/agtegra_bayer_20260529.json", "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)

print(f"Done — {len(rows)} rows written.")
