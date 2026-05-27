#this will get the results from the GET /skus API and save to a dated json file

#!/usr/bin/env python3
import requests
import json
from datetime import date, datetime
import os
from dotenv import load_dotenv

# --- LOAD ENV VARIABLES ---
load_dotenv("/Users/lorimartella/scheduled_jobs/agvend/agvend_python_scripts/.env")  # Adjust path if needed
API_KEY = os.getenv("API_KEY")

# --- SETTINGS ---
URL = "https://dev.calc.ag/skus"  # Replace with your endpoint
OUTPUT_DIR = "/Users/lorimartella/scheduled_jobs/agvend/skus"  # Change path as needed

HEADERS = {
    "Accept": "application/json",
    "x-api-key": os.getenv("API_KEY"),
}
response = requests.get(URL, headers=HEADERS)


# --- SCRIPT ---
try:
    response = requests.get(URL, headers=HEADERS, timeout=30)
    response.raise_for_status()
    data = response.json()
except Exception as e:
    print(f"[{datetime.now()}] ❌ API request failed: {e}")
    exit(1)

os.makedirs(OUTPUT_DIR, exist_ok=True)

filename = f"response_{date.today().isoformat()}.json"
filepath = os.path.join(OUTPUT_DIR, filename)

with open(filepath, "w") as f:
    json.dump(data, f, indent=2)

print(f"[{datetime.now()}] ✅ Saved: {filepath}")
