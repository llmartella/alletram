"""
send_json_payload.py

Sends a JSON file as the body of an HTTP request to an endpoint.
Handles large files by streaming them instead of loading the whole
thing into memory (which is usually why Postman chokes/crashes on
big payloads).

HOW TO USE
----------
Option 1 - Fully hardcoded:
    Fill in ENDPOINT_URL, AUTH, and FILE_PATH below, then hit Run in VS Code.

Option 2 - File picker popup:
    Fill in ENDPOINT_URL and AUTH, but leave FILE_PATH = None.
    When you run the script, a file picker window will pop up so you
    can browse to and select the JSON file.

Run with the VS Code "Run" button (or F5) like normal.
"""

import json
import os
import sys
from datetime import datetime

import requests

# ============================================================
# CONFIG - edit these
# ============================================================

ENDPOINT_URL = "https://dev.calc.ag/calculations?program_supplier_key=basf&time_frame=2026"

# Set FILE_PATH to a specific file to skip the popup (Option 1),
# or leave as None to get a file-picker window (Option 2).
# FILE_PATH = None
FILE_PATH = "/Users/lorimartella/Downloads/bayer_2026_invoices_and_purchase_invoices__request_agtegra.json"

# HTTP method to use - almost always POST or PUT
HTTP_METHOD = "POST"

# ---- Auth: fill in ONE of these blocks and set AUTH_TYPE ----
AUTH_TYPE = "api_key_header"  # options: "none", "bearer", "basic", "api_key_header"

BEARER_TOKEN = ""  # used if AUTH_TYPE == "bearer"

BASIC_AUTH_USERNAME = ""  # used if AUTH_TYPE == "basic"
BASIC_AUTH_PASSWORD = ""

API_KEY_HEADER_NAME = "x-api-key"  # used if AUTH_TYPE == "api_key_header"
API_KEY_VALUE = "TvJPHuzyDi3yuV"

# Any extra headers you need to send, beyond auth/content-type
EXTRA_HEADERS = {
    # "X-Custom-Header": "some-value",
}

# Request timeout in seconds (large payloads may need this higher)
TIMEOUT_SECONDS = 120

# Where to save the response. Leave as None to save it in the same folder
# as the input file (recommended). Set to a specific folder path if you'd
# rather it always go somewhere fixed, e.g. "/Users/lorimartella/Downloads"
OUTPUT_DIR = None

# Set to False if you don't want responses saved to disk at all
SAVE_RESPONSE = True

# ============================================================
# You shouldn't need to edit anything below this line
# ============================================================


def pick_file_via_dialog():
    """Opens a native file picker and returns the selected path, or None if cancelled."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # hide the empty root window
    root.attributes("-topmost", True)  # bring the dialog to the front

    path = filedialog.askopenfilename(
        title="Select JSON file to send",
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
    )

    root.destroy()
    return path or None


def build_headers():
    headers = {"Content-Type": "application/json"}
    headers.update(EXTRA_HEADERS)

    if AUTH_TYPE == "bearer":
        if not BEARER_TOKEN:
            sys.exit("AUTH_TYPE is 'bearer' but BEARER_TOKEN is empty. Fill it in and re-run.")
        headers["Authorization"] = f"Bearer {BEARER_TOKEN}"
    elif AUTH_TYPE == "api_key_header":
        if not API_KEY_VALUE:
            sys.exit("AUTH_TYPE is 'api_key_header' but API_KEY_VALUE is empty. Fill it in and re-run.")
        headers[API_KEY_HEADER_NAME] = API_KEY_VALUE
    elif AUTH_TYPE not in ("none", "basic"):
        sys.exit(f"Unrecognized AUTH_TYPE: {AUTH_TYPE!r}")

    return headers


def build_auth():
    if AUTH_TYPE == "basic":
        if not BASIC_AUTH_USERNAME or not BASIC_AUTH_PASSWORD:
            sys.exit("AUTH_TYPE is 'basic' but username/password is empty. Fill it in and re-run.")
        return (BASIC_AUTH_USERNAME, BASIC_AUTH_PASSWORD)
    return None


def validate_json_file(path):
    """Quick sanity check that the file is valid JSON before we send it."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            json.load(f)
    except json.JSONDecodeError as e:
        sys.exit(f"File does not look like valid JSON: {path}\n{e}")
    except OSError as e:
        sys.exit(f"Could not read file: {path}\n{e}")


def save_response(input_file_path, response):
    """Saves the full response body (and a few key details) to a file
    next to the input file, or in OUTPUT_DIR if that's set."""

    input_dir = os.path.dirname(os.path.abspath(input_file_path))
    input_name = os.path.splitext(os.path.basename(input_file_path))[0]

    output_dir = OUTPUT_DIR or input_dir
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Try to pretty-print as JSON if possible, otherwise save raw text
    content_type = response.headers.get("Content-Type", "")
    is_json = "json" in content_type.lower()

    extension = "json" if is_json else "txt"
    output_path = os.path.join(
        output_dir, f"{input_name}_response_{timestamp}.{extension}"
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(f"Status code: {response.status_code}\n")
        f.write(f"URL: {response.url}\n")
        f.write("Headers:\n")
        for key, value in response.headers.items():
            f.write(f"  {key}: {value}\n")
        f.write("\nBody:\n")

        if is_json:
            try:
                f.write(json.dumps(response.json(), indent=2))
            except ValueError:
                f.write(response.text)
        else:
            f.write(response.text)

    print(f"\nFull response saved to: {output_path}")
    return output_path


def send_file(path):
    print(f"Sending file: {path}")
    print(f"To:           {ENDPOINT_URL}")
    print(f"Method:       {HTTP_METHOD}")

    headers = build_headers()
    auth = build_auth()

    # Open the file in binary mode and pass the file object directly as `data`.
    # requests will stream it from disk rather than loading the whole thing
    # into memory up front - this is the key difference vs. Postman, which
    # tends to load the entire body into memory (and crash on big files).
    with open(path, "rb") as f:
        response = requests.request(
            method=HTTP_METHOD,
            url=ENDPOINT_URL,
            headers=headers,
            auth=auth,
            data=f,
            timeout=TIMEOUT_SECONDS,
        )

    print(f"\nStatus code: {response.status_code}")

    # Print response body, truncated so a huge response doesn't flood the terminal
    body_preview = response.text[:2000]
    print("Response body (first 2000 chars):")
    print(body_preview)
    if len(response.text) > 2000:
        print(f"... [truncated, {len(response.text)} chars total]")

    if not response.ok:
        print(f"\nRequest failed with status {response.status_code}.")

    if SAVE_RESPONSE:
        save_response(path, response)


def main():
    file_path = FILE_PATH

    if not file_path:
        file_path = pick_file_via_dialog()
        if not file_path:
            sys.exit("No file selected. Exiting.")

    validate_json_file(file_path)
    send_file(file_path)


if __name__ == "__main__":
    main()