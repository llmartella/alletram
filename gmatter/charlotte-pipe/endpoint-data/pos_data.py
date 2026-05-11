import requests
import json
import time
import os
from typing import Dict, Any, Optional

def fetch_and_save_data(
    endpoint_url: str,
    output_file: str,
    headers: Optional[Dict[str, str]] = None,
    params: Optional[Dict[str, Any]] = None,
    chunk_size: int = 8192,
    max_retries: int = 3,
    timeout: int = 300  # bumped to 300 — large POS datasets can be slow to start
):
    session = requests.Session()
    if headers:
        session.headers.update(headers)

    for attempt in range(max_retries):
        try:
            print(f"Attempt {attempt + 1}/{max_retries} — fetching: {endpoint_url}")

            response = session.get(
                endpoint_url,
                params=params,
                stream=True,
                timeout=(15, 300)  # (connect timeout, read timeout) — separate concerns
            )
            response.raise_for_status()

            content_type = response.headers.get('content-type', '').lower()
            total_bytes = 0

            with open(output_file, 'w', encoding='utf-8') as f:
                if 'application/json' in content_type:
                    print("Streaming JSON response...")
                    # Stream raw chunks to file instead of loading all into memory
                    for chunk in response.iter_content(chunk_size=chunk_size, decode_unicode=True):
                        if chunk:
                            f.write(chunk)
                            total_bytes += len(chunk.encode('utf-8'))
                            if total_bytes % (chunk_size * 100) == 0:
                                print(f"  Received: {total_bytes:,} bytes")
                else:
                    print("Streaming text response...")
                    for chunk in response.iter_content(chunk_size=chunk_size, decode_unicode=True):
                        if chunk:
                            f.write(chunk)
                            total_bytes += len(chunk.encode('utf-8'))

            print(f"✅ Saved to: {output_file} ({get_file_size(output_file)})")
            return True

        except requests.exceptions.Timeout as e:
            print(f"⏱ Timeout on attempt {attempt + 1}: {e}")
        except requests.exceptions.ConnectionError as e:
            print(f"🔌 Connection error on attempt {attempt + 1}: {e}")
        except requests.exceptions.HTTPError as e:
            print(f"🚫 HTTP {response.status_code} on attempt {attempt + 1}: {response.text[:300]}")
            # Don't retry on auth errors — they won't resolve themselves
            if response.status_code in (401, 403):
                print("Auth error — check credentials. Not retrying.")
                return False
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            return False

        if attempt < max_retries - 1:
            wait_time = 5 * (2 ** attempt)  # 5s, 10s, 20s — more breathing room
            print(f"Retrying in {wait_time}s...")
            time.sleep(wait_time)

    print("Max retries reached.")
    return False


def get_file_size(file_path: str) -> str:
    """Get human-readable file size."""
    try:
        size = os.path.getsize(file_path)
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"
    except:
        return "Unknown"


def main():
    ENDPOINT_URL = "https://cpf-api.charlottepipe.com/GET_ContractorDB_POS"
    OUTPUT_FILE = "api_data.json"

    HEADERS = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
    }

    PARAMS = {
        "AuthorizationNumber": "ClairVoyant_4GV",  # fill in
        "Passcode": "258a1a49-a5b4-48fd-bf65-b549ea36143a",             # fill in
    }

    print("Starting data fetch...")
    success = fetch_and_save_data(
        endpoint_url=ENDPOINT_URL,
        output_file=OUTPUT_FILE,
        headers=HEADERS,
        params=PARAMS
    )

    if success:
        print("Script completed successfully!")
    else:
        print("Script failed — check logs above.")


if __name__ == "__main__":
    main()
