import os
import re
import json
import logging
from datetime import datetime
import requests
import pandas as pd
import beautify

# ==========================================
# Step 1: Set variables for URL, Headers, and Parameters
# ==========================================
GET_URL = "https://www.ccilindia.com/archive"   # Replace with your GET URL
POST_URL = "https://www.ccilindia.com/archive?p_p_id=FPIArchival_FPIArchivalPortlet_INSTANCE_scnt&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view&p_p_resource_id=CG&p_p_cacheability=cacheLevelPage"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Referer":"https://www.ccilindia.com/archive",
    "Origin":"https://www.ccilindia.com"
}

# Prompt user for body parameter value
user_val = input("Enter Date in format <YYYY-MM-DD>: ").strip()

# Form-urlencoded body parameters (update key names to match your form)
PAYLOAD = {
    "_FPIArchival_FPIArchivalPortlet_INSTANCE_scnt_fromDate1": user_val
}

# ==========================================
# Setup Output Folder & Built-in Logging
# ==========================================
safe_folder = re.sub(r'[\\/*?:"<>|]', "_", user_val) or "output"
os.makedirs(safe_folder, exist_ok=True)

# 1. Simple Status Logger (what function did what & when)
status_logger = logging.getLogger("status_logger")
status_logger.setLevel(logging.INFO)
status_handler = logging.FileHandler(os.path.join(safe_folder, "status.log"), encoding="utf-8")
status_handler.setFormatter(logging.Formatter("%(asctime)s | %(funcName)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
status_logger.addHandler(status_handler)

# 2. Full Requests Dump Logger
req_logger = logging.getLogger("requests_logger")
req_logger.setLevel(logging.INFO)
req_handler = logging.FileHandler(os.path.join(safe_folder, "requests_dump.txt"), encoding="utf-8")
req_handler.setFormatter(logging.Formatter("%(asctime)s\n%(message)s\n" + "=" * 60, datefmt="%Y-%m-%d %H:%M:%S"))
req_logger.addHandler(req_handler)

ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def log_http_exchange(label, response):
    req = response.request
    dump = (
        f"[{label}]\n"
        f"REQUEST: {req.method} {req.url}\n"
        f"REQUEST HEADERS: {dict(req.headers)}\n"
        f"REQUEST BODY: {req.body}\n"
        f"RESPONSE STATUS: {response.status_code}\n"
        f"RESPONSE HEADERS: {dict(response.headers)}\n"
        f"RESPONSE BODY:\n{response.text}"
    )
    req_logger.info(dump)


def clean_cell(val):
    if isinstance(val, str):
        return ILLEGAL_XML_RE.sub("", val)
    if isinstance(val, float):
        return round(val, 12)
    return val


def clean_sheet_name(name):
    cleaned = re.sub(r'[\\/*?:\[\]]', "_", str(name))
    return cleaned[:31] or "Sheet"


def sanitize_dataframe(df):
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass
    return df.map(clean_cell)


def fetch_session_and_post():
    status_logger.info("Initializing session and sending GET request")
    session = requests.Session()
    session.headers.update(HEADERS)

    get_res = session.get(GET_URL, timeout=15)
    log_http_exchange("GET REQUEST", get_res)
    get_res.raise_for_status()
    status_logger.info(f"Completed GET request (Status: {get_res.status_code})")

    status_logger.info("Sending POST request with form-urlencoded payload")
    post_res = session.post(POST_URL, data=PAYLOAD, timeout=15)
    log_http_exchange("POST REQUEST", post_res)
    post_res.raise_for_status()
    status_logger.info(f"Completed POST request (Status: {post_res.status_code})")

    return post_res


def parse_and_export(post_res):
    status_logger.info("Parsing JSON response body")
    raw_text = post_res.content.decode("utf-8-sig").strip()
    outer_json = json.loads(raw_text)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = os.path.join(safe_folder, f"{user_val}_raw_archive_{timestamp}.xlsx")

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        sheets_written = 0

        for key, val in outer_json.items():
            if isinstance(val, str):
                try:
                    val = json.loads(val)
                except json.JSONDecodeError:
                    continue

            if isinstance(val, list) and len(val) > 0:
                df = sanitize_dataframe(pd.DataFrame(val))
                sheet_name = clean_sheet_name(key)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_written += 1
                print(f"Added sheet '{sheet_name}' ({len(df)} rows)")
                status_logger.info(f"Sanitized and wrote sheet '{sheet_name}' ({len(df)} rows)")

    if sheets_written > 0:
        print(f"\nSuccess! Exported {sheets_written} sheet(s) to: {output_filename}")
        beautify.beautify(output_filename)
        status_logger.info(f"Export completed successfully: {output_filename}")
    else:
        print("\nWarning: Response parsed, but no tabular lists were found to export.")
        status_logger.warning("No tabular lists found in parsed JSON; nothing exported")


def main():
    status_logger.info(f"Started execution for user input: '{user_val}'")
    try:
        post_res = fetch_session_and_post()
        parse_and_export(post_res)
        status_logger.info("Execution finished successfully")
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        status_logger.error(f"HTTP Request failed: {e}")
    except json.JSONDecodeError as e:
        print(f"JSON parsing error: {e}")
        status_logger.error(f"JSON parsing failed: {e}")
    except Exception as e:
        print(f"Error: {e}")
        status_logger.exception(f"Unexpected error: {e}")


if __name__ == "__main__":
    main()
