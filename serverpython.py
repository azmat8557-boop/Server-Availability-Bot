"""
Robot Framework integration layer.

Robot controls the browser (Selenium). This module handles the "data work":
- Read rows from `input.xlsx`
- Write the scraped `Status` back into `input.xlsx`
- Send one summary email via Outlook at the end of the run
"""

import pandas as pd
import win32com.client

RECEIVER_EMAIL = "[Email Address]"   # Change this to your receiver address


def get_server_data():
    """
    Reads `input.xlsx` and returns rows as Python dictionaries.

    The website dropdown value is not just the server code or just the IP.
    This function generates `ServerOption` to match the dropdown format.
    Example: Server Code 'TS' + IP '202.32.5' -> 'TS-202-32-5'
    """
    df = pd.read_excel("input.xlsx")
    
    # Build the combined dropdown value.
    # The website uses ALL dashes: TS + 202.32.5 -> TS-202-32-5 (dots become dashes).
    df['ServerOption'] = df['Server Code'].astype(str) + '-' + df['IP'].astype(str).str.replace('.', '-', regex=False)
    
    return df.to_dict('records')

def save_server_status(code, status):
    """
    Updates `input.xlsx` for one server code with the scraped `Status`.

    This function is intentionally defensive: if something goes wrong while writing,
    it records an ERROR message into the `Status` column and continues.
    """
    try:
        df = pd.read_excel("input.xlsx")
        df['Status'] = df['Status'].astype(str)
        
        # The page sometimes returns multi-line text like: "ServerName\\nOnline".
        # We keep the last line as the final "Online/Offline" style result.
        clean_status = status.strip().split('\n')[-1].strip()
        
        df.loc[df['Server Code'] == code, 'Status'] = clean_status
        df.to_excel("input.xlsx", index=False)
        print(f"✅ Saved: {code} → {clean_status}")
    except Exception as e:
        # If anything goes wrong, record it into the row and keep the run going.
        df.loc[df['Server Code'] == code, 'Status'] = f"ERROR: {e}"
        df.to_excel("input.xlsx", index=False)
        print(f"⚠️ Row {code} had an issue: {e} — kept running!")


def send_status_email(max_rows=0):
    """
    Sends ONE summary email summarizing scraped `Status` values from `input.xlsx`.

    max_rows:
      - 0 or negative: send all rows
      - positive: only send the first N rows (useful for testing)
    """
    try:
        df = pd.read_excel("input.xlsx")

        if "Server Code" not in df.columns or "Status" not in df.columns:
            print("⚠️ Email not sent: required columns missing in input.xlsx")
            return

        # Normalize and handle blanks safely so email building never crashes.
        df["Status"] = df["Status"].fillna("").astype(str)

        # Optional testing limit
        try:
            max_rows_int = int(max_rows)
        except Exception:
            max_rows_int = 0

        if max_rows_int > 0:
            df = df.head(max_rows_int)

        # Keep only rows that actually have a status value.
        df = df[df["Status"].str.strip() != ""]

        if df.empty:
            print("⚠️ Email not sent: no statuses found in input.xlsx")
            return

        lines = []
        for _, row in df.iterrows():
            # Expected row fields:
            # - Server Code
            # - Status
            # - IP (optional, but we include it when present)
            server_code = str(row.get("Server Code", "")).strip()
            status = str(row.get("Status", "")).strip()

            # Include IP if available to make the email easier to verify.
            ip = ""
            if "IP" in df.columns:
                ip = str(row.get("IP", "")).strip()
            if ip:
                lines.append(f"- {server_code} ({ip}): {status}")
            else:
                lines.append(f"- {server_code}: {status}")

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem
        mail.To = RECEIVER_EMAIL
        mail.Subject = f"Server Status Alert: {len(df)} servers"

        body_lines = [
            "Hello,",
            "",
            "Automated Status Notification:",
            "",
        ]
        body_lines.extend(lines)
        body_lines.extend(["", "Regards,", "Automated Bot"])

        mail.Body = "\n".join(body_lines)

        mail.Send()
        print(f"✅ Email sent! Rows: {len(df)}")
    except Exception as e:
        # Do not break the Robot run if email fails
        print(f"⚠️ Email send failed: {e}")