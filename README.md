# Server Availability Bot
This bot uses Robot Framework + Python to check server availability on https://botsdna.com/ServerAvailability/ and send you one summary email with the results (via Outlook on the same PC)

## Run
1. Install dependencies: `pip install -r requirements.txt`
2. Start: `robot Serverrobot.robot`

## What it does
- Downloads `input.xlsx`
- For each row: logs in (UID/PWD), selects the server, starts it, and reads the page `Status` in the web broswer.
- Writes the `Status` back into `input.xlsx`
- Sends **one summary email** at the end, then closes the browser

## `input.xlsx` columns required
- `UID`, `PWD`, `Server Code`, `IP`, `Status`
(`Server---Option---` is generated automatically by the Python code.)

## Email rows (test control)
In `Serverrobot.robot`, it is set at:
- `${MAX_EMAIL_ROWS}` (default `5`)


## Main files
- `Serverrobot.robot`, `serverpython.py`, `requirements.txt`
