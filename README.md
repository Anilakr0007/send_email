# Excel Candidate Email Sender

This project reads candidate rows from an Excel file, sends a personalized email to each valid email address, and records every attempt in a CSV log.

## What you need to change first

1. **SMTP settings in `.env`**
   - `SMTP_HOST`: your mail provider SMTP server.
   - `SMTP_PORT`: usually `587` for STARTTLS or `465` for SSL.
   - `SMTP_USERNAME`: the mailbox used to send emails.
   - `SMTP_PASSWORD`: an app password or SMTP password. Do not commit this value.
   - `SENDER_NAME`: name recipients should see.
   - `SENDER_EMAIL`: normally the same as `SMTP_USERNAME`.
2. **Email content in `templates/candidate_email.txt`**
   - Replace the sample subject and body with the actual message.
   - `{candidate_name}` is replaced from the Excel sheet.
   - `{contact_no}` is replaced from the Excel sheet.
3. **Excel input file**
   - Put one `.xlsx` or `.xls` file in `source/`.
   - Required column names: `Candidate Name`, `Contact No`, `Email`.
   - The header names are matched case-insensitively and extra columns are allowed.
   - Keep the email column populated with one recipient email per row.
4. **Optional folders/settings**
   - `source/`: incoming Excel files.
   - `archive/`: successfully processed Excel files when archive mode is enabled.
   - `logs/email_log.csv`: send history.
   - `attachments/`: optional files to attach to every email. Set `ATTACHMENTS` in `.env` if needed.

## Important email-provider setup

- Gmail and Microsoft 365 commonly require an app password or approved SMTP authentication; a normal account password may not work.
- Confirm that your organization permits automated SMTP sending.
- Start with `DRY_RUN=true`. This validates the workbook and prints recipients without sending.
- The default behavior skips an email address already recorded as `SENT` in the log. Set `RESEND_SENT=true` only when you intentionally want to send again.
- Never commit `.env`, passwords, or candidate data to source control.

## Setup on Windows

Open PowerShell in this folder:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Edit `.env`, place the Excel file in `source/`, and customize the template.

## Run

Validate without sending:

```powershell
python send_candidate_emails.py --dry-run
```

Send emails:

```powershell
python send_candidate_emails.py
```

Use a specific workbook:

```powershell
python send_candidate_emails.py --file source\candidates.xlsx
```

Archive the workbook after a successful run:

```powershell
python send_candidate_emails.py --archive
```

## Log details

`logs/email_log.csv` records timestamp, candidate name, contact number, email, source file, status, and an error message when applicable. Status values are `SENT`, `SKIPPED`, `INVALID`, or `FAILED`.

The log is an operational record only. Protect it because it contains personal information.

## Excel example

| Candidate Name | Contact No | Email |
|---|---|---|
| Priya Kumar | 9876543210 | priya@example.com |
| Arjun Rao | 9123456780 | arjun@example.com |

## Project files

- `send_candidate_emails.py`: main program.
- `.env.example`: configuration template.
- `templates/candidate_email.txt`: subject and body template.
- `requirements.txt`: Python dependencies.
- `source/`, `archive/`, `logs/`, `attachments/`: runtime folders.
