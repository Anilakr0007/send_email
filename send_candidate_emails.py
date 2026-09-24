from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import smtplib
import ssl
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

REQUIRED_COLUMNS = {"candidate name", "contact no", "email"}
LOG_COLUMNS = [
    "timestamp_utc",
    "candidate_name",
    "contact_no",
    "email",
    "source_file",
    "status",
    "error",
]
EMAIL_PATTERN = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


@dataclass(frozen=True)
class Settings:
    smtp_host: str
    smtp_port: int
    smtp_username: str
    smtp_password: str
    smtp_use_tls: bool
    sender_name: str
    sender_email: str
    source_dir: Path
    archive_dir: Path
    log_file: Path
    template_file: Path
    attachments: tuple[Path, ...]
    dry_run: bool
    resend_sent: bool


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def load_settings() -> Settings:
    load_dotenv()
    attachments = tuple(
        Path(item.strip())
        for item in os.getenv("ATTACHMENTS", "").split(",")
        if item.strip()
    )
    return Settings(
        smtp_host=os.getenv("SMTP_HOST", ""),
        smtp_port=int(os.getenv("SMTP_PORT", "587")),
        smtp_username=os.getenv("SMTP_USERNAME", ""),
        smtp_password=os.getenv("SMTP_PASSWORD", ""),
        smtp_use_tls=env_bool("SMTP_USE_TLS", True),
        sender_name=os.getenv("SENDER_NAME", ""),
        sender_email=os.getenv("SENDER_EMAIL", os.getenv("SMTP_USERNAME", "")),
        source_dir=Path(os.getenv("SOURCE_DIR", "source")),
        archive_dir=Path(os.getenv("ARCHIVE_DIR", "archive")),
        log_file=Path(os.getenv("LOG_FILE", "logs/email_log.csv")),
        template_file=Path(
            os.getenv("TEMPLATE_FILE", "templates/candidate_email.txt")
        ),
        attachments=attachments,
        dry_run=env_bool("DRY_RUN", True),
        resend_sent=env_bool("RESEND_SENT", False),
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send personalized emails from an Excel workbook."
    )
    parser.add_argument(
        "--file",
        type=Path,
        help="Workbook path. Defaults to the newest .xlsx or .xls in SOURCE_DIR.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and print recipients without sending email.",
    )
    parser.add_argument(
        "--archive",
        action="store_true",
        help="Move the workbook to ARCHIVE_DIR after a run with no failures.",
    )
    return parser.parse_args()


def choose_workbook(source_dir: Path, requested: Path | None) -> Path:
    if requested:
        if not requested.is_file():
            raise FileNotFoundError(f"Workbook not found: {requested}")
        return requested

    workbooks = [
        path
        for path in source_dir.glob("*")
        if path.suffix.lower() in {".xlsx", ".xls"} and not path.name.startswith("~$")
    ]
    if not workbooks:
        raise FileNotFoundError(f"No .xlsx or .xls file found in {source_dir}")
    return max(workbooks, key=lambda path: path.stat().st_mtime)


def read_template(path: Path) -> tuple[str, str]:
    if not path.is_file():
        raise FileNotFoundError(f"Template not found: {path}")
    lines = path.read_text(encoding="utf-8").splitlines()
    if not lines or not lines[0].lower().startswith("subject:"):
        raise ValueError("The first template line must start with 'Subject:'")
    subject = lines[0].split(":", 1)[1].strip()
    body = "\n".join(lines[1:]).lstrip("\n")
    if not subject or not body.strip():
        raise ValueError("The template must contain a subject and a non-empty body")
    return subject, body


def normalized_columns(dataframe: pd.DataFrame) -> dict[str, str]:
    return {
        str(column).strip().lower(): str(column)
        for column in dataframe.columns
    }


def read_candidates(workbook: Path) -> pd.DataFrame:
    dataframe = pd.read_excel(workbook, dtype=str).fillna("")
    columns = normalized_columns(dataframe)
    missing = REQUIRED_COLUMNS - columns.keys()
    if missing:
        missing_text = ", ".join(sorted(missing))
        raise ValueError(f"Missing required Excel columns: {missing_text}")

    return dataframe.rename(
        columns={
            columns["candidate name"]: "candidate_name",
            columns["contact no"]: "contact_no",
            columns["email"]: "email",
        }
    )


def read_sent_emails(log_file: Path) -> set[str]:
    if not log_file.is_file():
        return set()
    try:
        log = pd.read_csv(log_file, dtype=str).fillna("")
    except (pd.errors.EmptyDataError, pd.errors.ParserError):
        return set()
    if "status" not in log or "email" not in log:
        return set()
    return {
        str(email).strip().lower()
        for email, status in zip(log["email"], log["status"])
        if str(status).strip().upper() == "SENT"
    }


def append_log(log_file: Path, values: dict[str, str]) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    is_new = not log_file.exists()
    with log_file.open("a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=LOG_COLUMNS)
        if is_new:
            writer.writeheader()
        writer.writerow(values)


def render(value: str, candidate_name: str, contact_no: str) -> str:
    return value.format(candidate_name=candidate_name, contact_no=contact_no)


def build_message(
    sender_name: str,
    sender_email: str,
    recipient: str,
    subject: str,
    body: str,
    attachments: tuple[Path, ...],
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    for attachment in attachments:
        if not attachment.is_file():
            raise FileNotFoundError(f"Attachment not found: {attachment}")
        message.add_attachment(
            attachment.read_bytes(),
            maintype="application",
            subtype="octet-stream",
            filename=attachment.name,
        )
    return message


def open_smtp(settings: Settings) -> smtplib.SMTP:
    if settings.smtp_port == 465:
        server: smtplib.SMTP = smtplib.SMTP_SSL(
            settings.smtp_host, settings.smtp_port, context=ssl.create_default_context()
        )
    else:
        server = smtplib.SMTP(settings.smtp_host, settings.smtp_port)
        if settings.smtp_use_tls:
            server.starttls(context=ssl.create_default_context())
    server.login(settings.smtp_username, settings.smtp_password)
    return server


def log_values(
    candidate_name: str,
    contact_no: str,
    email: str,
    workbook: Path,
    status: str,
    error: str = "",
) -> dict[str, str]:
    return {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "candidate_name": candidate_name,
        "contact_no": contact_no,
        "email": email,
        "source_file": str(workbook),
        "status": status,
        "error": error,
    }


def main() -> int:
    args = parse_args()
    settings = load_settings()
    dry_run = settings.dry_run or args.dry_run
    workbook = choose_workbook(settings.source_dir, args.file)
    subject_template, body_template = read_template(settings.template_file)
    candidates = read_candidates(workbook)
    sent_emails = set() if settings.resend_sent else read_sent_emails(settings.log_file)

    print(f"Workbook: {workbook}")
    print(f"Candidates found: {len(candidates)}")
    if dry_run:
        print("DRY RUN: no email will be sent")

    smtp: smtplib.SMTP | None = None
    failures = 0
    sent_count = 0
    try:
        if not dry_run:
            missing_settings = [
                name
                for name, value in {
                    "SMTP_HOST": settings.smtp_host,
                    "SMTP_USERNAME": settings.smtp_username,
                    "SMTP_PASSWORD": settings.smtp_password,
                    "SENDER_EMAIL": settings.sender_email,
                }.items()
                if not value
            ]
            if missing_settings:
                raise ValueError(
                    "Missing SMTP configuration: " + ", ".join(missing_settings)
                )
            smtp = open_smtp(settings)

        for row_number, row in candidates.iterrows():
            candidate_name = str(row["candidate_name"]).strip()
            contact_no = str(row["contact_no"]).strip()
            email = str(row["email"]).strip()
            common = dict(
                candidate_name=candidate_name,
                contact_no=contact_no,
                email=email,
                workbook=workbook,
            )
            if not email or not EMAIL_PATTERN.fullmatch(email):
                append_log(settings.log_file, log_values(**common, status="INVALID", error="Invalid or missing email address"))
                failures += 1
                print(f"Row {row_number + 2}: invalid email ({email or 'blank'})")
                continue
            if email.lower() in sent_emails and not settings.resend_sent:
                append_log(settings.log_file, log_values(**common, status="SKIPPED", error="Already recorded as SENT"))
                print(f"Row {row_number + 2}: skipped {email} (already sent)")
                continue

            message = build_message(
                settings.sender_name,
                settings.sender_email,
                email,
                render(subject_template, candidate_name, contact_no),
                render(body_template, candidate_name, contact_no),
                settings.attachments,
            )
            if dry_run:
                print(f"Row {row_number + 2}: would send to {email}")
                continue

            try:
                smtp.send_message(message)
                append_log(settings.log_file, log_values(**common, status="SENT"))
                sent_emails.add(email.lower())
                sent_count += 1
                print(f"Row {row_number + 2}: sent to {email}")
            except Exception as error:  # SMTP errors vary by provider.
                append_log(settings.log_file, log_values(**common, status="FAILED", error=str(error)))
                failures += 1
                print(f"Row {row_number + 2}: failed for {email}: {error}")
    finally:
        if smtp is not None:
            smtp.quit()

    if args.archive and not dry_run and failures == 0:
        settings.archive_dir.mkdir(parents=True, exist_ok=True)
        destination = settings.archive_dir / workbook.name
        if destination.exists():
            destination = settings.archive_dir / f"{workbook.stem}_{datetime.now():%Y%m%d%H%M%S}{workbook.suffix}"
        shutil.move(str(workbook), destination)
        print(f"Archived workbook: {destination}")

    print(f"Completed. Sent: {sent_count}; failures/invalid: {failures}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
