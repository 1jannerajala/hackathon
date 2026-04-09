import argparse
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

from email_summary_agent import run_agent

TIMEZONE = ZoneInfo("Europe/Helsinki")


def run_scheduled_summary(output_path: str, max_emails: int, client_id: str, tenant_id: str, authority: str, token_cache_path: str, client_secret: str, mailbox: str, no_env: bool):
    start_time = datetime.now(TIMEZONE).isoformat()
    print(f"[{start_time}] Starting scheduled Outlook email summary job.")
    try:
        run_agent(
            output_path=output_path,
            max_emails=max_emails,
            client_id=client_id,
            tenant_id=tenant_id,
            authority=authority,
            token_cache_path=token_cache_path,
            client_secret=client_secret,
            mailbox=mailbox,
            no_env=no_env,
        )
        print(f"[{datetime.now(TIMEZONE).isoformat()}] Job completed successfully.")
    except Exception as exc:
        print(f"[{datetime.now(TIMEZONE).isoformat()}] Job failed: {exc}")


def parse_args():
    parser = argparse.ArgumentParser(description="Schedule the Outlook email summary agent for weekdays at 07:00 Helsinki time.")
    parser.add_argument("--output", default="email_summary.md", help="Path to the Markdown summary output file.")
    parser.add_argument("--max-emails", type=int, default=None, help="Optional maximum number of emails to fetch.")
    parser.add_argument("--client-id", default=None, help="Azure app client ID. Can also be set via AZURE_CLIENT_ID.")
    parser.add_argument("--tenant-id", default=None, help="Azure tenant ID. Can also be set via AZURE_TENANT_ID.")
    parser.add_argument("--authority", default=None, help="Optional Azure authority URL. Can also be set via AZURE_AUTHORITY.")
    parser.add_argument("--client-secret", default=None, help="Azure app client secret for app-only authentication. Can also be set via AZURE_CLIENT_SECRET.")
    parser.add_argument("--mailbox", default=None, help="Service mailbox to summarize when using app-only authentication. Can also be set via SERVICE_MAILBOX.")
    parser.add_argument("--token-cache", default=None, help="Path to persistent MSAL token cache. Can also be set via TOKEN_CACHE_PATH.")
    parser.add_argument("--no-env", action="store_true", help="Do not load values from a .env file.")
    return parser.parse_args()


def main():
    args = parse_args()
    trigger = CronTrigger(day_of_week="mon-fri", hour=7, minute=0, timezone=TIMEZONE)
    scheduler = BlockingScheduler(timezone=TIMEZONE)
    scheduler.add_job(
        run_scheduled_summary,
        trigger,
        args=[
            args.output,
            args.max_emails,
            args.client_id,
            args.tenant_id,
            args.authority,
            args.token_cache,
            args.client_secret,
            args.mailbox,
            args.no_env,
        ],
        id="outlook_email_summary",
        name="Outlook email summary every weekday at 07:00 Helsinki time",
        replace_existing=True,
    )
    print("Scheduler started. The job will run on weekdays at 07:00 Europe/Helsinki.")
    scheduler.start()


if __name__ == "__main__":
    main()
