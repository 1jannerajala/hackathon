import argparse
import datetime
import os
import re
import sys
from collections import Counter
from pathlib import Path

import msal
import requests

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
SCOPE = ["Mail.Read"]
CLIENT_CREDENTIALS_SCOPE = ["https://graph.microsoft.com/.default"]


class EmailSummaryAgent:
    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        authority: str = None,
        max_emails: int = None,
        token_cache_path: str = None,
        client_secret: str = None,
        mailbox: str = None,
    ):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.authority = authority or f"https://login.microsoftonline.com/{tenant_id}"
        self.max_emails = max_emails
        self.client_secret = client_secret
        self.mailbox = mailbox
        self.token_cache_path = token_cache_path or "token_cache.bin"
        self.token_cache = msal.SerializableTokenCache()
        self._load_token_cache()
        self.is_app_only = bool(self.client_secret)
        self.scope = CLIENT_CREDENTIALS_SCOPE if self.is_app_only else SCOPE
        self.app = self._build_msal_app()
        self.access_token = None

    def _build_msal_app(self):
        if self.is_app_only:
            return msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret,
                token_cache=self.token_cache,
            )
        return msal.PublicClientApplication(
            self.client_id,
            authority=self.authority,
            token_cache=self.token_cache,
        )

    def _load_token_cache(self):
        cache_file = Path(self.token_cache_path)
        if cache_file.exists():
            self.token_cache.deserialize(cache_file.read_bytes())

    def _save_token_cache(self):
        cache_file = Path(self.token_cache_path)
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        cache_file.write_bytes(self.token_cache.serialize())

    def authenticate(self):
        if self.is_app_only:
            if not self.client_secret:
                raise RuntimeError("Client secret is required for app-only authentication.")
            if not self.mailbox:
                raise RuntimeError("SERVICE_MAILBOX must be provided for app-only mailbox access.")

            result = self.app.acquire_token_for_client(scopes=self.scope)
            if "access_token" not in result:
                raise RuntimeError(f"Authentication failed: {result.get('error_description', result)}")

            self.access_token = result["access_token"]
            self._save_token_cache()
            print("App-only authentication succeeded.")
            return

        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(scopes=self.scope, account=accounts[0])
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                print("Loaded token from cache silently.")
                return

        flow = self.app.initiate_device_flow(scopes=self.scope)
        if "user_code" not in flow:
            raise RuntimeError("Failed to initiate device flow. Check your Azure app registration settings.")

        print("Please authenticate with Microsoft Graph using the code below:")
        print(f"User code: {flow['user_code']}")
        print(f"Verification URL: {flow['verification_uri']}")
        print("Open the URL in your browser and paste the code, then return here.")

        result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Authentication failed: {result.get('error_description', result)}")

        self.access_token = result["access_token"]
        self._save_token_cache()
        print("Authentication succeeded.")

    def _make_request(self, url: str, params: dict = None):
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        return response.json()

    def fetch_messages(self):
        if not self.access_token:
            raise RuntimeError("Agent is not authenticated. Call authenticate() first.")

        if self.is_app_only:
            mailbox = self.mailbox
            if not mailbox:
                raise RuntimeError("SERVICE_MAILBOX is required for app-only Graph mailbox access.")
            url = f"{GRAPH_ENDPOINT}/users/{mailbox}/messages"
        else:
            url = f"{GRAPH_ENDPOINT}/me/messages"

        params = {
            "$select": "subject,receivedDateTime,from,isRead,bodyPreview,importance,flag",
            "$top": 50,
            "$orderby": "receivedDateTime desc",
        }

        messages = []
        while url:
            data = self._make_request(url, params)
            page_messages = data.get("value", [])
            messages.extend(page_messages)

            if self.max_emails and len(messages) >= self.max_emails:
                messages = messages[: self.max_emails]
                break

            url = data.get("@odata.nextLink")
            params = None

        return messages

    def summarize_messages(self, messages):
        total = len(messages)
        unread = sum(1 for msg in messages if not msg.get("isRead", True))
        senders = Counter()
        subjects = Counter()
        date_range = [None, None]
        preview_lines = []

        for msg in messages:
            sender = msg.get("from", {}).get("emailAddress", {}).get("name") or msg.get("from", {}).get("emailAddress", {}).get("address") or "Unknown sender"
            senders[sender] += 1

            subject = msg.get("subject") or "(No subject)"
            subjects[subject] += 1

            received = msg.get("receivedDateTime")
            if received:
                received_ts = datetime.datetime.fromisoformat(received.replace("Z", "+00:00"))
                if date_range[0] is None or received_ts < date_range[0]:
                    date_range[0] = received_ts
                if date_range[1] is None or received_ts > date_range[1]:
                    date_range[1] = received_ts

            preview = msg.get("bodyPreview", "").strip()
            preview = re.sub(r"\s+", " ", preview)
            preview_lines.append((received, sender, subject, preview[:180]))

        summary = {
            "total_messages": total,
            "unread_messages": unread,
            "date_range_start": date_range[0],
            "date_range_end": date_range[1],
            "top_senders": senders.most_common(12),
            "top_subjects": subjects.most_common(10),
            "preview_messages": preview_lines[:10],
        }
        return summary

    def write_markdown(self, summary, messages, output_path):
        lines = [
            "# Outlook Email Account Summary",
            "",
            f"Generated: {datetime.datetime.now(datetime.timezone.utc).astimezone().isoformat()}",
            "",
            f"- Total messages fetched: **{summary['total_messages']}**",
            f"- Unread messages: **{summary['unread_messages']}**",
        ]

        if summary["date_range_start"] and summary["date_range_end"]:
            lines.append(f"- Date range: **{summary['date_range_start'].isoformat()}** to **{summary['date_range_end'].isoformat()}**")

        lines.extend(["", "## Top senders", ""])
        if summary["top_senders"]:
            for sender, count in summary["top_senders"]:
                lines.append(f"- {sender}: {count}")
        else:
            lines.append("No senders found.")

        lines.extend(["", "## Top subjects", ""])
        if summary["top_subjects"]:
            for subject, count in summary["top_subjects"]:
                escaped = subject.replace("\n", " ").strip()
                lines.append(f"- {escaped}: {count}")
        else:
            lines.append("No subjects found.")

        lines.extend(["", "## Sample recent messages", ""])
        for received, sender, subject, preview in summary["preview_messages"]:
            received_display = received or "Unknown date"
            lines.append(f"### {subject}")
            lines.append(f"- Sender: {sender}")
            lines.append(f"- Received: {received_display}")
            lines.append(f"- Preview: {preview}")
            lines.append("")

        lines.append("## Fetch options")
        lines.append(f"- Max emails requested: {self.max_emails or 'unlimited'}")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        print(f"Summary written to: {output_path}")


def load_env_file():
    if load_dotenv:
        load_dotenv()


def run_agent(
    output_path: str = "email_summary.md",
    max_emails: int = None,
    client_id: str = None,
    tenant_id: str = None,
    authority: str = None,
    token_cache_path: str = None,
    client_secret: str = None,
    mailbox: str = None,
    no_env: bool = False,
):
    if not no_env:
        load_env_file()

    client_id = client_id or get_env_variable("AZURE_CLIENT_ID")
    tenant_id = tenant_id or get_env_variable("AZURE_TENANT_ID")
    authority = authority or get_env_variable("AZURE_AUTHORITY")
    token_cache_path = token_cache_path or get_env_variable("TOKEN_CACHE_PATH") or "token_cache.bin"
    client_secret = client_secret or get_env_variable("AZURE_CLIENT_SECRET")
    mailbox = mailbox or get_env_variable("SERVICE_MAILBOX")
    max_emails = max_emails or (get_env_variable("MAX_EMAILS") and int(get_env_variable("MAX_EMAILS")))

    if not client_id or not tenant_id:
        raise ValueError("AZURE_CLIENT_ID and AZURE_TENANT_ID must be provided via flags or environment variables.")

    agent = EmailSummaryAgent(
        client_id=client_id,
        tenant_id=tenant_id,
        authority=authority,
        max_emails=max_emails,
        token_cache_path=token_cache_path,
        client_secret=client_secret,
        mailbox=mailbox,
    )
    agent.authenticate()
    messages = agent.fetch_messages()
    summary = agent.summarize_messages(messages)
    agent.write_markdown(summary, messages, output_path)


def get_env_variable(name, fallback=None):
    value = os.getenv(name)
    return value if value is not None else fallback


def parse_args():
    parser = argparse.ArgumentParser(description="Outlook email summary agent using Microsoft Graph OAuth.")
    parser.add_argument("--output", default="email_summary.md", help="Path to the Markdown summary output file.")
    parser.add_argument("--max-emails", type=int, default=None, help="Optional maximum number of emails to fetch.")
    parser.add_argument("--client-id", default=None, help="Azure app client ID. Can also be set via AZURE_CLIENT_ID.")
    parser.add_argument("--tenant-id", default=None, help="Azure tenant ID. Can also be set via AZURE_TENANT_ID.")
    parser.add_argument("--authority", default=None, help="Optional Azure authority URL. Can also be set via AZURE_AUTHORITY.")
    parser.add_argument("--client-secret", default=None, help="Azure app client secret for app-only authentication. Can also be set via AZURE_CLIENT_SECRET.")
    parser.add_argument("--mailbox", default=None, help="Service mailbox to summarize when using app-only authentication. Can also be set via SERVICE_MAILBOX.")
    parser.add_argument("--token-cache", default=None, help="Path to a persistent MSAL token cache file. Can also be set via TOKEN_CACHE_PATH.")
    parser.add_argument("--no-env", action="store_true", help="Do not load values from a .env file.")
    return parser.parse_args()


def main():
    args = parse_args()
    try:
        run_agent(
            output_path=args.output,
            max_emails=args.max_emails,
            client_id=args.client_id,
            tenant_id=args.tenant_id,
            authority=args.authority,
            token_cache_path=args.token_cache,
            client_secret=args.client_secret,
            mailbox=args.mailbox,
            no_env=args.no_env,
        )
    except ValueError as exc:
        print(f"Error: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
