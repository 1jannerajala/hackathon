# 🚀 From Code to Impact: GitHub Copilot for AI Builders

Welcome to this day full of GitHub Copilot! Below you'll find the agenda for the day and links to workshop materials.

---

## 📅 Agenda

| Time | Session | Who | Length |
|-------|---------|-----|--------|
| 9:00 | **Arrival and Breakfast (Singapore)** | | 30 min |
| 09:30 | Welcome & Introduction | Vibha Deshpande, Microsoft | 5 min |
| 09:35 | From Code to Impact: Why GitHub Copilot Matters Now | Lillie Harris, Microsoft | 15 min |
| 9:50 | End to End Agentic Engineering | Lukas Lundin, Microsoft | 20 min |
| 10:10 | Getting started with GitHub Copilot | Raghib, Aapo, Microsoft | 20 min |
| 10:30 | Break and transfer to hack area | | 15 min |
| 10:45 | Open Hack (Customer Center) | Breakout Rooms | 75 min |
| 12:00 | Lunch (Nest) | | 60 min |
| 13:00 | Open Hack (Customer Center) | Breakout Rooms | 45 min |
| 14:15 | Coffee | | |
| 14:30 | Open Hack (Customer Center) | Breakout Rooms | 30 min |
| 15:45 | Showcase and Wrap Up | Breakout Rooms | 15 min |

---

## 🧰 Getting Started with GitHub Copilot

Before diving into the scenarios, make sure you're set up with GitHub Copilot in your tool of choice:

- 🖥️ **[GitHub Copilot in VS Code](https://github.com/features/copilot/ai-code-editor)** — AI-powered code completions, chat, and agent mode right inside your editor.
- 💻 **[GitHub Copilot in the CLI](https://github.com/features/copilot/cli/)** — Use Copilot directly from the terminal for shell commands, git operations, and more.

---

## 🎯 Hackathon Scenarios

Pick one of the three scenarios below for today's hack session:

### 1. 🌱 Greenfield — Bring Your Own Idea

Have a project idea you've been wanting to build? This is your chance! Build something from scratch with the help of GitHub Copilot. Our team will be on hand to help you along the way — just raise your hand or find us in the breakout rooms.

> **💡 Tip:** Not sure what to build? Use **Plan Mode** in GitHub Copilot (press `Shift+Tab` in the CLI or select "Plan" in VS Code Copilot Chat) to brainstorm and spar with Copilot. Describe your interests or problem domain and let Copilot help you shape an idea, outline an architecture, and break it down into actionable steps — before writing a single line of code!

### 2. 🛠️ GitHub Copilot Hands-On Workshop

Follow a guided, hands-on workshop to learn GitHub Copilot from the ground up using the CLI:

👉 **[Copilot CLI for Beginners](https://github.com/github/copilot-cli-for-beginners/)**

> **Note:** The techniques and patterns you learn in this workshop apply equally to VS Code with GitHub Copilot — completions, chat, and agent mode all work the same way. Feel free to follow along in whichever tool you prefer!

### 3. 🔄 Application Modernization with GitHub Copilot

Explore real-world app modernization scenarios — migrate, refactor, and modernize existing applications with AI assistance:

👉 **[GitHub Copilot App Modernization](https://learn.microsoft.com/en-us/azure/developer/github-copilot-app-modernization/overview)** — End-to-end guidance for modernizing applications using GitHub Copilot on Azure.

---

## Outlook Email Summary Agent

This repository now includes a Python agent that authenticates with Microsoft Outlook via OAuth and writes a Markdown summary file for your mailbox.

### Added files

- `email_summary_agent.py` — Python implementation for the summary agent
- `requirements.txt` — required packages for authentication and Graph API calls
- `.env.example` — sample environment variables for Azure app settings

### Setup

1. Register an Azure app in the Azure portal and grant it the `Mail.Read` Microsoft Graph permission.
2. Set `AZURE_CLIENT_ID` and `AZURE_TENANT_ID` in your environment or create a `.env` file based on `.env.example`.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

### Usage

Run the agent and generate a Markdown summary:

```bash
python email_summary_agent.py --output outlook_summary.md
```

Limit the number of messages if your mailbox is large:

```bash
python email_summary_agent.py --output outlook_summary.md --max-emails 500
```

Use app-only authentication for unattended enterprise scheduling:

```bash
python email_summary_agent.py --output outlook_summary.md --client-secret "$AZURE_CLIENT_SECRET" --mailbox "service-mailbox@yourdomain.com"
```

### Scheduled weekday runs at 07:00 Helsinki time

The scheduler script can run the same summary job automatically on weekdays at 07:00 Europe/Helsinki, including correct handling of DST transitions.

```bash
python email_summary_scheduler.py --output outlook_summary.md --client-secret "$AZURE_CLIENT_SECRET" --mailbox "service-mailbox@yourdomain.com"
```

Keep the scheduler process running in the background or as a service so it triggers every weekday at the local Helsinki morning hour.

### Code quality checks

This repository includes a nightly GitHub Actions CodeQL scan configured in `.github/workflows/codeql-analysis.yml`.

- Runs daily at 23:00 UTC, which is around midnight Helsinki time
- Executes on push to `main`, pull requests against `main`, and on a schedule
- Provides automated static analysis and security code scanning for the Python codebase

### Notes

- The agent supports both delegated authentication (device code flow) and unattended app-only authentication using a client secret.
- For app-only access to a service mailbox, set `AZURE_CLIENT_SECRET` and `SERVICE_MAILBOX` in your environment or `.env` file.
- App-only Graph permissions require admin consent and typically use `Mail.Read.All` for mailbox access.
- After successful sign-in or token acquisition, the access token is optionally cached in `token_cache.bin` so scheduled runs can execute silently.
- Output is written to a Markdown file containing total messages, unread count, top senders, top subjects, and sample previews.
