# 🤖 360° Feedback Slack Bot (Serverless)

A powerful, serverless HR tool for collecting 360-degree feedback directly in Slack. Built with **Google Apps Script** and **Google Sheets** as a database.

![Приклад роботи бота, стартовий меседж для респондента]([https://github.com/IhorFedorov/360-feedback-slack-bot/blob/7489fbb08f4347a56404075cc75fe93802733df0/bot-start-screen.png))

## ✨ Key Features

* **Slack-First Experience:** Interactive buttons, modal reports, and direct messages.
* **Smart Deadlines:** Automatically calculates due dates (e.g., +3 working days) excluding weekends.
* **Snooze Logic:** Users can delay feedback (10m, 1h, 1d) with automated follow-ups.
* **Spam Protection:** Smart logic prevents duplicate requests for the same person within 30 days.
* **Admin Dashboard:**
    * View real-time stats in Slack.
    * Send urgent reminders ("🔥 Nudge") to pending users.
    * Generate sidebar reports directly in Google Sheets.
* **Email Reports:** Generates a formatted summary email (Draft) in Gmail with one click.

## 🛠 Tech Stack

* **Backend:** Google Apps Script (Node.js-like environment).
* **Database:** Google Sheets.
* **UI:** Slack Block Kit.
* **Report View:** HTML/CSS (served via GAS Web App).

## 🚀 Installation & Setup

### 1. Google Sheets Setup
1.  Create a new Google Sheet.
2.  Rename the active sheet to `Database`.
3.  Create the header row (Row 1) with these columns:
    `User ID | Subject | Token | Status | Date Created | Date Completed | Last Opened | Snooze Until | Q1 | Q2 ...`

### 2. Google Apps Script
1.  Go to **Extensions** > **Apps Script**.
2.  Copy the contents of `Code.js`, `index.html`, and `report.html` from this repository.
3.  **Crucial:** Update the `CONFIGURATION` section in `Code.js`:
    ```javascript
    const SLACK_TOKEN = 'xoxb-your-bot-token';
    const SPREADSHEET_ID = 'your-sheet-id';
    const WEB_APP_URL = 'your-deployed-web-app-url';
    const ADMIN_IDS = ['U12345', 'U67890']; // Your Slack User IDs
    ```

### 3. Slack App Configuration
1.  Create a new app at [api.slack.com](https://api.slack.com/apps).
2.  **OAuth & Permissions:** Add the following scopes:
    * `chat:write`
    * `commands`
    * `im:write`
    * `users:read`
3.  **Interactivity & Shortcuts:** Enable it and paste your **Web App URL** (from Step 4).

### 4. Deployment
1.  In the Script Editor, click **Deploy** > **New Deployment**.
2.  Select type: **Web App**.
3.  **Execute as:** `Me`.
4.  **Who has access:** `Anyone`.
5.  Copy the generated URL and update `WEB_APP_URL` in your code and Slack settings.

## 📸 Screenshots

### Request & Snooze Options
*(Insert screenshot of the bot message with "Snooze" buttons)*

### Admin Report in Slack
*(Insert screenshot of the "Status" card with stats)*

### Gmail Summary Draft
*(Insert screenshot of the generated email)*

## 🤝 Contributing
Feel free to fork this repository and submit Pull Requests.

## 📄 License
MIT License.
