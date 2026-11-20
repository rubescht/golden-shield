# Golden Shield — by Code Black Group

A simple, senior-friendly app for aged care residents: cyber safety lessons, AI made simple, a scam checker, classes, trusted videos and an "I Need Help" button that reaches the Code Black Group team.

## Run locally
    npm install
    STAFF_PASSWORD=choose-one npm start
    # open http://localhost:3000   (staff dashboard: /#staff)

## Environment variables (set in Render)
| Name | What it does |
|---|---|
| `STAFF_PASSWORD` | Password for the staff dashboard (`/#staff`). Required. |
| `TRIAL_PASSWORD` | Password clients type on the sign-in screen. Default `GoldenShield2026`. Checked on the server, never sent to the browser. |
| `NOTIFY_TO` | Who gets the alert emails. Default `vip@codeblackgroup.com.au` (comma-separate for more). |
| `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_SENDER` | **Recommended.** Sends alerts from a Microsoft 365 mailbox through Microsoft Graph (works on Render's free plan, which blocks SMTP). Needs an Entra app registration with the `Mail.Send` application permission and admin consent. `MS_SENDER` is the mailbox it sends from. |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS` | Plain SMTP instead. Does not work on Render's free plan (SMTP ports are blocked there). |
| `RESEND_API_KEY` | Alternative to SMTP: send alerts through Resend instead. |
| `MAIL_FROM` | Optional sender, e.g. `Golden Shield <noreply@codeblackgroup.com.au>`. Defaults to the SMTP user. |
| `ANTHROPIC_API_KEY` | Turns on the AI scam checker and the "Goldie" practice chat. Optional — without it the checker uses built-in warning-sign rules. |
| `MODEL_CHECK`, `MODEL_CHAT` | Optional model overrides. |
| `DATA_DIR` | Where help requests are stored (default `./data`). |

## Email alerts
Every trial sign-in, help request and class booking is emailed to `NOTIFY_TO` the moment it happens, signed off "Golden Shield Dashboard". If sending fails, the email is retried every 5 minutes. The staff dashboard shows whether email is set up and has a "Send a test email" button.

## Note on storage
Help requests and bookings are saved to a JSON file. On Render's free plan that file is wiped whenever the service restarts or redeploys — fine for testing; move to a Render Postgres database or a persistent disk before real residents use it.
