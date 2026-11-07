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
| `ANTHROPIC_API_KEY` | Turns on the AI scam checker and the "Goldie" practice chat. Optional — without it the checker uses built-in warning-sign rules. |
| `MODEL_CHECK`, `MODEL_CHAT` | Optional model overrides. |
| `DATA_DIR` | Where help requests are stored (default `./data`). |

## Note on storage
Help requests and bookings are saved to a JSON file. On Render's free plan that file is wiped whenever the service restarts or redeploys — fine for testing; move to a Render Postgres database or a persistent disk before real residents use it.
