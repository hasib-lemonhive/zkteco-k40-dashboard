# ZKTeco SQLite Dashboard

A simple Flask dashboard for the `attendance.db` created by the K40 sync script.

## Setup

Place this project in the same directory as `attendance.db`.

Create/activate your virtual environment, then from the root `zkteco` directory:

```bash
pip install -r requirements.txt
```

Run:

```bash
python app.py
```

Open:

http://localhost:5000

The dashboard has:

- `/` — attendance with user name and pagination
- `/users` — users with pagination
