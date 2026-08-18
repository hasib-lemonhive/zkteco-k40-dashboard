# ZKTeco K40 Attendance Sync & Web Dashboard

An end-to-end attendance management system and web dashboard for **ZKTeco K40** biometric fingerprint/card attendance devices. 

It connects to the ZKTeco device over LAN (via the ZK communication protocol), syncs user information and punch logs into a local SQLite database (`attendance.db`), and provides a clean, responsive web dashboard with **Real-Time Live Event Streaming**, **HR Analytics & KPI Cards**, **Daily In & Out Summaries**, **Excel Exports**, and **Database Configuration & Maintenance Tools**.

---

## 📑 Table of Contents
- [System Architecture & File Tree](#-system-architecture--file-tree)
- [Key Features](#-key-features)
- [Database Schema & Persistence Guarantee](#-database-schema--persistence-guarantee)
- [Prerequisites](#-prerequisites)
- [Installation & Setup](#-installation--setup)
- [How to Run](#-how-to-run)
  - [Option A: Windows 1-Click Launcher (Recommended for End Users)](#option-a-windows-1-click-launcher-recommended-for-end-users)
  - [Option B: Manual Terminal Execution (Mac / Linux / Windows)](#option-b-manual-terminal-execution-mac--linux--windows)
- [Hardware & Network Configuration](#-hardware--network-configuration)
- [Automated Testing & Verification](#-automated-testing--verification)
- [Developer Guide (Extending the Project)](#-developer-guide-extending-the-project)

---

## 🏗 System Architecture & File Tree

```text
zkteco/
├── attendance.db               # SQLite database (permanent local storage)
├── requirements.txt            # Centralized project dependencies
├── run_dashboard.bat           # 1-Click Windows batch launcher (Sync + Web UI)
├── sync_k40.py                 # Device communication, historical sync & live listener
├── README.md                   # Project documentation
│
└── dashboard/
    ├── app.py                  # Flask web application & export engine
    ├── test_app.py             # Automated unit & integration test suite (16 tests)
    │
    ├── static/
    │   └── style.css           # Modern, responsive CSS design
    │
    └── templates/
        ├── base.html           # Base layout template with navbar & flash alerts
        ├── attendance.html     # Attendance table, KPI cards & export tools
        ├── users.html          # Registered users list with pagination
        ├── config.html         # Configuration, device stats & deletion tools
        └── _pagination.html    # Reusable query-preserving pagination component
```

---

## ✨ Key Features

### 1. ⚡ Real-Time Live Capture
- Listens to device punch events in **real time** using `pyzk` (`conn.live_capture()`).
- As soon as an employee scans their finger on the machine, the punch is inserted into `attendance.db` within milliseconds.
- Auto-reconnection mechanism ensures resilience against network drops.

### 2. 📊 HR Analytics & KPI Cards
- Displayed at the top of the Attendance page:
  - 👥 **Total Staff**: Total registered employee count.
  - 🟢 **Present Today**: Count of active employees today with percentage attendance rate.
  - 🔴 **Absent Today**: Count of absent employees with an interactive expandable list (**"View absent staff"**) listing their names.

### 3. 🕒 Daily In & Out Summaries
- Aggregates multiple daily punches into a single row per employee per day:
  - **Date**
  - **Employee Name**
  - **In Time (First)**: First punch of the day formatted in 12-hour AM/PM (`10:36:34 AM`).
  - **Out Time (Last)**: Last punch of the day, or `- (Single Punch)` if only one punch occurred.
  - **Total Hours**: Calculated elapsed work duration (`Xh Ym`, e.g. `4h 07m`) or `N/A` for single punches.
  - **Total Punches & Punch History Details**.

### 4. 🔍 Flexible Filtering & Pagination
- **Filter by User**: Dropdown list of all registered employees.
- **Filter by Date Range**: Optional `From` and `To` date pickers.
- **Query-Preserving Pagination**: Retains active filters across page navigation.

### 5. 📥 Excel Export (`.xlsx`)
- Custom spreadsheet generator built with `openpyxl`:
  - Formatted dark blue headers, light borders, centered time columns, and auto-fitted column widths.
  - **Ascending Date Sorting**: Exports records chronologically from oldest date to newest date.
  - **1-Click Quick Preset Buttons**:
    - **Export Excel**: Exports active filter scope (or defaults to the latest month's data).
    - **Last Week**: Exports records for the last 7 days with date range in filename (`attendance_last_week_YYYY-MM-DD_to_YYYY-MM-DD.xlsx`).
    - **Last Month**: Exports records for the previous calendar month (`attendance_last_month_YYYY-MM-DD_to_YYYY-MM-DD.xlsx`).

### 6. ⚙️ Configuration & Maintenance Page (`/config`)
- **Side-by-Side Comparison**:
  - **Local Database**: Records count, registered users, SQLite storage file size.
  - **ZKTeco K40 Device**: Online/Offline status, records in memory vs. 80,000 capacity, users and fingerprints count.
- **Selective Attendance Deletion**:
  - Delete attendance records **By Month** (e.g., `2025-08`) or **By Year** (e.g., `2025`).
  - Mandatory confirmation warning dialogs before deletion.
  - Automatically runs SQLite `VACUUM` to compact storage and reclaim disk space.

---

## 🗄 Database Schema & Persistence Guarantee

The system uses SQLite located at `attendance.db`:

```sql
CREATE TABLE IF NOT EXISTS users (
    user_id TEXT PRIMARY KEY,
    name TEXT,
    privilege TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS attendance (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL,
    timestamp TEXT NOT NULL,
    status INTEGER,
    punch INTEGER,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (user_id, timestamp, status, punch),
    FOREIGN KEY (user_id) REFERENCES users(user_id)
);
```

### 🛡 Data Retention & Safety Guarantee:
`sync_k40.py` uses `INSERT OR IGNORE INTO attendance` and `INSERT INTO users ... ON CONFLICT / UPDATE`.  
**Existing records in `attendance.db` are never deleted when syncing from the device.** If attendance logs or users are cleared from the physical device memory, your local database permanently preserves all historical data.

---

## 📦 Prerequisites

1. **Python 3.10 or higher** installed ([python.org](https://www.python.org/downloads/)).
   - *On Windows, ensure "Add Python to PATH" is checked during installation.*
2. **Network Connection**: The computer running the script must be on the same local network (LAN / Wi-Fi) as the ZKTeco K40 device.

---

## 🚀 Installation & Setup

1. Clone or copy the project repository:
   ```bash
   cd zkteco
   ```

2. Create a Python virtual environment:
   ```bash
   python -m venv .venv
   ```

3. Activate the virtual environment:
   - **Windows (cmd.exe)**:
     ```cmd
     .venv\Scripts\activate.bat
     ```
   - **Windows (PowerShell)**:
     ```powershell
     .venv\Scripts\Activate.ps1
     ```
   - **Mac / Linux**:
     ```bash
     source .venv/bin/activate
     ```

4. Install dependencies from the project root:
   ```bash
   pip install -r requirements.txt
   ```

---

## 🖥 How to Run

### Option A: Windows 1-Click Launcher (Recommended for End Users)
Simply **double-click [`run_dashboard.bat`](file:///Users/ranger/development/zkteco/run_dashboard.bat)**.

It automatically:
1. Verifies/creates `.venv` and installs packages from `requirements.txt`.
2. Starts the background **Live Listener** (`sync_k40.py`).
3. Starts the **Flask Web Dashboard** (`dashboard/app.py`).
4. Opens `http://localhost:5000` in the default browser.

---

### Option B: Manual Terminal Execution (Mac / Linux / Windows)

#### 1. Run Device Sync & Live Listener
```bash
# Sync history + listen for live punch events continuously:
python sync_k40.py

# Or perform one-time sync only and exit:
python sync_k40.py --once
```

#### 2. Start the Web Dashboard
In another terminal window:
```bash
python dashboard/app.py
```
Open **`http://localhost:5000`** in your browser.

---

## 📡 Hardware & Network Configuration

The default K40 device configuration is defined in [`sync_k40.py`](file:///Users/ranger/development/zkteco/sync_k40.py) and [`dashboard/app.py`](file:///Users/ranger/development/zkteco/dashboard/app.py):

```python
K40_IP = "192.168.10.67"  # IP address of the K40 device
K40_PORT = 4370            # Standard ZKTeco communication port
K40_PASSWORD = 0          # Comm key / password (default 0)
```

To change the device IP, simply update `K40_IP` in `sync_k40.py` and `dashboard/app.py`.

---

## 🧪 Automated Testing & Verification

The project includes an automated test suite verifying all routes, calculation logic, exports, and maintenance tools.

Run the test suite using:
```bash
python dashboard/test_app.py
```

### Verified Test Cases (16/16):
- Default attendance route with KPI summary cards (`Total Staff`, `Present`, `Absent`)
- 12-hour AM/PM time visualization (`strftime('%I:%M:%S %p', ...)`)
- `Total Hours` duration calculation (`Xh Ym`) and single punch `N/A` handling
- User filter dropdown (`?user_id=...`)
- Single date and date range filters (`?from_date=...&to_date=...`)
- Filtered pagination link preservation
- Excel export endpoint with active filters (`.xlsx`)
- Excel export default fallback (latest month)
- Excel export quick presets (`last_week`, `last_month`) with date ranges in filenames
- Excel export chronological ascending date order verification
- Configuration page metrics (`/config`)
- Attendance deletion by month and by year (`/config/delete-attendance`) with database verification

---

## 👩‍💻 Developer Guide (Extending the Project)

### Adding New Routes / Views
- Define new Flask endpoints in [`dashboard/app.py`](file:///Users/ranger/development/zkteco/dashboard/app.py).
- Use `get_db()` helper to query `attendance.db`. SQLite rows are returned as dict-like objects (`db.row_factory = sqlite3.Row`).
- Render templates located in `dashboard/templates/` extending `base.html`.

### Modifying Excel Export Columns or Styles
- Update `export_excel()` in [`dashboard/app.py`](file:///Users/ranger/development/zkteco/dashboard/app.py).
- Openpyxl styling utilities (`PatternFill`, `Font`, `Alignment`, `Border`) are imported and applied at the cell level.

### Hardware Commands via `pyzk`
The `pyzk` library supports direct hardware control through `zk.connect()`:
- `conn.get_users()` / `conn.set_user(...)` — Manage users on machine
- `conn.get_attendance()` — Read all logs
- `conn.clear_attendance()` — Clear device log memory
- `conn.set_time(datetime.now())` — Synchronize hardware clock with PC time
- `conn.test_voice(index=0)` — Play voice prompt ("Thank You")
- `conn.live_capture()` — Event iterator for real-time punch streaming

---

## 📄 License
Internal Attendance Management System. Built for ZKTeco K40 devices.
