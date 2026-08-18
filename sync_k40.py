import argparse
from datetime import datetime
from pathlib import Path
import sqlite3
import sys
import time
from zk import ZK


# =========================
# K40 CONFIGURATION
# =========================

K40_IP = "192.168.10.67"
K40_PORT = 4370
K40_PASSWORD = 0

# SQLite database file
DB_FILE = Path(__file__).resolve().parent / "attendance.db"


# =========================
# DATABASE
# =========================

def create_database():
    db = sqlite3.connect(DB_FILE)

    db.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id TEXT PRIMARY KEY,
            name TEXT,
            privilege TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    db.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            user_id TEXT NOT NULL,
            timestamp TEXT NOT NULL,

            status INTEGER,
            punch INTEGER,

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            UNIQUE (
                user_id,
                timestamp,
                status,
                punch
            ),

            FOREIGN KEY (user_id)
                REFERENCES users(user_id)
        )
    """)

    db.commit()
    return db


# =========================
# SYNC USERS
# =========================

def sync_users(device, db):
    print("Syncing users from device...")

    try:
        users = device.get_users()
        print(f"Found {len(users)} users on device.")

        inserted = 0
        updated = 0

        for user in users:
            user_id = str(user.user_id)
            name = user.name or ""
            privilege = str(user.privilege)

            existing = db.execute(
                "SELECT user_id FROM users WHERE user_id = ?",
                (user_id,)
            ).fetchone()

            if existing:
                db.execute(
                    """
                    UPDATE users
                    SET name = ?, privilege = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE user_id = ?
                    """,
                    (name, privilege, user_id)
                )
                updated += 1
            else:
                db.execute(
                    """
                    INSERT INTO users (user_id, name, privilege)
                    VALUES (?, ?, ?)
                    """,
                    (user_id, name, privilege)
                )
                inserted += 1

        db.commit()
        print(f"Users sync complete: {inserted} inserted, {updated} updated.")
    except Exception as e:
        print(f"[WARNING] Failed to sync users: {e}")


# =========================
# SYNC ATTENDANCE HISTORY
# =========================

def sync_attendance(device, db):
    print("\nSyncing attendance log history...")

    try:
        records = device.get_attendance()
        print(f"Found {len(records)} attendance records on device.")

        inserted = 0
        skipped = 0

        for record in records:
            user_id = str(record.user_id)
            timestamp = record.timestamp.isoformat()
            status = record.status
            punch = record.punch

            cursor = db.execute(
                """
                INSERT OR IGNORE INTO attendance (user_id, timestamp, status, punch)
                VALUES (?, ?, ?, ?)
                """,
                (user_id, timestamp, status, punch)
            )

            if cursor.rowcount == 1:
                inserted += 1
            else:
                skipped += 1

        db.commit()
        print(f"Attendance history sync complete: {inserted} new inserted, {skipped} existing skipped.")
    except Exception as e:
        print(f"[WARNING] Failed to sync attendance history: {e}")


# =========================
# LIVE CAPTURE LISTENER
# =========================

def start_live_capture(zk, db):
    """
    Listens for real-time punch events sent by the device.
    Automatically reconnects if the connection drops.
    """
    print("\n==============================")
    print("LIVE CAPTURE MODE ACTIVE")
    print("==============================")
    print("Listening for real-time punch events... (Press Ctrl+C to stop)\n")

    while True:
        device = None
        try:
            device = zk.connect()
            device.disable_device()  # Lock device briefly while reading live events
            device.enable_device()

            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Live capture connected to K40 device.")

            for attendance in device.live_capture():
                if attendance is None:
                    continue

                user_id = str(attendance.user_id)
                timestamp = attendance.timestamp.isoformat()
                status = attendance.status
                punch = attendance.punch

                cursor = db.execute(
                    """
                    INSERT OR IGNORE INTO attendance (user_id, timestamp, status, punch)
                    VALUES (?, ?, ?, ?)
                    """,
                    (user_id, timestamp, status, punch)
                )
                db.commit()

                time_str = datetime.now().strftime("%I:%M:%S %p")
                if cursor.rowcount == 1:
                    print(f"⚡ [{time_str}] LIVE PUNCH RECORDED -> User ID: {user_id} | Time: {timestamp}")
                else:
                    print(f"ℹ️ [{time_str}] Live Punch (Already Recorded) -> User ID: {user_id} | Time: {timestamp}")

        except KeyboardInterrupt:
            print("\n[LIVE CAPTURE STOPPED] User exited.")
            if device:
                try:
                    device.disconnect()
                except Exception:
                    pass
            break
        except Exception as e:
            print(f"[CONNECTION LOST] Live capture error: {e}. Reconnecting in 5 seconds...")
            if device:
                try:
                    device.disconnect()
                except Exception:
                    pass
            time.sleep(5)


# =========================
# MAIN
# =========================

def main():
    parser = argparse.ArgumentParser(description="ZKTeco K40 Attendance Sync & Live Capture")
    parser.add_argument("--once", action="store_true", help="Perform one-time historical sync and exit (no live capture)")
    args = parser.parse_args()

    print("Opening SQLite database...")
    db = create_database()

    zk = ZK(
        K40_IP,
        port=K40_PORT,
        timeout=10,
        password=K40_PASSWORD,
        force_udp=False,
    )

    device = None
    try:
        print(f"Connecting to K40 at {K40_IP}:{K40_PORT}...")
        device = zk.connect()
        print("Connected successfully!")
        print(f"Device: {device.get_device_name()} | Serial: {device.get_serialnumber()} | Firmware: {device.get_firmware_version()}")

        print("\n==============================")
        print("1. HISTORICAL SYNC")
        print("==============================")

        sync_users(device, db)
        sync_attendance(device, db)

    except Exception as e:
        print(f"[ERROR] Could not connect to K40 device: {e}")
    finally:
        if device:
            try:
                device.disconnect()
                print("Disconnected from device after historical sync.")
            except Exception:
                pass

    if args.once:
        print("\nOne-time sync complete. Exiting.")
        db.close()
        sys.exit(0)

    # Start continuous Live Capture
    start_live_capture(zk, db)
    db.close()


if __name__ == "__main__":
    main()
