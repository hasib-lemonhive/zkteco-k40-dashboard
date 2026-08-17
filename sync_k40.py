import sqlite3
from zk import ZK


# =========================
# K40 CONFIGURATION
# =========================

K40_IP = "192.168.10.67"
K40_PORT = 4370
K40_PASSWORD = 0

# SQLite database file
DB_FILE = "attendance.db"


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
    print("Getting users...")

    users = device.get_users()

    print(f"Found {len(users)} users")

    inserted = 0
    updated = 0

    for user in users:
        user_id = str(user.user_id)
        name = user.name or ""
        privilege = str(user.privilege)

        existing = db.execute(
            """
            SELECT user_id
            FROM users
            WHERE user_id = ?
            """,
            (user_id,)
        ).fetchone()

        if existing:
            db.execute(
                """
                UPDATE users
                SET
                    name = ?,
                    privilege = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE user_id = ?
                """,
                (
                    name,
                    privilege,
                    user_id,
                )
            )

            updated += 1

        else:
            db.execute(
                """
                INSERT INTO users (
                    user_id,
                    name,
                    privilege
                )
                VALUES (?, ?, ?)
                """,
                (
                    user_id,
                    name,
                    privilege,
                )
            )

            inserted += 1

    db.commit()

    print(f"Users inserted: {inserted}")
    print(f"Users updated:  {updated}")


# =========================
# SYNC ATTENDANCE
# =========================

def sync_attendance(device, db):
    print("\nGetting attendance records...")

    records = device.get_attendance()

    print(f"Found {len(records)} attendance records")

    inserted = 0
    skipped = 0

    for record in records:

        user_id = str(record.user_id)

        timestamp = record.timestamp.isoformat()

        status = record.status
        punch = record.punch

        cursor = db.execute(
            """
            INSERT OR IGNORE INTO attendance (
                user_id,
                timestamp,
                status,
                punch
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                user_id,
                timestamp,
                status,
                punch,
            )
        )

        if cursor.rowcount == 1:
            inserted += 1
        else:
            skipped += 1

    db.commit()

    print(f"Attendance inserted: {inserted}")
    print(f"Attendance skipped:  {skipped}")


# =========================
# MAIN
# =========================

def main():

    print("Creating/opening SQLite database...")

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

        print("Connected!")
        print(f"Device: {device.get_device_name()}")
        print(f"Serial: {device.get_serialnumber()}")
        print(f"Firmware: {device.get_firmware_version()}")

        print("\n==============================")
        print("SYNC USERS")
        print("==============================")

        sync_users(device, db)

        print("\n==============================")
        print("SYNC ATTENDANCE")
        print("==============================")

        sync_attendance(device, db)

        print("\n==============================")
        print("SYNC COMPLETE")
        print("==============================")

    except Exception as e:

        print(f"\nERROR: {e}")

    finally:

        if device:
            device.disconnect()
            print("Disconnected from K40.")

        db.close()


if __name__ == "__main__":
    main()
