import io
import sqlite3
import sys
from pathlib import Path
import unittest

import openpyxl

# Ensure dashboard directory is in path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from app import app, DB_FILE


class DashboardTestCase(unittest.TestCase):
    def setUp(self):
        app.config["TESTING"] = True
        self.client = app.test_client()

    def test_default_attendance_page(self):
        """Test default attendance route (daily In & Out summary view for all users)."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Attendance", html)
        self.assertIn("Filter User:", html)
        self.assertIn("From:", html)
        self.assertIn("To (optional):", html)
        self.assertIn("-- All Users --", html)
        self.assertIn("In Time (First)", html)
        self.assertIn("Out Time (Last)", html)
        self.assertIn("Total Hours", html)
        self.assertNotIn("<th>Status</th>", html)

    def test_navigation_header_links(self):
        """Test that header includes Attendance, Users, and Configuration links."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('href="/"', html)
        self.assertIn('href="/users"', html)
        self.assertIn('href="/config"', html)

    def test_config_page_loads(self):
        """Test that Configuration page loads and displays stats and maintenance forms."""
        response = self.client.get("/config")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Configuration", html)
        self.assertIn("Maintenance", html)
        self.assertIn("Local Database", html)
        self.assertIn("ZKTeco K40 Device", html)
        self.assertIn("Delete Attendance Records", html)
        self.assertIn("Delete by Month", html)
        self.assertIn("Delete by Year", html)

    def test_delete_attendance_by_month(self):
        """Test deleting attendance records by month with confirmation action."""
        # Insert a temporary test record for 1999-01
        db = sqlite3.connect(DB_FILE)
        db.execute(
            "INSERT OR IGNORE INTO attendance (user_id, timestamp, status, punch) VALUES ('1', '1999-01-15T09:00:00', 1, 0)"
        )
        db.commit()
        db.close()

        response = self.client.post(
            "/config/delete-attendance",
            data={"delete_type": "month", "target_value": "1999-01"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Successfully deleted", html)
        self.assertIn("1999-01", html)

        # Verify record is gone
        db = sqlite3.connect(DB_FILE)
        count = db.execute("SELECT COUNT(*) FROM attendance WHERE strftime('%Y-%m', timestamp) = '1999-01'").fetchone()[0]
        db.close()
        self.assertEqual(count, 0)

    def test_delete_attendance_by_year(self):
        """Test deleting attendance records by year."""
        # Insert a temporary test record for 1998
        db = sqlite3.connect(DB_FILE)
        db.execute(
            "INSERT OR IGNORE INTO attendance (user_id, timestamp, status, punch) VALUES ('1', '1998-05-20T10:00:00', 1, 0)"
        )
        db.commit()
        db.close()

        response = self.client.post(
            "/config/delete-attendance",
            data={"delete_type": "year", "target_value": "1998"},
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Successfully deleted", html)
        self.assertIn("1998", html)

        # Verify record is gone
        db = sqlite3.connect(DB_FILE)
        count = db.execute("SELECT COUNT(*) FROM attendance WHERE strftime('%Y', timestamp) = '1998'").fetchone()[0]
        db.close()
        self.assertEqual(count, 0)

    def test_kpi_analytics_cards(self):
        """Test that retained KPI cards (Total Staff, Present, Absent) are rendered."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Total Staff", html)
        self.assertIn("Present", html)
        self.assertIn("Absent", html)
        self.assertNotIn("Late Arrivals", html)
        self.assertNotIn("Avg. Duration", html)

    def test_total_hours_calculation(self):
        """Test total hours calculation and N/A fallback for single punch."""
        response = self.client.get("/?user_id=1&from_date=2026-08-11&to_date=2026-08-11")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("<th>Total Hours</th>", html)
        # On 2026-08-11 User 1 has punches 10:36:34 to 14:43:51 (4h 07m)
        self.assertIn("4h 07m", html)

    def test_user_id_column_removed_from_ui(self):
        """Test that 'User ID' table header column is removed from UI view."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertNotIn("<th>User ID</th>", html)

    def test_12_hour_am_pm_time_format(self):
        """Test that timestamps are formatted in 12-hour AM/PM format."""
        response = self.client.get("/?user_id=1&from_date=2026-08-11&to_date=2026-08-11")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("AM", html)
        self.assertIn("PM", html)

    def test_user_filter(self):
        """Test filtering by user_id='1'."""
        response = self.client.get("/?user_id=1")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Showing daily attendance for User 1", html)
        self.assertIn("In Time (First)", html)
        self.assertIn("Out Time (Last)", html)
        self.assertIn("Total Hours", html)
        self.assertIn("Clear Filter", html)

    def test_from_date_filter_only(self):
        """Test filtering by from_date='2026-08-01' (optional to_date omitted)."""
        response = self.client.get("/?from_date=2026-08-01")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("from 2026-08-01 onwards", html)
        self.assertIn("In Time (First)", html)

    def test_export_excel_active_filters(self):
        """Test Excel export endpoint with active filters."""
        response = self.client.get("/export?user_id=1&from_date=2026-08-01&to_date=2026-08-10")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.assertTrue(len(response.data) > 1000)

    def test_export_excel_total_hours_header_and_data(self):
        """Test that exported Excel spreadsheet includes 'Total Hours' and no 'Status' column."""
        response = self.client.get("/export?user_id=1&from_date=2026-08-11&to_date=2026-08-11")
        self.assertEqual(response.status_code, 200)
        wb = openpyxl.load_workbook(io.BytesIO(response.data))
        ws = wb.active
        headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        self.assertIn("Total Hours", headers)
        self.assertNotIn("Status", headers)
        total_hours_index = headers.index("Total Hours")
        data_rows = [row for row in ws.iter_rows(min_row=2, values_only=True)]
        self.assertTrue(len(data_rows) > 0)
        self.assertEqual(data_rows[0][total_hours_index], "4h 07m")

    def test_export_excel_ascending_date_order(self):
        """Test that exported Excel records are sorted in ascending date order."""
        response = self.client.get("/export?user_id=1&from_date=2026-08-01&to_date=2026-08-10")
        self.assertEqual(response.status_code, 200)
        wb = openpyxl.load_workbook(io.BytesIO(response.data))
        ws = wb.active
        dates = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]
        self.assertTrue(len(dates) > 1)
        self.assertEqual(dates, sorted(dates))

    def test_export_excel_default_latest_month(self):
        """Test Excel export endpoint with no filters (defaults to latest month)."""
        response = self.client.get("/export")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.assertTrue(len(response.data) > 1000)

    def test_export_excel_presets(self):
        """Test Excel export endpoint with preset parameter and date range filename."""
        for preset in ["last_week", "last_month"]:
            response = self.client.get(f"/export?preset={preset}")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.mimetype, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            disposition = response.headers.get("Content-Disposition", "")
            self.assertIn(f"filename=attendance_{preset}_", disposition)
            self.assertIn("_to_", disposition)
            self.assertTrue(len(response.data) > 1000)


if __name__ == "__main__":
    unittest.main()
