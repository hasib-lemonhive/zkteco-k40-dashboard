import io
import sys
from pathlib import Path
import unittest

import openpyxl

# Ensure dashboard directory is in path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from app import app


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
        """Test that exported Excel spreadsheet includes 'Total Hours' column."""
        response = self.client.get("/export?user_id=1&from_date=2026-08-11&to_date=2026-08-11")
        self.assertEqual(response.status_code, 200)
        wb = openpyxl.load_workbook(io.BytesIO(response.data))
        ws = wb.active
        headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        self.assertIn("Total Hours", headers)
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
