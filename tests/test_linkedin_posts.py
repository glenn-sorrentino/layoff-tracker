from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from scripts.lib.linkedin_posts import WarnRow, existing_archive_index, select_post, write_post_archive


def make_row(
    *,
    company: str,
    county: str,
    effective_date: date,
    employees: int,
    address: str,
    industry: str = "Technology",
    layoff_closure: str = "Layoff Permanent",
) -> WarnRow:
    return WarnRow(
        county=county,
        county_normalized=county.replace(" County", ""),
        notice_date=None,
        processed_date=None,
        effective_date=effective_date,
        company=company,
        layoff_closure=layoff_closure,
        employees=employees,
        address=address,
        industry=industry,
    )


class LinkedInPostSelectionTests(unittest.TestCase):
    def test_uses_todays_event_for_single_listing_today(self) -> None:
        rows = [
            make_row(
                company="Riot Games",
                county="Los Angeles County",
                effective_date=date(2026, 4, 27),
                employees=26,
                address="123 Main St",
            ),
            make_row(
                company="Later Co",
                county="Alameda County",
                effective_date=date(2026, 4, 28),
                employees=10,
                address="1 Later St",
            ),
        ]

        candidate = select_post(rows, date(2026, 4, 27), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "todays_event")
        self.assertIn("Company: Riot Games", candidate["linkedin"])
        self.assertIn("Learn more: https://ca-warn.github.io/layoff-tracker/", candidate["linkedin"])

    def test_uses_todays_summary_for_multiple_listings_today(self) -> None:
        rows = [
            make_row(
                company="Riot Games",
                county="Los Angeles County",
                effective_date=date(2026, 4, 27),
                employees=26,
                address="123 Main St",
            ),
            make_row(
                company="Amazon LAX18",
                county="Los Angeles County",
                effective_date=date(2026, 4, 27),
                employees=7,
                address="456 Second St",
            ),
        ]

        candidate = select_post(rows, date(2026, 4, 27), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "todays_summary")
        self.assertIn("Today, there are 33 scheduled layoffs", candidate["linkedin"])
        self.assertIn("Company: Riot Games", candidate["linkedin"])
        self.assertIn("Company: Amazon LAX18", candidate["linkedin"])

    def test_uses_quarterly_total_on_first_day_of_quarter_when_no_today_rows(self) -> None:
        rows = [
            make_row(
                company="Past Co",
                county="Los Angeles County",
                effective_date=date(2026, 3, 31),
                employees=20,
                address="Past St",
            ),
            make_row(
                company="Future Co",
                county="Alameda County",
                effective_date=date(2026, 4, 5),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 4, 1), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "total_layoffs")
        self.assertIn("California WARN file spans", candidate["linkedin"])

    def test_uses_monthly_impact_on_first_of_month_when_not_quarter_start(self) -> None:
        rows = [
            make_row(
                company="April Co",
                county="Los Angeles County",
                effective_date=date(2026, 5, 3),
                employees=20,
                address="A St",
            ),
            make_row(
                company="May Co",
                county="Alameda County",
                effective_date=date(2026, 5, 4),
                employees=10,
                address="B St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 1), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "this_month_impact")
        self.assertIn("In May, there are 2 layoff notices scheduled", candidate["linkedin"])

    def test_uses_no_layoffs_when_no_today_rows_and_no_scheduled_monthly_rollup(self) -> None:
        rows = [
            make_row(
                company="Future Co",
                county="Alameda County",
                effective_date=date(2026, 5, 4),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 5), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "no_layoffs")
        self.assertIn("There are no scheduled layoffs in California today", candidate["linkedin"])

    def test_uses_no_layoffs_this_week_on_monday_when_week_has_no_notices(self) -> None:
        rows = [
            make_row(
                company="Later This Month Co",
                county="Alameda County",
                effective_date=date(2026, 5, 14),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 4), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "no_layoffs_this_week")
        self.assertIn("There are no scheduled layoffs in California this week", candidate["linkedin"])

    def test_skips_weekly_no_layoffs_on_non_monday_when_week_has_no_notices(self) -> None:
        rows = [
            make_row(
                company="Later This Month Co",
                county="Alameda County",
                effective_date=date(2026, 5, 14),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 6), set())
        self.assertIsNone(candidate)

    def test_uses_monthly_zero_post_on_first_when_month_has_no_notices(self) -> None:
        rows = [
            make_row(
                company="Future Co",
                county="Alameda County",
                effective_date=date(2026, 6, 3),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 1), set())

        self.assertIsNotNone(candidate)
        assert candidate is not None
        self.assertEqual(candidate["post_type"], "this_month_impact")
        self.assertIn("In May, there are no layoff notices currently scheduled", candidate["linkedin"])

    def test_skips_daily_no_layoffs_when_month_has_no_notices(self) -> None:
        rows = [
            make_row(
                company="Future Co",
                county="Alameda County",
                effective_date=date(2026, 6, 3),
                employees=10,
                address="Future St",
            ),
        ]

        candidate = select_post(rows, date(2026, 5, 2), set())
        self.assertIsNone(candidate)

    def test_existing_archive_index_reads_fingerprint(self) -> None:
        rows = [
            make_row(
                company="Riot Games",
                county="Los Angeles County",
                effective_date=date(2026, 4, 27),
                employees=26,
                address="123 Main St",
            )
        ]
        candidate = select_post(rows, date(2026, 4, 27), set())
        self.assertIsNotNone(candidate)
        assert candidate is not None

        with tempfile.TemporaryDirectory() as tempdir:
            archive_root = Path(tempdir)
            write_post_archive(
                archive_root=archive_root,
                publish_date=date(2026, 4, 27),
                workbook_path=Path("data/warn_report.xlsx"),
                rows=rows,
                candidate=candidate,
            )

            archive_index = existing_archive_index(archive_root)
            self.assertIn(candidate["fingerprint"], archive_index["fingerprints"])
            self.assertIn("2026-04-27", archive_index["archive_keys"])


if __name__ == "__main__":
    unittest.main()
