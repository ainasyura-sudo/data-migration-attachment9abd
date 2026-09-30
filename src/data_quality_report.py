"""Read-only CSV quality report; runs before any database load."""

import csv
from collections import Counter
from pathlib import Path

import yaml


BASE = Path(__file__).resolve().parent
config = yaml.safe_load((BASE / "environment.yaml").read_text(encoding="utf-8"))
folder = BASE / config["output"]["directory"]


def read(name):
    with (folder / name).open(newline="", encoding="utf-8-sig") as stream:
        return list(csv.DictReader(stream))


master = read("masterplan.csv")
summary = read("decom_summary.csv")
facility = read("decom_facility.csv")
combined = read("decom_combined.csv")
report = []
details = []


def check(name, count, meaning, severity="review"):
    report.append({"check": name, "count": count, "severity": severity,
                   "meaning": meaning})
    print(f"{name}: {count}")


def duplicates(name, rows, key_columns):
    counts = Counter(tuple(row[col] for col in key_columns) for row in rows)
    found = {key: count for key, count in counts.items() if count > 1}
    check(name, len(found), "Number of duplicated natural keys", "blocker")
    for key, count in found.items():
        details.append({"issue": name, "record": repr(key), "detail": str(count)})


duplicates("masterplan_duplicate_field_year", master, ("field", "year"))
duplicates("decom_summary_duplicate_field", summary, ("field",))
duplicates("decom_facility_duplicate_field_source_row", facility,
           ("field", "source_row_9h"))

check("masterplan_blank_field", sum(not row["field"].strip() for row in master),
      "A field name is required", "blocker")
check("masterplan_missing_work_amount",
      sum(row["work_amount_9a_rm_million"] == "" for row in master),
      "Review 9a(ii) blanks; do not replace with zero")
check("masterplan_nonzero_9b_activity",
      sum(row["activity_fy_rm_million"] not in ("", "0", "0.0") for row in master),
      "Only nonzero FY activities are extracted")
check("masterplan_unexpected_year",
      sum(row["year"] not in {str(year) for year in range(2026, 2031)}
          for row in master), "Year is outside 2026–2030", "blocker")

facility_fields = {row["field"] for row in facility}
summary_fields = {row["field"] for row in summary}
missing = sorted(summary_fields - facility_fields)
check("9g_fields_without_9h_facilities", len(missing),
      "One blank-facility combined row expected for each such field")
for field in missing:
    details.append({"issue": "9g_without_9h", "record": field,
                    "detail": "No 9h section"})
unknown = sorted(facility_fields - summary_fields)
check("9h_fields_without_9g_summary", len(unknown),
      "9h sections cannot join to 9g", "blocker")
for field in unknown:
    details.append({"issue": "9h_without_9g", "record": field,
                    "detail": "No 9g summary"})
expected = sum(1 for row in facility if row["field"] in summary_fields) + len(missing)
check("combined_row_count_difference", abs(len(combined) - expected),
      f"Expected {expected} joined rows, found {len(combined)}", "blocker")
check("combined_blank_facility", sum(not row["facility"] for row in combined),
      "Compare with 9g fields without 9h sections")

meta_columns = ("psc_name", "psc_contractor", "psc_focal_point",
                "region_psc", "psc_expiry", "awpb_year")
for name, rows in (("masterplan", master), ("decom_summary", summary),
                   ("decom_facility", facility), ("decom_combined", combined)):
    missing_meta = sum(any(not row.get(key, "").strip() for key in meta_columns)
                       for row in rows)
    check(f"{name}_incomplete_psc_metadata", missing_meta,
          "Confirm cover-page metadata", "blocker")


def write(name, columns, rows):
    with (folder / name).open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


write("data_quality_report.csv", ("check", "count", "severity", "meaning"), report)
write("data_quality_details.csv", ("issue", "record", "detail"), details)
print(f"Reports created in {folder}")
