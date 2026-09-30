"""Load the reviewed Attachment 9 CSVs into PostgreSQL staging tables.

Run build_csv.py and data_quality_report.py first, then run this script.
No Excel workbook is read here. All four CSVs are loaded in one transaction.
"""

import csv
import getpass
import os
from collections import Counter, defaultdict
from pathlib import Path

import psycopg
from psycopg import sql
import yaml

BASE = Path(__file__).resolve().parent
cfg = yaml.safe_load((BASE / "environment.yaml").read_text(encoding="utf-8"))
db = cfg["database"]
folder = BASE / cfg.get("output", {}).get("directory", "output")
names = cfg.get("tables", {})


def read_csv(name):
    with (folder / name).open(newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


plan = read_csv("masterplan.csv")
summary = read_csv("decom_summary.csv")
facility = read_csv("decom_facility.csv")
combined = read_csv("decom_combined.csv")


def unique(rows, keys, name):
    ids = [tuple(row[key] for key in keys) for row in rows]
    if len(ids) != len(set(ids)):
        raise ValueError(f"Duplicate key in {name}: {keys}")


unique(plan, ("field", "year"), "masterplan.csv")
unique(summary, ("field",), "decom_summary.csv")
unique(facility, ("field", "source_row_9h"), "decom_facility.csv")
if len(combined) != len(facility) + len(set(r["field"] for r in summary) - set(r["field"] for r in facility)):
    raise ValueError("decom_combined.csv row count does not match summary/facility")

# CSV combined deliberately omits source_row_9h. Match repeated facility labels
# to facility CSV in original order, so the two Terminal sections stay distinct.
fac_lookup = defaultdict(list)
for row in facility:
    fac_lookup[(row["field"], row["facility"])].append(row["source_row_9h"])
fac_position = Counter()
for row in combined:
    key = row["field"], row["facility"]
    if row["facility"]:
        position = fac_position[key]
        if position >= len(fac_lookup[key]):
            raise ValueError(f"Cannot match combined facility to 9h: {key}")
        row["source_row_9h"] = fac_lookup[key][position]
        fac_position[key] += 1
    else:
        row["source_row_9h"] = "0"
unique(combined, ("field", "source_row_9h"), "decom_combined.csv")


def convert(column, raw):
    if raw in (None, ""):
        return None
    if column in {"year", "source_row_9h", "awpb_year", "psc_expiry", "source_row_9b"}:
        return int(raw)
    if column in {"field", "field_id", "facility", "activity_9b", "asset_9b", "remarks_9b",
                  "psc_name", "psc_contractor", "psc_focal_point", "region_psc"}:
        return raw
    return float(raw)


meta = ("psc_name", "psc_contractor", "psc_focal_point", "region_psc", "psc_expiry", "awpb_year")
plan_cols = ("field", "field_id", "year", "work_amount_9a_rm_million", "source_row_9b", "asset_9b",
             "activity_9b", "q1_rm_million", "q2_rm_million", "q3_rm_million", "q4_rm_million",
             "activity_fy_rm_million", "remarks_9b") + meta
summary_cols = ("field", "platform_rm_million", "well_rm_million", "pipeline_rm_million",
                "floater_rm_million", "subsea_facility_rm_million", "terminal_rm_million",
                "total_rt_rm_million", "total_mod_rm_million") + meta
facility_cols = ("field", "source_row_9h", "facility", "total_cost_rm_million") + meta
combined_cols = ("field", "field_id", "platform_rm_million", "well_rm_million", "pipeline_rm_million",
                 "floater_rm_million", "subsea_facility_rm_million", "terminal_rm_million",
                 "total_rt_rm_million", "total_mod_rm_million", "source_row_9h", "facility",
                 "facility_cost_9h_rm_million") + meta

password = db.get("password") or os.environ.get(db.get("password_env", "DIC_DB_PASSWORD"))
if not password:
    password = getpass.getpass("PostgreSQL password: ")


def table(key, fallback):
    return sql.Identifier(db.get("schema", "public"), names.get(key, fallback))


def upsert(cur, target, rows, columns, keys):
    if not rows:
        raise ValueError(f"Refusing empty CSV load: {target.as_string(cur)}")
    for row in rows:
        missing = set(columns) - row.keys()
        if missing:
            raise ValueError(f"CSV missing columns: {sorted(missing)}")
    query = sql.SQL("INSERT INTO {} ({}) VALUES ({}) ON CONFLICT ({}) DO UPDATE SET {}").format(
        target,
        sql.SQL(", ").join(map(sql.Identifier, columns)),
        sql.SQL(", ").join(sql.Placeholder() for _ in columns),
        sql.SQL(", ").join(map(sql.Identifier, keys)),
        sql.SQL(", ").join(
            sql.SQL("{} = COALESCE(EXCLUDED.{}, {}.{})").format(
                sql.Identifier(col), sql.Identifier(col), target, sql.Identifier(col))
            if col == "field_id" else
            sql.SQL("{} = EXCLUDED.{}").format(sql.Identifier(col), sql.Identifier(col))
            for col in columns if col not in keys
        ),
    )
    cur.executemany(query, [tuple(convert(col, row[col]) for col in columns) for row in rows])


def reconcile_combined_keys(cur, target):
    """Assign source keys to legacy rows without changing IDs or created_at."""
    fields = sorted({row["field"] for row in combined})
    cur.execute(sql.SQL(
        "SELECT id, field, facility, facility_cost_9h_rm_million, source_row_9h "
        "FROM {} WHERE field = ANY(%s) ORDER BY id"
    ).format(target), (fields,))
    existing = cur.fetchall()
    old_groups = defaultdict(list)
    new_groups = defaultdict(list)
    for row in existing:
        old_groups[(row[1], row[2] or "")].append(row)
    for row in combined:
        new_groups[(row["field"], row["facility"])].append(row)
    # Older joins sometimes generated two empty rows for a summary-only field.
    # Remove the excess only when both rows have no facility or facility cost.
    for group, old_rows in list(old_groups.items()):
        expected = len(new_groups[group])
        if (group[1] == "" and expected == 1 and len(old_rows) == 2
                and all(r[2] is None and r[3] is None for r in old_rows)
                and all(r[4] in (None, 0) for r in old_rows)):
            retained = next((r for r in old_rows if r[4] == 0), old_rows[0])
            excess = next(r for r in old_rows if r[0] != retained[0])
            cur.execute(sql.SQL(
                "DELETE FROM {} WHERE id = %s AND field = %s "
                "AND facility IS NULL AND facility_cost_9h_rm_million IS NULL "
                "RETURNING id"
            ).format(target), (excess[0], group[0]))
            if cur.fetchone() is None:
                raise ValueError(f"Could not remove verified extra row for {group}")
            old_groups[group] = [retained]
            print(f"Removed one extra legacy empty-facility row for {group[0]}")
    old_counts = Counter({k: len(v) for k, v in old_groups.items()})
    new_counts = Counter({k: len(v) for k, v in new_groups.items()})
    if old_counts != new_counts:
        print("Existing combined table versus CSV (field, facility: DB -> CSV):")
        for group in sorted(old_counts.keys() | new_counts.keys()):
            if old_counts[group] != new_counts[group]:
                print(f"  {group}: {old_counts[group]} -> {new_counts[group]}")
        raise ValueError(
            "Legacy decom rows do not match CSV field/facility counts; "
            "no rows changed. Inspect the existing table before reload."
        )
    assignments = []
    for group, old_rows in old_groups.items():
        incoming = new_groups[group]
        occupied = {r[4] for r in old_rows if r[4] is not None}
        wanted = {int(r["source_row_9h"]) for r in incoming}
        if not occupied <= wanted:
            raise ValueError(f"Unexpected existing source row in {group}: {occupied - wanted}")
        remaining = [r for r in incoming if int(r["source_row_9h"]) not in occupied]
        missing = [r for r in old_rows if r[4] is None]
        if len(remaining) != len(missing):
            raise ValueError(f"Cannot reconcile legacy rows for {group}")
        # Match equal costs first; repeated equal-cost rows follow existing ID order.
        for old in missing[:]:
            match = next((r for r in remaining if
                          (float(r["facility_cost_9h_rm_million"]) if r["facility_cost_9h_rm_million"] else None)
                          == old[3]), None)
            if match is not None:
                assignments.append((int(match["source_row_9h"]), old[0]))
                remaining.remove(match)
                missing.remove(old)
        for old, row in zip(missing, sorted(remaining, key=lambda r: int(r["source_row_9h"]))):
            assignments.append((int(row["source_row_9h"]), old[0]))
    if assignments:
        cur.executemany(sql.SQL("UPDATE {} SET source_row_9h = %s WHERE id = %s").format(target), assignments)
        print(f"Reconciled {len(assignments)} legacy decom combined keys; IDs retained")


targets = (
    ("masterplan", "abd_masterplan_combined_stg", plan, plan_cols, ("field", "year")),
    ("decom_summary", "abd_decom_summary_stg", summary, summary_cols, ("field",)),
    ("decom_facility", "abd_decom_facility_stg", facility, facility_cols, ("field", "source_row_9h")),
    ("decom_combined", "abd_decom_combined_stg", combined, combined_cols, ("field", "source_row_9h")),
)

with psycopg.connect(host=db["host"], port=int(db.get("port", 5432)), dbname=db["name"],
                     user=db["user"], password=password, connect_timeout=10) as conn:
    with conn.cursor() as cur:
        for key, fallback, rows, columns, keys in targets:
            target = table(key, fallback)
            # Existing staging tables are retained. Add only CSV metadata columns.
            for column, kind in (("psc_name", "TEXT"), ("psc_contractor", "TEXT"),
                                 ("psc_focal_point", "TEXT"), ("region_psc", "TEXT"),
                                 ("psc_expiry", "INTEGER"), ("awpb_year", "INTEGER")):
                cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS {} {}").format(
                    target, sql.Identifier(column), sql.SQL(kind)))
            if key == "decom_combined":
                cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS source_row_9h INTEGER").format(target))
                reconcile_combined_keys(cur, target)
                cur.execute(sql.SQL("CREATE UNIQUE INDEX IF NOT EXISTS {} ON {} (field, source_row_9h)").format(
                    sql.Identifier(names.get(key, fallback) + "_field_source_row_uk"), target))
            upsert(cur, target, rows, columns, keys)
            cur.execute(sql.SQL("SELECT COUNT(*) FROM {}").format(target))
            count = cur.fetchone()[0]
            print(f"{names.get(key, fallback)}: CSV {len(rows)}, DB {count}")
            if count != len(rows):
                print("  Extra older rows remain in this table; no rows were deleted.")

print("CSV load complete. Refresh tables in DBeaver.")
