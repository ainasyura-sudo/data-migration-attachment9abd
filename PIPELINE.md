# ETL pipeline - Attachment 9 (Balingian PSC AWPB 2026)

## Lineage diagram

```
┌──────────────────────────────────────────┐      ┌──────────────────────────────────┐
│ AWPB 2026 - Balingian PSC -              │      │ pac_name_mapping.yaml            │
│ Attachment 9 2.xlsx                      │      │ PSC cover-page metadata          │
│ sheets 9a, 9b, 9f, 9g, 9h                │      │ (name, contractor, expiry)       │
└─────────────────────┬────────────────────┘      └─────────────────┬────────────────┘
                      │                                             │
                      │ read (openpyxl)                             │
                      │                                             │
                      ▼                                             │
┌──────────────────────────────────────────┐                        │
│ build_csv.py -> masterplan + decom       │        reads metadata  │
│ cess_contribution.py -> CESS             ┤◄───────────────────────┘
│ (no database access in this step)        │
└─────────────────────┬────────────────────┘
                      │
                      │ writes
                      │
                      ▼
┌──────────────────────────────────────────┐
│ output/                                  │
│   masterplan.csv                         │
│   decom_summary.csv                      │    intermediate files:
│   decom_facility.csv                     │    open and review them by hand
│   decom_combined.csv                     │
│   cess_combined.csv                      │
└─────────────────────┬────────────────────┘
                      │
                      │ validate (read-only)
                      ▼
┌──────────────────────────────────────────┐
│ data_quality_report.py                   │
│ checks the 4 masterplan/decom CSVs       │
│ (CESS CSV is reviewed by hand)           │
│ -> data_quality_report.csv               │
│ -> data_quality_details.csv              │
└─────────────────────┬────────────────────┘
                      │
                      │ blockers?
                      ▼
┌──────────────────────────────────────────┐
│ Manual review gate                       │
│ blocker found -> fix, rerun build_csv.py │
│ no blocker    -> continue to load        │
└─────────────────────┬────────────────────┘
                      │
                      │ approved
                      ▼
┌──────────────────────────────────────────┐            ┌────────────────────────────┐
│ field_id_rule.py (shared by all loaders) │ exact match│ amir.am_field              │
│ 1. schema gate: id must allow NULL       ┤◄───────────├ read-only lookup           │
│ 2. exact name match -> am_field.id       │            └────────────────────────────┘
│ 3. post-check: id = field_id             │
└─────────────────────┬────────────────────┘
                      │
                    ┌─┴─────────────────────────────────────────┐
                    ▼                                           ▼
┌────────────────────────────────────────┐  ┌────────────────────────────────────────┐
│ load_csv.py                            │  │ cess_contribution.py --load-db         │
│ masterplan + decom tables              │  │ CESS table                             │
│ reads output/*.csv, upserts            │  │ reads cess_combined.csv, upserts       │
└───────────────────┬────────────────────┘  └───────────────────┬────────────────────┘
                    │                                           │
                    │ one transaction                           │
                    ▼                                           ▼
┌────────────────────────────────────────┐  ┌────────────────────────────────────────┐
│ amir.abd_masterplan_combined_stg       │  │ amir.abd_cess_combined_stg             │
│ amir.abd_decom_combined_stg            │  │                                        │
│ (+ decom_summary, decom_facility)      │  │ key: field + year                      │
│ key: field+year / field+source_row     │  │                                        │
└───────────────────┬────────────────────┘  └───────────────────┬────────────────────┘
                    │                                           │
                    └─────────────────────┬─────────────────────┘
                                          │
                                          ▼
                        ┌────────────────────────────────────┐
                        │ DBeaver                            │
                        │ compare row counts and samples     │
                        └────────────────────────────────────┘
```

Each stage only hands over a file or a table, so any problem can be traced to one stage. Nothing touches PostgreSQL until the CSVs have been reviewed, and each load runs in a single transaction: if any check fails, nothing is written.

## Field ID rule

Applies to the masterplan, decom combined and CESS tables.

```
          source field name
                  │
                  ▼
┌────────────────────────────────────┐
│ exact, case-sensitive match        │
│ in am_field.field_name ?           │
└─────────┬─────────────────┬────────┘
          │ yes             │ no
          ▼                 ▼
┌────────────────────┐  ┌──────────────────────────────────────┐
│ id = field_id      │  │ id = NULL, field_id = NULL           │
│ = am_field.id      │  │ no UUID generated                    │
└────────────────────┘  │ no row added to am_field             │
                        └──────────────────────────────────────┘
```

`id` is no longer a row key. All rows of a field (years, facility sections) share one ID, so a row is identified by its natural key. `sheet_id`, `submission_date`, `region` and `operator` stay NULL until a verified source exists.

## Stage 1 - extract: `build_csv.py` and `cess_contribution.py`

Reads the workbook with openpyxl (values only, no formulas) and flattens it into CSVs. Blank cells and `-` become empty values; zero stays zero. PSC metadata from `pac_name_mapping.yaml` is copied onto every row. No database connection is opened.

## Stage 2 - validate: `data_quality_report.py`

Read-only checks on the four masterplan/decom CSVs: duplicate natural keys, blank fields, years outside 2026-2030, 9g fields without 9h sections, row-count match, and incomplete PSC metadata. It writes `data_quality_report.csv` and `data_quality_details.csv`. Rows marked `blocker` must be fixed before loading.

## Stage 3 - load: `load_csv.py` and `cess_contribution.py --load-db`

Both call `field_id_rule.py` before writing. It refuses to run if `id` is still `NOT NULL`, has a default, or has a unique index; it resolves each field by exact match against `am_field`; and after loading it checks that every row satisfies `id = field_id`. Rows are upserted on the natural key, so reruns update in place.

| Input | CSV | Database table (schema `amir`) | Grain |
| --- | --- | --- | --- |
| 9a(ii) + nonzero FY activities from 9b | `output/masterplan.csv` | `abd_masterplan_combined_stg` | field + year |
| 9g | `output/decom_summary.csv` | `abd_decom_summary_stg` | field |
| 9h facility sections | `output/decom_facility.csv` | `abd_decom_facility_stg` | field + source row |
| 9g left join 9h | `output/decom_combined.csv` | `abd_decom_combined_stg` | facility section or unmatched 9g field |
| 9a + 9f | `output/cess_combined.csv` | `abd_cess_combined_stg` | field + year |

## Run order in `C:\devops\data migration`

1. Put the workbook and the `.py`/`.yaml`/`.sql` files in this folder. Install `openpyxl` and `pyyaml`; database loading also needs `psycopg`.
2. Verify the PSC metadata in `pac_name_mapping.yaml` against the workbook cover page, especially spelling and the `SK` region code.
3. Run `python build_csv.py` and `python cess_contribution.py`. They read Excel and write CSVs to `output/` without touching PostgreSQL.
4. Run `python data_quality_report.py`. Review `output/data_quality_report.csv` and `output/data_quality_details.csv`. Unexpected duplicate or unmatched activity is a blocker; workbook blanks can be reviewed separately.
5. **One-time schema change:** review `proposed_id_schema_change.sql`, then run Part 2 manually in `MPM_DIC_DEV`. It removes the primary key, `NOT NULL` and default on `id`. The loaders refuse to write until this is done.
6. Set the `DIC_DB_PASSWORD` environment variable (the config holds only its name, never the password). Load with `python load_csv.py` and `python cess_contribution.py --load-db`. `migrate_next.py` is the older path that reads Excel directly; prefer the CSV-first loaders.
7. In DBeaver compare sample records and row counts. Earlier workbook counts: 35 masterplan, 7 9g summaries, 35 9h sections, 37 combined rows.

## Rules and limits

- Blank numeric cells and `-` become empty CSV cells; zero stays zero. The 9a(ii) years become `(field, year)` rows. The 9b extraction retains FY costs that are nonzero.
- Masterplan key: `(field, year)`. 9g key: `field`. 9h key: `(field, source_row_9h)`. A field can have repeated facility names, including two `Terminal` sections. Keep the source-row key in the facility staging CSV even though the combined presentation CSV hides it.
- 9g field totals repeat on each matching facility row. Do **not** sum them across `decom_combined.csv`. An unmatched 9g field has a blank facility; previously Bayan NAG and Ridwan.
- Workbook-level PSC metadata repeats on exported rows. `psc_contractor` is a cover-page code, and is not automatically the `operator` DB value. The PAC name aliases remain empty until exact legal names and codes are verified.
- The small EUR/cess table in 9a(ii), item-level detail within 9h, and zero-FY activities in 9b are outside this current extraction.
- Loads are in-place upserts on the natural key, so reruns update matching rows instead of deleting and re-inserting them. Because `id` is shared across a field's rows, use the natural key, not `id`, to identify a row. Choose a primary key in Part 3 of the SQL file.
