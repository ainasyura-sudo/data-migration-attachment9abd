\# Attachment 9 Data Migration Pipeline



Python-based data migration pipeline for extracting, transforming, validating, and loading Attachment 9 data from Excel into PostgreSQL staging tables.



\## Overview



The project processes three main data areas from Attachment 9:



| Data Area | Source Sheets | Target Table |

|---|---|---|

| CESS Contribution | Attachment 9a + 9f | `amir.abd\_cess\_combined\_stg` |

| Master Abandonment Plan | Attachment 9a(ii) + 9b | `amir.abd\_masterplan\_combined\_stg` |

| Decommissioning Cost | Attachment 9g + 9h | `amir.abd\_decom\_combined\_stg` |



Supporting decommissioning staging tables are also populated:



\- `amir.abd\_decom\_summary\_stg`

\- `amir.abd\_decom\_facility\_stg`



\## Project Structure



```text

data-migration/

├── src/

│   ├── \_\_init\_\_.py

│   ├── cess\_contribution.py

│   ├── migrate\_next.py

│   ├── field\_mapping.py

│   └── data\_quality\_report.py

│

├── environment.example.yaml

├── pac\_name\_mapping.example.yaml

├── .gitignore

└── README.md

```



Local source workbooks, generated outputs, database credentials, and actual metadata mapping files are excluded from version control.



\## Migration Flow



\### CESS Contribution



`cess\_contribution.py` processes Attachment 9a and Attachment 9f.



```text

Attachment 9a + 9f

&#x20;       ↓

Extract and clean

&#x20;       ↓

Combine by field\_name + year

&#x20;       ↓

Generate cess\_combined.csv

&#x20;       ↓

Review CSV

&#x20;       ↓

Validate field\_id

&#x20;       ↓

Upsert PostgreSQL

```



The current validated dataset produces 77 CESS field/year records.



\### Master Abandonment Plan



`migrate\_next.py` combines annual abandonment planning data from Attachment 9a(ii) with activity information from Attachment 9b.



The current validated dataset produces 35 field/year records.



\### Decommissioning



`migrate\_next.py` processes:



\- Attachment 9g — field-level decommissioning cost summary

\- Attachment 9h — facility-level decommissioning cost breakdown



The information is loaded into supporting staging tables and the combined decommissioning staging table.



The current validated combined dataset produces 37 records.



\## Asset Hierarchy Mapping



`field\_id` represents an existing asset hierarchy record in `amir.am\_field`.



Migration rules:



\- Use only an existing registered `field\_id`.

\- Do not generate replacement hierarchy IDs.

\- Fields not registered in `am\_field` retain `field\_id = NULL`.

\- Staging table `id` is separate from `field\_id`.

\- PostgreSQL generates the staging UUID for new rows.



This prevents the migration process from creating artificial asset hierarchy records.



\## Configuration



Create local configuration files from the supplied examples:



```powershell

Copy-Item environment.example.yaml environment.yaml

Copy-Item pac\_name\_mapping.example.yaml pac\_name\_mapping.yaml

```



Update `environment.yaml` with the local database and workbook configuration.



Do not commit passwords directly to configuration files.



A database password can be provided through the environment variable configured by `password\_env`, for example:



```powershell

$env:DIC\_DB\_PASSWORD = "YOUR\_PASSWORD"

```



If the environment variable is not set, the migration scripts can prompt for the PostgreSQL password.



Update `pac\_name\_mapping.yaml` with verified metadata for the source workbook before building the CESS CSV.



\## Running CESS Migration



Run commands from the repository root.



\### 1. Build CESS CSV



```powershell

python -m src.cess\_contribution

```



This reads the Excel workbook and creates:



```text

output/cess\_combined.csv

```



Review the generated CSV before database loading.



\### 2. Load Reviewed CESS CSV



```powershell

python -m src.cess\_contribution --load-db

```



The loader validates registered hierarchy mappings before performing a rerunnable upsert into:



```text

amir.abd\_cess\_combined\_stg

```



The CESS business key is:



```text

(field\_name, year)

```



\## Running Masterplan and DECOM Migration



Run:



```powershell

python .\\src\\migrate\_next.py

```



The script processes the Masterplan and Decommissioning datasets and loads their corresponding PostgreSQL staging tables.



\## Current Validation Results



The current test workbook produced:



| Dataset | Rows |

|---|---:|

| CESS Combined | 77 |

| Masterplan Combined | 35 |

| DECOM Summary | 7 |

| DECOM Facility | 35 |

| DECOM Combined | 37 |



\## Data Cleaning Rules



The migration applies several validation and cleaning rules, including:



\- Blank values and `-` converted to `NULL` where appropriate

\- Valid year checking

\- Field names trimmed before processing

\- Duplicate Attachment 9a(i) field records resolved using the latest applicable row

\- Attachment 9b activities filtered using non-zero FY values

\- Existing asset hierarchy IDs validated before database loading



\## Security



The repository excludes:



\- `environment.yaml`

\- `pac\_name\_mapping.yaml`

\- Excel source workbooks

\- generated output files

\- Python virtual environments

\- Python cache files



Use the provided example configuration files when setting up the project in a new environment.

