# Blasting Automation

A local Streamlit app that copies data from an Excel worksheet into your template's column layout. It starts with **NEGATIVE** and applies the same separation rules to every selected worksheet: Christmas Amnesty, product, and level, followed by **Email** and **Viber files** for every group. GEO worksheets with an **Area** column first separate **Mindanao** and **R3**, combining all levels within each product. GEO/RECO and AUTOSTAT offer Excel downloads with literal text cells for account numbers. CSVs remain available for system uploads, whether the template is Excel, CSV, or the built-in blank template.

## Run

Requires Python 3.11 or newer. In PowerShell, from this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open **http://localhost:8501**. If the dependencies are already installed, run `python -m streamlit run app.py` directly. Stop the server with **Ctrl+C**.

Use the sidebar to switch between three separate pages:

- **GEO & RECO**: the original workbook-to-template workflow, with the existing Area, product, level, and Christmas Amnesty grouping and configured filenames.
- **NEGATIVE B2**: B2 and B4 account matching, separate CARDS/BEL blasting files for each bucket, and a B2 upload panel for Email, SMS and Viber activity logs.
- **AUTOSTAT**: Standard uploads compile into Email and Viber activity files; B2 uploads use separate Email, SMS and Viber templates.

Workbooks stay on the page you select. A DataGrid-looking sheet does not switch GEO/RECO into B2 matching.

Use **Clear all files** in the sidebar on any page to empty every upload box at once, including GEO/RECO sources and templates, B2 activity uploads, both AUTOSTAT modes, and the SBC database/DataGrid uploads. It also clears prepared downloads, previews and source mappings across all pages. Log settings and output preferences are kept. Files already saved to a folder are kept. You can immediately upload a new batch after clearing.

## Automatic AUTOSTAT on GEO & RECO

Uploading a source workbook on **GEO & RECO** automatically prepares the **AUTOSTAT — Email and Viber** section for the selected worksheet, before you click **Preview and prepare files**. Use **Download AUTOSTAT_Email.xlsx** or **Download AUTOSTAT_Viber.xlsx** for separate files, or **Download both AUTOSTAT files (ZIP)**. The combined **AUTOSTAT.xlsx** remains available with separate Email and Viber sheets. Each channel shows its valid account count. Viber follows **TEMPLATE_VIBER.xlsx**: `Time`, `Account No.`, `Status`, `Remark`, `Remark By`, the supplied `SMS BLAST SENT - WAITING FOR REPLY` status, and a remark ending with the validated mobile followed by `VIBER SENT - BLASTING`. A missing phone mapping or no valid mobile accounts produces an explanation beside the downloads. No exported campaign files need to be uploaded again. The logs read the original source cells directly, use the selected source columns and header row, and map **Financing/Card No. (`cardnumber`) to Account No.**, ahead of `primary_key` when both are present. Account digits and leading zeros remain text, zero decimal suffixes are removed, accounts are deduplicated by source identity, and each channel is sorted by OB from highest to lowest. Blank financing/card numbers are excluded without substituting primary_key; account-only files without a mapped cardnumber column retain the existing account-column fallback. Campaign exports retain primary_key and cardnumber in their own columns. The existing contact checks and scientific-notation rejection apply.

The collapsed **AutoStat log settings** panel lets you change **Remark By** (default `MBURGOS`) or leave Time blank. Otherwise, Time uses the current Manila date and time when the logs are prepared and stays fixed on routine reruns. Changing the upload, selected worksheet, source mapping, header row, author, timestamp option, or CSV encoding automatically rebuilds the logs. Removing the upload clears them. SBC INSURANCE, B2, and B4 files or worksheets are excluded from this standard log; use the **B2 files** panel on **NEGATIVE B2** for B2 activity files. The separate AUTOSTAT page remains available for compiling several files and worksheets into one pair of logs.

## AUTOSTAT

AUTOSTAT checks account lists for valid email addresses and mobile numbers, then compiles every selected file and sheet into one Email and one Viber activity log using the supplied `TEMPLATE_EMAIL.xlsx` and `TEMPLATE_VIBER.xlsx` layouts. Both layouts are built in; the sample account, email address, and old timestamp are not included. **Excel (.xlsx)** is the default output: individual downloads are **`AUTOSTAT_Email.xlsx`** and **`AUTOSTAT_Viber.xlsx`**, the ZIP contains those same two workbooks, and folder saves write the two workbooks. Every account cell is literal text. A combined **`AUTOSTAT.xlsx`** download contains the full logs on **Email** and **Viber** sheets.

1. Open **AUTOSTAT** in the sidebar and upload one or more `.xlsx` or `.csv` source files.
   `SBC INSURANCE` files and filenames containing a separate `B2` or `B4` label are excluded automatically from both channel totals and exports. Case, spaces, underscores and punctuation are supported. The collapsed **Excluded files** panel lists skipped uploads. This applies only to AUTOSTAT.
2. Source columns are detected automatically. All file and sheet controls are hidden inside one collapsed **Edit source columns** panel. Open it only to change selected sheets, header rows, account columns, or contact columns. Several email or phone columns can be selected. Account numbers must use a separate column from contacts.
3. Open **Log settings** to change **Remark By** (default `MBURGOS`) or the timestamp (default current Asia/Manila time). Select **Leave date and time blank** to leave the `Time` column empty in both Email and Viber files.
4. Click **Preview and prepare AUTOSTAT files**. The results show just **Valid email accounts** and **Valid phone accounts**, counting each exported account once per channel.
5. Use **Download Excel (keeps card numbers as text)** for the combined workbook, the two individual Excel downloads, or **Download both Excel files (ZIP)**. **Save to a folder** also saves Excel files by default. If a receiving system requires CSV, explicitly select **Output files → CSV (.csv)** in the sidebar; individual files, the ZIP, and folder saves then use CSV. Changing output format clears older prepared results. CSV has no cell types and can be rounded when opened directly in Excel. **Save to a folder** and **Preview** stay collapsed until opened.

CSV cannot specify text column types. Opening a CSV directly in Excel can remove leading zeros or round identifiers longer than 15 digits. Excel downloads write identifiers as literal text, without adding apostrophes or formulas, and are generated only when clicked. Format the destination account column as Text before pasting. **AUTOSTAT skips bad database rows and continues preparing valid accounts.** Rows are skipped when the selected account is written in scientific notation, or a selected source cell contains an Excel error or a formula without a saved value. The page shows the number of skipped rows. This applies to GEO/RECO's automatic logs and the Standard and B2/B4 activity uploads. Skipped accounts are never expanded, guessed, or replaced with a different identifier from the row. A source with no valid remaining accounts produces empty activity logs; GEO/RECO explains this without offering stale downloads. Numeric Excel account cells with 16 or more digits still trigger a precision warning; their supplied values are preserved. Missing digits are never guessed. Formatting an already rounded number as text cannot restore it. See [Microsoft's explanation of Excel number precision](https://support.microsoft.com/en-us/excel/keeping-leading-zeros-and-large-numbers).

Email validation uses the same local format checks as the existing workflow. For phones, AUTOSTAT removes spaces and common phone formatting, keeps the **rightmost 10 digits**, and accepts the result only when it **starts with 9**. For example, `09171234567`, `+63 (917) 123-4567`, and `9171234567` become `9171234567`. Short numbers, results beginning with another digit, and text containing letters are rejected. Numeric Excel phones and integral CSV numbers with `.0` or scientific notation are supported.

AUTOSTAT Standard uploads automatically select **`cardnumber`** as the source for **`Account No.`** in both logs, taking priority over `primary_key` and other account headers when present. Header matching ignores case, spaces, and punctuation. When `cardnumber` is absent, the existing account-header detection and manual column selection remain available. A blank `cardnumber` value is treated as a missing account; another identifier from the same row is never substituted automatically.

AUTOSTAT removes a decimal suffix made entirely of zeros from whole-number account values: `5156041234567897.0` becomes `5156041234567897`, and `000123.00` becomes `000123`. This uses the original text digits without a floating-point conversion. The cleaned account is used consistently in previews, both channel logs, the ZIP, folder saves, and the combined Excel download, including account deduplication. Alphanumeric accounts and nonzero decimal parts are preserved. Scientific account notation stops preparation because expanding an abbreviated value would not restore missing digits. Excel account cells remain **Text**, since Number format with zero decimal places can still replace digits beyond Excel's 15-digit precision limit with zeros.

AUTOSTAT keeps the first valid email and first valid phone per account. When `indexoneuid` (or `Old I.C.`) is present, its value distinguishes records that share the same card number; different source IDs are retained. Rows with the same account number and source ID are deduplicated independently for Email and Viber across selected rows, sheets, and files. Without a source ID, deduplication uses the account number alone. `Account No.` still comes from the selected account column, with `cardnumber` preferred. Matching trims surrounding spaces and ignores letter case while preserving leading zeros and punctuation. Multiple contacts in one cell can be separated by commas, semicolons, pipes, or line breaks; the first valid contact in the selected column order is kept for each channel. Missing account numbers are excluded. An account with valid email and phone details appears in both logs; an account with neither appears in neither. Source data is not changed.

All selected sources contribute to the same two files. After deduplication, each channel is sorted by **OB, highest to lowest**, across all uploaded files and selected sheets. OB is detected from `OB` or `OB (Masterlist)` headers, ignoring case and punctuation. Numeric amounts and comma-formatted amounts are compared numerically. Equal amounts retain upload, sheet, and row order; missing or invalid amounts go last. Sources without OB retain their relative order at the bottom. This order applies to previews, CSVs, the ZIP, folder saves, and the Excel download. BEL, CARDS, and other accounts are compiled together by channel. Duplicate filenames do not create additional output files. If a channel has no valid contacts, its file contains headers only. The Email output keeps all 11 supplied headers, including the extra reference header at the end; unused fields stay blank. The Viber output keeps its five headers. The Email remark formula is rendered as text with the validated address; the Viber `NUMBER` placeholder is replaced with the normalized phone. Template statuses are copied exactly, and Remark By and Time are filled from the page settings.

Preparing these logs does not send emails or Viber messages. Source files are never modified. Changing source files, sheets, columns, header rows, remark author, timestamp settings, or CSV encoding clears the prepared result so it can be rebuilt from the new inputs.

### B2 uploads in AUTOSTAT

Select **Upload type → B2** and use **B2 source files** to upload one or more Excel or CSV account files, including the outputs from NEGATIVE B2. Click **Preview and prepare B2 AUTOSTAT files** to create **AUTOSTAT_B2_Email**, **AUTOSTAT_B2_SMS**, and **AUTOSTAT_B2_Viber**. Excel downloads, a combined **AUTOSTAT_B2.xlsx**, ZIP downloads, previews and folder saves are available. Choose CSV in the sidebar when required for system upload.

B2 and B4 activity files use this mode to map source **`primary_key`** to Email **`Account No.`** and SMS/Viber **`Account number`**, even when `cardnumber` is also present. GEO and RECO use Financing/Card No. instead. This mapping applies to previews, CSVs, Excel files, ZIPs and folder saves, and account deduplication uses the same identifier. Blank primary keys are excluded without substituting another identifier. Header matching ignores case, spaces and punctuation. When a source has no `primary_key` header, the existing account-column detection and manual selection remain available for account-only files and the supplied activity layouts.

The built-in layouts copy the headers, status codes and remark formats from **BUCKET 2 EMAIL AS 10212026.xlsx**, **FOR AS SMS BLASTING 10212026.xlsx**, and **FOR AS VIBER BLASTING 10212026.xlsx**. Email has `Name`, `Account No.`, `STATUS CODE`, `REMARKS`, `REMARKS BY`, `REMARKS DATE`, `Email`, and `Collector`. SMS and Viber have `Account number`, `Name`, `STATUS CODE`, `REMARKS`, `REMARKS BY`, `REMARKS DATE`, and `Mobile Number`; the unused trailing blank column in the Viber sample is omitted.

Email remarks use `<email> | EMAIL SENT | <collector>`. SMS remarks use `<mobile> |  SMS SENT` (including the template's two spaces), and Viber remarks use `<mobile> | VIBER SENT`. Status codes are respectively `EMAIL BLAST SENT - WAITING FOR REPLY`, `SMS BLAST SENT - WAITING FOR REPLY`, and `VIBER BLAST SENT - WAITING FOR REPLY`. Mobile numbers use the template's local eleven-digit format (`09171234567`); both phone channels use the first valid mobile per account.

**Remarks By** defaults to the template's `PBBALUYOT`. **REMARKS DATE** uses the current Manila timestamp in `YYYY-MM-DD HH:MM:SS` format; log settings also support a fixed timestamp or blank dates. Names come from the current upload. **Terms supplies the collector agent code**, copied into the Email output's **Collector** column and email remarks, ahead of any Collector column in the source. If Terms is absent, Collector, Collectory or Collector Code can supply the code. Header matching ignores case, spaces and punctuation. Optional column overrides are available under **Edit source columns**. Missing values stay blank; **Default collector** supplies a fallback when desired. Historical dates, customer rows, and collectors from the sample workbooks are never inserted into new uploads. This mapping applies to B2 uploads on NEGATIVE B2 and in AUTOSTAT's B2 mode.

B2 accepts the files explicitly uploaded in its own mode, including filenames marked B2. It uses the same contact validation, account text preservation, scientific-notation rejection, identity-based deduplication and descending OB order as Standard. Each valid email produces an Email row; each valid phone produces one SMS row and one Viber row. Standard retains its existing B2/B4/insurance exclusions and output templates. B2 uses its own upload and result state, and changing its inputs clears its previous downloads. Preparing files does not send messages.

## NEGATIVE B2: match DataGrid with the SBC database

The **B2 files** panel on this page accepts multiple B2 Excel or CSV account files, including NEGATIVE B2 campaign outputs. Uploading them automatically prepares the supplied Email, SMS and Viber activity layouts, with primary_key in each account column and the existing B2 mobile validation. Use the combined Excel workbook, individual channel downloads, ZIP, or folder save. B2 defaults to Excel; its CSV option shares the page's sidebar encoding with matching. The activity uploads work independently of the SBC database/DataGrid uploads. Replacing a B2 activity file, including one with the same filename, rebuilds its logs; removing all activity uploads clears those logs. **Clear all files** clears both activity uploads and matching uploads. GEO & RECO contains its own source and template uploads.

Open **NEGATIVE B2** and upload the full SBC database and the DataGrid account list. The page selects the full `SBC` sheet by its required headers. DataGrid defines the account population; database accounts absent from DataGrid are excluded.

B2 matching runs only on its dedicated page. Daily filenames can change because matching uses source columns and file contents. Both B2 files use the exact 16-column layout from **1EMAIL_Blasting - Template.xlsx**: `email`, `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `phone`, `OB`, `Monthly`, `Terms`, `doclass`, `Level`, `Product`, `EndoDate`, `AREA`, `ChristmasAmnesty`, `NewEndo`. Sample rows from the template are never included.

Upload the latest database each day, then click **Match and prepare B2 files**. Replacing either upload clears previous results and reads the new file contents, including when a replacement has the same filename. The app uses the database you upload; it does not synchronize a changing database on your computer automatically.

`NEGATIVE B2 - BP.xlsx` and `NEGATIVE B2 - CARDS AND BEL.xlsx` are DataGrid account lists. Choose the list you want to process and pair it with the `sbc cards-bel database as of ...xlsx` contact database, selecting its full **SBC** sheet rather than **SBC Lite** or **SBC BEL Lite**. B2 recognizes reversed uploads from their columns and assigns the database and account-list roles automatically. If both uploads are account lists, it requests the missing database before matching; an account list is never treated as the contact database.

DataGrid `Account No.` matches database `Account Key`, and `Financing/Card No.` matches database `Contract Number`. Both identifiers must identify the same record when both are supplied. Matching retains leading zeros and uses exact identifiers after trimming spaces and ignoring letter case. Unmatched, conflicting, ambiguous, and missing identifiers appear in the matching report. Matching never uses customer names. Only database Bucket Code `2` / `B2` qualifies; other or missing buckets are excluded. Product Category Code `BEL` creates **`NEGATIVE B2 - BEL.csv`**; supported card categories such as `MC` create **`NEGATIVE B2 - CARDS.csv`**. `primary_key` retains Account Key and `cardnumber` retains Contract Number.

The first valid email and first valid mobile come from the matched database record. **EMAIL (01)** supplies the preferred email and **Mobile 1** supplies the preferred phone, regardless of their positions in the worksheet. Other recognized Email/Phone, Mobile, Cellphone, Home Phone, and Business Phone fields are checked in source order when the preferred value is missing or invalid, or the preferred header is absent. Header matching ignores case, spaces and punctuation. Numbers must normalize to ten digits starting with 9. Errors or uncached formulas in contact cells count as invalid contacts, allowing other contact fields to be checked. Matching/classification cell errors must be corrected. Duplicate DataGrid accounts are excluded. Each exported account has **one row**, with both contacts in `email` and `phone`; an account with only one valid contact is still included. Accounts without valid contacts produce no row.

DataGrid Old I.C. supplies `indexoneuid`. Database Name, Total Outstanding, Assignment Date, and Area supply `FullName`, `OB`, `EndoDate`, and `AREA`; recognized DataGrid fields supply these when database values are missing. `Product` is CARDS or BEL and `doclass` is SBC CARDS B2 or SBC BEL B2. **Monthly contains the assigned agent name and Terms contains the matching agent code**, using the even assignment below. Level, ChristmasAmnesty, and NewEndo are copied only when explicitly present; other unavailable fields stay blank. Dates use ISO text. B2 blasting exports have no activity statuses, remarks, author, or log timestamp; the B2 files panel prepares activity logs from those exports.

Review matching and contact counts, preview either group, download the two CSVs or their ZIP, or save both to a folder. B2 names are configurable in `AUTOSTAT_B2_FILENAMES` in `export_config.py`.

### B4 matching on the NEGATIVE B2 page

The **NEGATIVE B4 — match accounts and contacts** section uses the same matching rules with independent uploads and results. Upload **SPM ENDORSEMENT B4 as of 29OCTOBER2026.xlsx** as the **B4 endorsement database** and **NEGATIVE B4 DB.xlsx** as the **B4 account list**. Daily filenames may change. Account No. matches Account Key; Financing/Card No. matches Contract Number, with both required to agree when both are supplied. Only database Bucket Code `4` / `B4` qualifies. Other buckets, missing identifiers, ambiguous matches and duplicates are excluded and reported.

B4 uses exactly **Email (01) → email** and **Phone (01) → phone**. Both headers must be present. Invalid or blank 01 contacts stay blank; other Email, Phone and Mobile columns are not substituted. The existing email validation and mobile normalization apply. Accounts without a valid 01 contact are excluded. Outputs are **NEGATIVE B4 - CARDS.csv** and **NEGATIVE B4 - BEL.csv**, using the same 16-column template with `primary_key` from Account Key, `cardnumber` from Contract Number, and `doclass` set to SBC CARDS B4 or SBC BEL B4. B4 filenames are configurable in `AUTOSTAT_B4_FILENAMES`.

For **both B2 and B4**, each CARDS/BEL output alternates its exported accounts between the two agents below. Assignment happens after matching, deduplication and contact exclusions, so each agent gets the same number of accounts, or a one-account difference. Almera receives the extra account for an odd total. Names and codes stay paired on the same row, and each group's split starts independently. The summary shows each agent's account count. Existing collector assignments in the source do not control this requested split.

| Monthly | Terms |
| --- | --- |
| Almera Ranceli | AMRANCELI |
| Kyarill Palima | KPALIMA |

The agent pairs are configured in `COLLECTION_AGENTS` in `export_config.py`. Feeding a B2 blasting file into B2 AutoStat copies Terms into Collector and the email remarks. **Clear all files** clears B2 and B4 matching uploads, results and source selections together.

Some Excel exports incorrectly save their worksheet range as `A1:C1` despite containing more columns and thousands of rows. Both readers now use the actual worksheet cells, recovering these hidden headers and data without modifying the original files.

## Deploy on Streamlit Community Cloud

1. Sign in at [share.streamlit.io](https://share.streamlit.io/) with the GitHub account that can access the repository.
2. Click **Create app** and select the existing GitHub repository:
   - Repository: `Michael081503/BlastingAutomation`
   - Branch: `main`
   - Main file path: `app.py`
3. In **Advanced settings**, select **Python 3.13**, the version used for local testing.
4. Click **Deploy**.

See Streamlit's [official deployment instructions](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy). The pinned dependencies in `requirements.txt` install automatically. No secrets, API keys, or database setup are needed for this app.

Upload the source workbook and template in the deployed app. Input workbooks, generated files, environment files, and Streamlit secrets are excluded from Git; the repository contains the application code and tests.

Startup refreshes older modules held in memory when a deployment adds required settings, including GEO configuration. Routine reruns retain the same module classes and settings. If startup still reports missing names after an update, check that the app uses `main` and `app.py` from this repository, then choose **Manage app → ⋮ → Reboot app** to clear the cloud process and load the deployed files. See Streamlit's [reboot instructions](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app/reboot-your-app).

On Community Cloud, **Save files to folder** writes to the cloud server, so it cannot create a folder directly on your PC. Use individual file downloads or the ZIP download, then extract the ZIP into a folder on your PC. Cloud-generated files are temporary; see Streamlit's [storage guidance](https://docs.streamlit.io/develop/concepts/connections/connecting-to-data). Running the app locally enables direct saves to your own folders.

## Use

1. Upload your source `.xlsx` workbook.
2. Upload your template (`.xlsx` or `.csv`), or use the built-in blank CSV template with all 16 campaign headers. Both template formats produce CSV outputs.
   **CSV file type** defaults to **CSV (Comma delimited)**, using Windows-1252 without a UTF-8 marker. Choose **CSV UTF-8 (Comma delimited)** when your data needs characters beyond that format. Both choices use commas and `.csv` filenames.
3. Choose the source sheet. **NEGATIVE** is selected automatically.
4. Choose the destination sheet if your template has multiple sheets.
5. Check **Header rows** if the detected row is incorrect.
6. For worksheets with grouping columns, keep **Separate Christmas Amnesty groups** selected to create the groups below. Worksheets with Area default to **Separate Area and Christmas Amnesty groups** instead. If grouping columns are absent, the app automatically uses **All accounts in one group** and copies the available values. Every group is still split into Email and Viber files. Each worksheet remembers its own output-mode selection.
7. Click **Preview and prepare files**, then review the email-validation counts and export previews.
8. Under **Save to a folder**, enter an output location and click **Save files to folder**. The app displays the full path of the created folder. Individual downloads and **Download all files (ZIP)** are also available.

Changing a file, sheet, header row, export mode, or filename configuration clears the previous result. Prepare a new download after changing inputs. Filenames from `export_config.py` apply to individual downloads, ZIP contents, previews, and folder saves. Duplicate or invalid filenames must be corrected in that file before preparing exports.

After downloading the ZIP, right-click it and choose **Extract All**, then open the CSV files from the extracted folder to edit them. Opening a CSV directly inside the ZIP can make Excel show a read-only temporary copy and ask you to save another file. Individual CSV downloads and local folder saves can also be opened from their saved location. The app does not set read-only permissions or protect the exported files. See Microsoft's [ZIP extraction instructions](https://support.microsoft.com/en-us/windows/experience/storage-filemanagement/zip-and-unzip-files).

## Save directly to a folder

The default destination is the project's `output` folder. Each save creates a new folder, for example `output/your_workbook_NEGATIVE_2026-10-26_21-00-00/`, containing the separate CSV files directly. Dates and times use Asia/Manila. An existing export is never overwritten, even when you save twice in the same second. No ZIP is written into this folder.

You can change the **Output folder** field in the app or edit `OUTPUT_FOLDER` in [export_config.py](export_config.py). Relative paths start in the project folder; an absolute path selects another location. Files are saved on the computer running Streamlit: your computer when running locally, or the server when deployed online.

Files are written only after clicking **Save files to folder**. Preparing or previewing the files does not save them. The folder option works with Excel or CSV templates and exports that keep all accounts in one group. Folder saves accept only CSV files.

## Email and Viber files

Every nonempty group creates two CSV files using the same template and all its original headers:

- **Valid email only:** included in `_Email` only.
- **Phone only:** included in `_Viber` only.
- **Valid email and phone:** included in both files.
- **Invalid email and phone:** included in `_Viber` only.
- **No valid email and no phone:** excluded from both files.

The **email and phone headers remain in both files**. In `_Email` files, the `phone` column is blank; in `_Viber` files, the `email` column is blank. All other mapped values remain intact. Source rows are not changed, and their order is retained within each file.

For example, the base name `CS AMNESTY NEGA CARDS - L1` produces:

```text
CS AMNESTY NEGA CARDS - L1_Email.csv
CS AMNESTY NEGA CARDS - L1_Viber.csv
```

If a group has accounts but one channel has no matching accounts, that channel's file still contains the template headers with no account data. A group with no accounts produces no files. This behavior applies to qualified, blank, review, and other-sheet exports, including folder saves, individual downloads, and ZIP downloads.

The app shows account counts separately from Email and Viber row counts. Accounts with both contact details contribute a row to each channel, so adding the channel counts may exceed the number of included source accounts.

### Email validation

Email addresses are checked automatically before export using `email-validator`. Leading and trailing whitespace is trimmed, and valid addresses are normalized. Internal spaces, consecutive dots, dots at the start or end of the local part, missing `@` signs, incomplete domains, and other invalid formats are removed from Email exports. Valid dots and plus tags are retained; the app does not guess corrections to malformed addresses.

Every nonempty source account is checked, including accounts later excluded for having no valid contact details. The app displays **Accounts checked for email**, **Valid email formats**, **Invalid emails removed**, and **Missing emails**. The three email-status counts add up to the accounts checked; entirely empty rows are counted separately.

An account with an invalid email still goes into its Viber file when it has a phone number. Without a phone number, the account is excluded from both files. Invalid-address counts include accounts excluded entirely. Blank email cells are counted as missing, not invalid. This applies to GEO groups, standard NEGATIVE groups, and single-group CSV exports with either template format, ZIP downloads, and folder saves. Source workbooks are never changed.

Validation runs locally without DNS queries, sending messages, or contacting mail servers. It checks the address format, not whether the individual mailbox exists or whether a message will be delivered. A well-formed address may still bounce.

## Worksheet separation

### Daily RECO agent cycle

The **RECO agent cycle** panel shows the assignment date, agent names for `Monthly`, codes for `Terms`, and the upcoming three-day order. It appears before uploading files and when a RECO worksheet is selected. GEO does not use this three-agent split or rotation. Starting **30 October 2026**, the saved default order is:

| Monthly | Terms |
| --- | --- |
| Andrea Mamaril | APMAMARIL |
| Rosalie Vinoya | ROSVINOYA |
| Jenayrine Bataller | JBATALLER |

The first agent moves to the end each calendar day in **Asia/Manila**: 30 October starts with APMAMARIL, 31 October with ROSVINOYA, and 1 November with JBATALLER, then the cycle repeats. The assignment date defaults to the current Manila date each day; select another date to preview or generate its assignments.

Open **Change RECO agent cycle** in the sidebar to change the start date or reorder the three agents, then click **Save agent cycle**. Each agent must appear exactly once, and names and codes stay paired. The saved RECO cycle survives **Clear all files** and new browser sessions. Existing saved settings remain in `settings/geo_reco_agent_cycle.json` on the app's computer/server, separately from Git; back up this file when moving or replacing the server's storage. The defaults remain in `export_config.py`.

Each separated RECO product/status/level group distributes distinct accounts evenly across the day's order after sorting by OB, highest first. Counts differ by at most one account per agent. Email and Viber retain the same agent for the same account, and duplicate source rows within a group keep the assignment. Individual channel counts can differ because some accounts have only one contact method. This also applies to RECO's all-accounts mode, previews, CSVs, Excel files, ZIPs and folder saves. Changing the date or saved order clears previous RECO campaign downloads. Missing `Monthly` or `Terms` headers are appended to RECO templates. GEO keeps its previous fields without assigning Andrea, Rosalie or Jenayrine; GEO downloads do not depend on RECO cycle settings. B2/B4 retain their separate two-agent assignments.

**GEO and RECO normalize Product first:** `MC` becomes `CARDS` before any Area, Christmas Amnesty, product, level, or Email/Viber split. Matching ignores capitalization and surrounding spaces, so `mc` and ` MC ` also become `CARDS`. These accounts follow the existing CARDS destination rules, and the exported Product value is `CARDS` in previews, CSVs, Excel files, ZIPs, and folder saves, including all-accounts mode. Other products retain their values. Original uploaded workbooks are unchanged.

The following rules apply to NEGATIVE and all other source sheets without an Area column. Select one worksheet at a time using **Source sheet**. NEGATIVE keeps its configured filenames. Other worksheets prefix the same configured group filenames and ZIP name with their worksheet name, so their exports are distinguishable. For example, a sheet called `POSITIVE` creates `POSITIVE - CS AMNESTY NEGA CARDS - L1_Email.csv` for that populated group. The configured base names remain editable in `export_config.py`. Templates, validation, review groups, previews, ZIP downloads, and folder saves work identically for all worksheets.

| Christmas Amnesty | Product | Level | Base filename |
| --- | --- | --- | --- |
| QUALIFIED | CARDS | Level 1 | CS AMNESTY NEGA CARDS - L1 |
| QUALIFIED | CARDS | Level 6 | CS AMNESTY NEGA CARDS - L6 |
| QUALIFIED | BEL | All levels | CS AMNESTY NEGA BEL - PL |
| Blank / not qualified | CARDS | Level 1 | NEGATIVE_BLANK_CARDS_LEVEL_1 |
| Blank / not qualified | CARDS | Level 6 | NEGATIVE_BLANK_CARDS_LEVEL_6 |
| Blank / not qualified | BEL | All levels | NEGATIVE_BLANK_BEL |
| Blank / not qualified | Other than CARDS and BEL | All levels | NEGATIVE_BLANK_OTHER_PRODUCTS |

The template's `ChristmasAmnesty` column contains `QUALIFIED` when the source is qualified; otherwise it is blank. Matching ignores case and surrounding whitespace. `LEVEL 1` and `Level 1` belong to the same group. BEL includes every level.

Accounts are removed when **there is no valid `debtor_email` and `phone_1` is empty**. An account with a valid email or a phone number is kept in its matching channel file(s). Whitespace-only cells count as empty. This condition applies to qualified and blank groups, as well as other source sheets. The app displays the number of accounts removed separately from entirely empty rows and the number of invalid emails removed.

Empty groups produce no file. Qualified rows outside the product/level rules and blank-status CARDS at levels other than 1 or 6 are put in separate, clearly identified review groups with Email/Viber files. They are never mixed into the other-products group. Every account with at least one contact field is assigned to exactly one group, then routed to its matching channel(s).

Before email validation, the sample dated October 26, 2026 had qualified group counts of **8**, **4,277**, and **131**, and blank group counts of **3,066**, **2,155**, **121**, and **4,059**. It included **13,817** source accounts after removing **541** without either contact field, with **12,414 Email rows** and **13,489 Viber rows**. Validation can reduce Email counts and remove accounts with an invalid email and no phone; prepare the sample again to see its current counts.

## GEO Area separation

Use **Download Excel files (ZIP, keeps account numbers as text)** to open or edit GEO/RECO accounts in Excel, or to pass them into AUTOSTAT. Each campaign file has an `.xlsx` copy with the same headers, full rows, and order. **Numeric balances and levels retain their source number types, and dates remain Excel dates**. Names, email addresses, labels and headers use General formatting. Only `primary_key`, `cardnumber`, `indexoneuid` and `phone` use Text formatting to preserve digits and leading zeros. Text values from the source stay literal text, including values beginning with `=`, and extra template columns stay blank. **Download Excel for this export** saves just the file selected in Preview. Workbooks are generated when downloaded. The existing CSVs and ZIP remain under **CSV files for system upload**, and folder saves still write CSVs. Opening and resaving those CSVs in Excel can round long identifiers; use the `.xlsx` copies instead. An Excel copy preserves the supplied values but cannot restore digits already missing from a rounded source.

GEO and RECO exports sort every Email and Viber file by **OB, highest to lowest**, including review groups and all-accounts mode. Equal amounts keep source order; blank or invalid amounts go last. Previews, CSV downloads, ZIPs, and folder saves all use the same order. Source values are preserved.

For GEO workbooks, the app detects the **Area** source column and defaults to **Separate Area and Christmas Amnesty groups**. `Mindanao` and `Region 3` / `R3` form separate groups, ignoring case, spacing, and punctuation. Each Area uses these five main groups, combining **all levels** within each product:

| Christmas Amnesty | Product |
| --- | --- |
| QUALIFIED | CARDS |
| QUALIFIED | BEL |
| Blank / not qualified | CARDS |
| Blank / not qualified | BEL |
| Any status for CSL/PL/SLA; blank / not qualified for other loan products | LOANS |

Every nonempty group then produces Email and Viber CSV files, so five populated groups in each Area produce **20 files**. **CSL, PL and SLA always go directly to the area's existing LOANS files:** `GEO NEGA LOANS - MINDANAO_Email/Viber.csv` or `GEO NEGA LOANS - R3_Email/Viber.csv`. This routing takes priority over Christmas Amnesty grouping, including QUALIFIED accounts. The Product and ChristmasAmnesty columns retain each account's source values. Level values are retained and do not affect GEO grouping. Unrecognized qualified products, and other or missing Area values, are kept in clearly named review files when present. RECO retains its existing grouping rules.

The GEO source headers `email`, `old_ic`, `account`, `card_no`, `assign_date`, `placement`, `Product`, `OB (MASTERLIST}`, and `Phone 1` are accepted. `assign_date` supplies `EndoDate`; `placement` supplies `Level`. The source Area is copied into `AREA`, preserving `Region 3` as the source value even though its group is named R3. The `doclass` template column is filled automatically with **`SBC GEO Mindanao`** for Mindanao accounts and **`SBC GEO R3`** for Region 3 / R3 accounts, ignoring capitalization, spacing, and punctuation in Area. This applies to both Email and Viber files, previews, ZIPs, folder saves, and the all-accounts mode. Other or missing Areas leave `doclass` blank. Area is optional for existing workbooks. GEO grouped exports require an AREA template header; the built-in 16-column template already includes AREA and doclass.

GEO grouping works on each worksheet with an Area column. **NEGATIVE** remains the default; select one source sheet at a time. Other GEO sheets prefix the configured filenames and ZIP name with their worksheet name. **OVERALL** and the other sheets may overlap, so processing them separately does not create a deduplicated combined export.

Edit the GEO filenames directly in `GEO_OUTPUT_FILENAMES` at the top of [export_config.py](export_config.py). Mindanao and R3 each have five entries, followed by the review filenames. Each quoted value is the full base name for that Area and group. The app appends `_Email` or `_Viber` and `.csv`. For example, `CS AMNESTY GEO CARDS_MINDANAO` produces `CS AMNESTY GEO CARDS_MINDANAO_Email.csv` and `CS AMNESTY GEO CARDS_MINDANAO_Viber.csv`. Change the ZIP name in `GEO_ZIP_FILENAME`.

## Edit filenames

Open **[export_config.py](export_config.py)** and edit the quoted filenames on the right. For GEO, the entries are listed separately for Mindanao and R3, for example:

```python
"mindanao_qualified_cards": "CS AMNESTY GEO CARDS_MINDANAO",
"r3_qualified_cards": "CS AMNESTY GEO CARDS_R3",
```

The standard NEGATIVE filenames are in the same file:

```python
OUTPUT_FILENAMES = {
    "qualified_cards_level_1": "CS AMNESTY NEGA CARDS - L1",
    "qualified_cards_level_6": "CS AMNESTY NEGA CARDS - L6",
    "qualified_bel": "CS AMNESTY NEGA BEL - PL",
    "blank_cards_level_1": "NEGATIVE_BLANK_CARDS_LEVEL_1",
    "blank_cards_level_6": "NEGATIVE_BLANK_CARDS_LEVEL_6",
    "blank_bel": "NEGATIVE_BLANK_BEL",
    "blank_other_products": "NEGATIVE_BLANK_OTHER_PRODUCTS",
    "qualified_review": "NEGATIVE_QUALIFIED_REVIEW",
    "blank_review": "NEGATIVE_BLANK_REVIEW",
}
CHANNEL_SUFFIXES = {"email": "_Email", "viber": "_Viber"}
ZIP_FILENAME = "NEGATIVE_CHRISTMAS_AMNESTY_FILES.zip"
OUTPUT_FOLDER = "output"
```

Keep the keys on the left unchanged. Edit the base names in `GEO_OUTPUT_FILENAMES` or `OUTPUT_FILENAMES`; the app adds `_Email` or `_Viber` before `.csv`. You can edit these suffixes separately in `CHANNEL_SUFFIXES`. Final filenames must be unique and must not contain folder paths. Save the code, rerun the app, and prepare the files again to use the updated names.

## Mapping

| Source header | Template header |
| --- | --- |
| debtor_email | email |
| OLD I.C | indexoneuid |
| Account no. | primary_key |
| financing/Card no. | cardnumber |
| Name | FullName |
| endo date | endodate |
| level | level |
| product | product |
| OB (Masterlist) | OB |
| phone_1 | phone |
| Christmas Amnesty | ChristmasAmnesty |
| Area (optional) | AREA |
| Area (calculated) | doclass |

Header matching ignores capitalization, whitespace, and punctuation. Source files accept both the original source names and the template names, including `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `EndoDate`, `email`, and `phone`. `Product Type` is also accepted for `product`, and `OB` is accepted for `OB (Masterlist)`. Phone aliases include `Phone number`, `Mobile`, `Mobile number`, and `Contact number`. GEO aliases are listed above.

Source files need an account identifier column (`Account no.` / `primary_key` or `financing/Card no.` / `cardnumber`) and at least one email or phone column. Other missing source fields stay blank in the template, and the app lists them before export. It never copies one identifier into a different identifier field or invents missing values. **Source columns** shows the actual headers and sample values. Select **Map source columns manually** to map differently named columns or resolve ambiguous headers; the controls open automatically when account/contact columns are not recognized. Mapping changes clear previous downloads. Account identifiers and contacts must use separate columns. Standard grouping requires Christmas Amnesty, product, and level; GEO grouping requires Area, Christmas Amnesty, and product. When grouping columns are missing, the app prepares a single Email file and a single Viber file with the available values. Uploaded templates still require the eleven mapped destination headers. Header detection scans the first 50 rows; change the header row manually when headers appear later. A usable account/contact sheet is preferred over summary sheets when choosing the default.

The built-in template uses your original 16-column layout: `email`, `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `phone`, `OB`, `Monthly`, `Terms`, `doclass`, `Level`, `Product`, `EndoDate`, `AREA`, `ChristmasAmnesty`, and `NewEndo`. `Monthly` and `Terms` come from the daily RECO agent cycle for RECO exports and stay blank for GEO; `NewEndo` remains blank. AREA is filled when present in the source and otherwise stays blank. `doclass` is calculated from Area for Mindanao and R3; other or missing Areas leave it blank. No source doclass column is required. Uploaded templates keep their own header labels and order; missing Monthly/Terms columns are appended for RECO agent assignment. When they include a doclass header, it is filled by the same rule. Export previews show all file headers, including columns with blank values and channels with no rows.

## Output behavior

- Source order is retained within each output group. Hidden and filtered source rows are included.
- Rows with no values in any mapped field are skipped. Accounts with neither a valid email nor a phone are removed and counted separately. Invalid emails are removed and counted. Duplicates are retained. Accounts with only a valid email or only a phone go to that channel; accounts with both go to both channels.
- Campaign Email/Viber files are CSV, including individual downloads, files inside ZIP downloads, and folder saves. The blank template download is also CSV. AUTOSTAT additionally offers its text-formatted Excel download.
- IDs, account numbers, card numbers, and phones retain existing text leading zeros and simple Excel zero masks such as `000000` in the CSV data. Digits already lost in the original Excel file cannot be recovered. CSV has no text column types: import identifier columns as text when opening it in Excel to preserve leading zeros.
- Dates use ISO text and balances use their source values without currency formatting. The default **CSV (Comma delimited)** uses Windows-1252 without a BOM and Windows line endings. **CSV UTF-8 (Comma delimited)** uses UTF-8 with BOM. Characters unsupported by the standard format produce an actionable error; they are never silently replaced or dropped.
- Excel templates supply the selected sheet's header labels and column order from the chosen header row. Their CSV outputs use commas. Content above the header, stale template rows, styles, and other worksheets are omitted. Extra template columns are blank. The original template is never changed.
- CSV templates supply their first-row headers and column order. Comma, semicolon, tab, and pipe delimiters are accepted as inputs; outputs always use commas. Extra template columns are blank.
- Formula source cells use values cached by Excel. A mapped formula without a saved result stops processing; recalculate and save the source in Excel before uploading again.
- Use ordinary `.xlsx` or `.csv` templates with a single, unambiguous header row containing all mapped fields. Excel formatting and workbook objects are not included in CSV outputs.
- Uploaded files and prepared outputs are held in memory. Clicking **Save files to folder** writes output copies to a new folder. Originals are never overwritten. Files are not emailed or sent to external services.

## Verification

```powershell
python -m unittest discover -s tests -v
```

The tests use generated sample data, including CSV output from Excel and CSV templates, reordered headers, leading-zero identifiers, multiple sheets, stale template rows, Excel formulas/errors, all qualified/blank group boundaries, email format validation and removal, channel routing, retained contact headers, header-only channels, contact exclusions, folder saves, ZIP contents, and the Streamlit upload-to-download flow.

The app uses Streamlit's [file uploader](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader) and [download button](https://docs.streamlit.io/develop/api-reference/widgets/st.download_button).
