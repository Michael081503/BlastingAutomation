# Blasting Automation

A local Streamlit app that copies data from an Excel worksheet into your template's column layout. It starts with **NEGATIVE** and applies the same separation rules to every selected worksheet: Christmas Amnesty, product, and level, followed by **Email** and **Viber CSV files** for every group. GEO worksheets with an **Area** column first separate **Mindanao** and **R3**, combining all levels within each product. Outputs are always `.csv`, whether the template is Excel, CSV, or the built-in blank template.

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
- **NEGATIVE B2**: DataGrid matching against the latest SBC database and the two B2 blasting files.
- **AUTOSTAT**: all uploaded account lists compiled into one Email activity file and one Viber activity file.

Workbooks stay on the page you select. A DataGrid-looking sheet does not switch GEO/RECO into B2 matching.

## AUTOSTAT

AUTOSTAT checks account lists for valid email addresses and mobile numbers, then compiles every selected file and sheet into **`AUTOSTAT_Email.csv`** and **`AUTOSTAT_Viber.csv`** using the supplied `TEMPLATE_EMAIL.xlsx` and `TEMPLATE_VIBER.xlsx` layouts. Both layouts are built in; the sample account, email address, and old timestamp are not included. Outputs are CSV files, as in the existing workflow.

1. Open **AUTOSTAT** in the sidebar and upload one or more `.xlsx` or `.csv` source files.
2. Review **Check source columns**. All workbook sheets with recognized account and contact headers are selected by default. For other layouts, select the sheet, set its header row, and choose its account, email, and phone columns. Several email or phone columns can be selected. Account numbers must use a separate column from contacts.
3. Check **Remark By**, which defaults to `MBURGOS`. The log timestamp defaults to the time you prepare the files in Asia/Manila. Uncheck **Use current Manila date and time** to enter a date and time.
4. Click **Preview and prepare AUTOSTAT files**. Review the per-sheet counts for invalid contacts, missing accounts, and rows with no valid contacts, then preview the logs.
5. Download the two compiled CSVs, download both as a ZIP, or save them directly to a folder on the computer running the app.

Email validation uses the same local format checks as the existing workflow. For phones, AUTOSTAT removes spaces and common phone formatting, keeps the **rightmost 10 digits**, and accepts the result only when it **starts with 9**. For example, `09171234567`, `+63 (917) 123-4567`, and `9171234567` become `9171234567`. Short numbers, results beginning with another digit, and text containing letters are rejected. Numeric Excel phones and integral CSV numbers with `.0` or scientific notation are supported.

AUTOSTAT automatically selects **`cardnumber`** as the source for **`Account No.`** in both logs, taking priority over `primary_key` and other account headers when present. Header matching ignores case, spaces, and punctuation. When `cardnumber` is absent, the existing account-header detection and manual column selection remain available. A blank `cardnumber` value is treated as a missing account; another identifier from the same row is never substituted automatically.

AUTOSTAT exports each account number once per channel, keeping its first valid email and first valid phone. Repeated accounts across selected rows, sheets, and files are removed independently for Email and Viber. The per-sheet summary shows **Duplicate Email rows removed** and **Duplicate Viber rows removed** alongside the exported channel counts. Account matching trims surrounding spaces and ignores letter case while preserving leading zeros and punctuation. Multiple contacts in one cell can be separated by commas, semicolons, pipes, or line breaks; the first valid contact in the selected column order is kept for each channel. Missing account numbers are excluded. An account with valid email and phone details appears in both logs; an account with neither appears in neither. Source data is not changed.

All selected sources contribute to the same two files, in upload and selected-sheet order. BEL, CARDS, and other accounts are compiled together by channel. Duplicate filenames do not create additional output files. If a channel has no valid contacts, its file contains headers only. The Email output keeps all 11 supplied headers, including the extra reference header at the end; unused fields stay blank. The Viber output keeps its five headers. The Email remark formula is rendered as text with the validated address; the Viber `NUMBER` placeholder is replaced with the normalized phone. Template statuses are copied exactly, and Remark By and Time are filled from the page settings. The per-source summary still identifies each contributing file and sheet.

Preparing these logs does not send emails or Viber messages. Source files are never modified. Changing source files, sheets, columns, header rows, remark author, timestamp settings, or CSV encoding clears the prepared result so it can be rebuilt from the new inputs.

## NEGATIVE B2: match DataGrid with the SBC database

Open **NEGATIVE B2** and upload the full SBC database and the DataGrid account list. The page selects the full `SBC` sheet by its required headers. DataGrid defines the account population; database accounts absent from DataGrid are excluded.

B2 matching runs only on its dedicated page. Daily filenames can change because matching uses source columns and file contents. Both B2 files use the exact 16-column layout from **1EMAIL_Blasting - Template.xlsx**: `email`, `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `phone`, `OB`, `Monthly`, `Terms`, `doclass`, `Level`, `Product`, `EndoDate`, `AREA`, `ChristmasAmnesty`, `NewEndo`. Sample rows from the template are never included.

Upload the latest database each day, then click **Match and prepare B2 files**. Replacing either upload clears previous results and reads the new file contents, including when a replacement has the same filename. The app uses the database you upload; it does not synchronize a changing database on your computer automatically.

DataGrid `Account No.` matches database `Account Key`, and `Financing/Card No.` matches database `Contract Number`. Both identifiers must identify the same record when both are supplied. Matching retains leading zeros and uses exact identifiers after trimming spaces and ignoring letter case. Unmatched, conflicting, ambiguous, and missing identifiers appear in the matching report. Matching never uses customer names. Only database Bucket Code `2` / `B2` qualifies; other or missing buckets are excluded. Product Category Code `BEL` creates **`NEGATIVE B2 - BEL.csv`**; supported card categories such as `MC` create **`NEGATIVE B2 - CARDS.csv`**. `primary_key` retains Account Key and `cardnumber` retains Contract Number.

The first valid email and first valid mobile come from the matched database record. Numbered Email/Phone fields and recognized Mobile, Cellphone, Home Phone, and Business Phone columns are checked in source order. Numbers must normalize to ten digits starting with 9. Errors or uncached formulas in contact cells count as invalid contacts, allowing other contact fields to be checked. Matching/classification cell errors must be corrected. Duplicate DataGrid accounts are excluded. Each exported account has **one row**, with both contacts in `email` and `phone`; an account with only one valid contact is still included. Accounts without valid contacts produce no row.

DataGrid Old I.C. supplies `indexoneuid`. Database Name, Total Outstanding, Assignment Date, and Area supply `FullName`, `OB`, `EndoDate`, and `AREA`; recognized DataGrid fields supply these when database values are missing. `Product` is CARDS or BEL and `doclass` is SBC CARDS B2 or SBC BEL B2. **Monthly and Terms remain blank**, as requested. Level, ChristmasAmnesty, and NewEndo are copied only when explicitly present; other unavailable fields stay blank. Dates use ISO text. B2 exports have no activity statuses, remarks, author, or log timestamp. AUTOSTAT compiles activity logs on its own page.

Review matching and contact counts, preview either group, download the two CSVs or their ZIP, or save both to a folder. B2 names are configurable in `AUTOSTAT_B2_FILENAMES` in `export_config.py`.

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

For GEO workbooks, the app detects the **Area** source column and defaults to **Separate Area and Christmas Amnesty groups**. `Mindanao` and `Region 3` / `R3` form separate groups, ignoring case, spacing, and punctuation. Each Area uses these five main groups, combining **all levels** within each product:

| Christmas Amnesty | Product |
| --- | --- |
| QUALIFIED | CARDS |
| QUALIFIED | BEL |
| Blank / not qualified | CARDS |
| Blank / not qualified | BEL |
| Blank / not qualified | Other than CARDS and BEL |

Every nonempty group then produces Email and Viber CSV files, so five populated groups in each Area produce **20 files**. Level values are retained in the output column and do not affect GEO grouping. Qualified products other than CARDS and BEL, and other or missing Area values, are kept in clearly named review files when present. The standard NEGATIVE workflow without an Area column retains its original level rules.

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

The built-in template uses your original 16-column layout: `email`, `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `phone`, `OB`, `Monthly`, `Terms`, `doclass`, `Level`, `Product`, `EndoDate`, `AREA`, `ChristmasAmnesty`, and `NewEndo`. The three columns without mapping rules (`Monthly`, `Terms`, and `NewEndo`) are retained and left blank. AREA is filled when present in the source and otherwise stays blank. `doclass` is calculated from Area for Mindanao and R3; other or missing Areas leave it blank. No source doclass column is required. Uploaded templates keep their own header labels and order; when they include a doclass header, it is filled by the same rule. Export previews show all file headers, including columns with blank values and channels with no rows.

## Output behavior

- Source order is retained within each output group. Hidden and filtered source rows are included.
- Rows with no values in any mapped field are skipped. Accounts with neither a valid email nor a phone are removed and counted separately. Invalid emails are removed and counted. Duplicates are retained. Accounts with only a valid email or only a phone go to that channel; accounts with both go to both channels.
- All generated Email/Viber files are CSV, including individual downloads, files inside ZIP downloads, and folder saves. The blank template download is also CSV.
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
