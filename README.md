# Blasting Automation

A local Streamlit app that copies data from an Excel worksheet into your template's column layout. It starts with **NEGATIVE**, separates rows by Christmas Amnesty, product, and level, then creates **Email** and **Viber CSV files** for every group. GEO worksheets with an **Area** column first separate **Mindanao** and **R3**. Other sheets without Area also get Email/Viber CSV files while their grouping rules are pending. Outputs are always `.csv`, whether the template is Excel, CSV, or the built-in blank template.

## Run

Requires Python 3.11 or newer. In PowerShell, from this folder:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open **http://localhost:8501**. If the dependencies are already installed, run `python -m streamlit run app.py` directly. Stop the server with **Ctrl+C**.

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
6. For NEGATIVE, keep **Separate Christmas Amnesty groups** selected to create the groups below. Worksheets with Area default to **Separate Area and Christmas Amnesty groups** instead. **All accounts in one group** skips these group filters. Every group is still split into Email and Viber files.
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

## NEGATIVE separation

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

The GEO source headers `email`, `old_ic`, `account`, `card_no`, `assign_date`, `placement`, `Product`, `OB (MASTERLIST}`, and `Phone 1` are accepted. `assign_date` supplies `EndoDate`; `placement` supplies `Level`. The source Area is copied into `AREA`, preserving `Region 3` as the source value even though its group is named R3. Area is optional for existing workbooks. GEO grouped exports require an AREA template header; the built-in 16-column template already includes it.

GEO grouping works on each worksheet with an Area column. **NEGATIVE** remains the default; select one source sheet at a time. **OVERALL** and the other sheets may overlap, so processing them separately does not create a deduplicated combined export.

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

Header matching ignores capitalization, whitespace, and punctuation. `Product Type` is also accepted for `product`, matching your sample workbook. `OB` is also accepted as a source header for `OB (Masterlist)`. GEO aliases are listed above. All eleven required source and destination headers must be present; Area is optional outside grouped GEO exports. Missing or ambiguous headers are reported before export, including duplicate optional Area headers. Header detection scans the first 50 rows; enter a row manually when headers appear later.

The built-in template uses your original 16-column layout: `email`, `primary_key`, `indexoneuid`, `cardnumber`, `FullName`, `phone`, `OB`, `Monthly`, `Terms`, `doclass`, `Level`, `Product`, `EndoDate`, `AREA`, `ChristmasAmnesty`, and `NewEndo`. The four columns without mapping rules (`Monthly`, `Terms`, `doclass`, and `NewEndo`) are retained and left blank. AREA is filled when present in the source and otherwise stays blank. Uploaded templates keep their own header labels and order. Export previews show all file headers, including columns with blank values and channels with no rows.

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
