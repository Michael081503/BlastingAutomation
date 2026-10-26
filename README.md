# Blasting Automation

A local Streamlit app that copies data from an Excel worksheet into your template. It starts with **NEGATIVE**, separates rows by Christmas Amnesty, product, and level, then creates **Email** and **Viber** files for every group. Other sheets also get Email/Viber files while their grouping rules are pending.

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

On Community Cloud, **Save files to folder** writes to the cloud server, so it cannot create a folder directly on your PC. Use individual file downloads or the ZIP download, then extract the ZIP into a folder on your PC. Cloud-generated files are temporary; see Streamlit's [storage guidance](https://docs.streamlit.io/develop/concepts/connections/connecting-to-data). Running the app locally enables direct saves to your own folders.

## Use

1. Upload your source `.xlsx` workbook.
2. Upload your template (`.xlsx` or `.csv`), or use the built-in blank template.
3. Choose the source sheet. **NEGATIVE** is selected automatically.
4. Choose the destination sheet if your template has multiple sheets.
5. Check **Header rows** if the detected row is incorrect.
6. For NEGATIVE, keep **Separate Christmas Amnesty groups** selected to create the groups below. **All accounts in one group** skips these group filters. Every group is still split into Email and Viber files.
7. Click **Preview and prepare files** and review the group counts and previews.
8. Under **Save to a folder**, enter an output location and click **Save files to folder**. The app displays the full path of the created folder. Individual downloads and **Download all files (ZIP)** are also available.

Changing a file, sheet, header row, export mode, or filename configuration clears the previous result. Prepare a new download after changing inputs.

## Save directly to a folder

The default destination is the project's `output` folder. Each save creates a new folder, for example `output/your_workbook_NEGATIVE_2026-10-26_21-00-00/`, containing the separate Excel files directly. Dates and times use Asia/Manila. An existing export is never overwritten, even when you save twice in the same second. No ZIP is written into this folder.

You can change the **Output folder** field in the app or edit `OUTPUT_FOLDER` in [export_config.py](export_config.py). Relative paths start in the project folder; an absolute path selects another location. Files are saved on the computer running Streamlit: your computer when running locally, or the server when deployed online.

Files are written only after clicking **Save files to folder**. Preparing or previewing the files does not save them. The folder option also works with CSV templates and exports that keep all accounts in one group.

## Email and Viber files

Every nonempty group creates two files using the same template and all its original headers:

- **Email only:** included in `_Email` only.
- **Phone only:** included in `_Viber` only.
- **Both email and phone:** included in both files.
- **Neither email nor phone:** excluded from both files.

The **email and phone headers remain in both files**. In `_Email` files, the `phone` column is blank; in `_Viber` files, the `email` column is blank. All other mapped values remain intact. Source rows are not changed, and their order is retained within each file.

For example, the base name `CS AMNESTY NEGA CARDS - L1` produces:

```text
CS AMNESTY NEGA CARDS - L1_Email.xlsx
CS AMNESTY NEGA CARDS - L1_Viber.xlsx
```

If a group has accounts but one channel has no matching accounts, that channel's file still contains the template headers with no account data. A group with no accounts produces no files. This behavior applies to qualified, blank, review, and other-sheet exports, including folder saves, individual downloads, and ZIP downloads.

The app shows account counts separately from Email and Viber row counts. Accounts with both contact details contribute a row to each channel, so adding the channel counts may exceed the number of included source accounts.

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

Accounts are removed only when **both `debtor_email` and `phone_1` are empty**. An account with either email or phone is kept in its matching channel file(s). Whitespace-only cells count as empty. This condition applies to qualified and blank groups, as well as other source sheets. The app displays the number of accounts removed separately from entirely empty rows.

Empty groups produce no file. Qualified rows outside the product/level rules and blank-status CARDS at levels other than 1 or 6 are put in separate, clearly identified review groups with Email/Viber files. They are never mixed into the other-products group. Every account with at least one contact field is assigned to exactly one group, then routed to its matching channel(s).

With the sample dated October 26, 2026, the three qualified group counts are **8**, **4,277**, and **131**. The four blank group counts are **3,066**, **2,155**, **121**, and **4,059**. This includes **13,817** source accounts after removing **541** without either contact field. The seven groups produce **14 files**, containing **12,414 Email rows** and **13,489 Viber rows**. No accounts need review in this sample.

## Edit filenames

Open **[export_config.py](export_config.py)** and edit the quoted values on the right:

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

Keep the keys on the left unchanged. Edit the base names in `OUTPUT_FILENAMES`; the app adds `_Email` or `_Viber` before `.xlsx` or `.csv`. You can edit these suffixes separately in `CHANNEL_SUFFIXES`. Final filenames must be unique and must not contain folder paths. Save the code, rerun the app, and prepare the files again. The app also shows this code under **Filename settings (export_config.py)**.

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

Header matching ignores capitalization, whitespace, and punctuation. `Product Type` is also accepted for `product`, matching your sample workbook. `OB` is also accepted as a source header for `OB (Masterlist)`. All eleven source and destination headers must be present; missing or ambiguous headers are reported before export. Header detection scans the first 50 rows; enter a row manually when headers appear later.

Your supplied template has 16 columns. The five columns without mapping rules (`Monthly`, `Terms`, `doclass`, `AREA`, and `NewEndo`) are retained and left blank.

## Output behavior

- Source order is retained within each output group. Hidden and filtered source rows are included.
- Rows with no values in any mapped field are skipped. Accounts with neither email nor phone are removed and counted separately. Duplicates are retained. Accounts with only email or only phone go to that channel; accounts with both go to both channels.
- IDs, account numbers, card numbers, and phones are written as Excel text. Existing text leading zeros and simple Excel zero masks such as `000000` are preserved. Digits already lost in the original Excel file cannot be recovered.
- Dates remain Excel dates and balances retain their original value type. No currency or date-string conversion is applied.
- Excel exports retain the destination header labels/order, content above the header, the first data row's cell styles, and other worksheets. Data below the selected header is cleared and replaced. Extra template columns are blank. Do not use the selected template sheet to store footer notes or formulas that must be kept.
- CSV exports retain the template's first-row headers, column order, and delimiter; dates use ISO text, and files use UTF-8 with BOM. CSV has no text column types: import identifier columns as text when opening it in Excel. Use `.xlsx` output when Excel must preserve leading zeros automatically.
- Formula source cells use values cached by Excel. A mapped formula without a saved result stops processing; recalculate and save the source in Excel before uploading again.
- Use ordinary `.xlsx` templates. Macros, merged data areas, and tables with totals rows are unsupported. Advanced Excel objects such as shapes may not survive an openpyxl round trip; see the [openpyxl documentation](https://openpyxl.readthedocs.io/en/stable/tutorial.html).
- Uploaded files and prepared outputs are held in memory. Clicking **Save files to folder** writes output copies to a new folder. Originals are never overwritten. Files are not emailed or sent to external services.

## Verification

```powershell
python -m unittest discover -s tests -v
```

The tests use generated sample data, including reordered headers, leading-zero identifiers, multiple sheets, stale template rows, Excel formulas/errors, all qualified/blank group boundaries, channel routing, retained contact headers, header-only channels, contact exclusions, folder saves, and the Streamlit upload-to-download flow.

The app uses Streamlit's [file uploader](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader) and [download button](https://docs.streamlit.io/develop/api-reference/widgets/st.download_button).
