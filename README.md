# FBT Manifest Split

Splits a FedEx `booking` export into the two FBT batch upload templates
(**US** / **Non-US**) by filling in **your actual template files** —
same formatting, same dropdown lists, same everything — with **live
Excel formulas**, saved back out as true Excel 97-2003 (`.xls`).

## Why this needs LibreOffice

A `.xls` (BIFF8) file can't be edited directly in Python while keeping
its formatting and dropdown (data validation) lists intact — no
maintained Python library reads+writes BIFF8 with full fidelity, and no
browser/JavaScript library can write a real formula into that format at
all (only the calculated value).

So this script does a three-step round trip:

1. Converts your template `.xls` → `.xlsx` with LibreOffice (lossless —
   keeps formatting and dropdowns).
2. Edits the `.xlsx` with `openpyxl` (keeps formatting and dropdowns,
   writes real formula text like `=AG2-(AE2+AA2+AB2)`).
3. Converts the edited `.xlsx` back → `.xls` with LibreOffice, which
   writes genuine BIFF8 formula records (not just cached numbers) and
   carries the dropdown lists through.

Verified end to end: the output contains real formula records (checked by
inspecting the BIFF8 structure), the Currency/Country/UOM dropdowns still
work, and the computed values match your reference files exactly.

One honest caveat: because the file passes through two format
conversions, a very small number of cosmetic style details (e.g. a
header's bold weight) can shift slightly. Column widths, fill colors,
and all data validation lists came through correctly in testing.

## Easy Windows use (recommended for non-technical users)

The repository also contains a simple Windows GUI (`fbt_gui.py`) and a GitHub Actions
workflow that builds a single `FBT-Manifest-Generator.exe`.

The generated Windows app:
- asks the user only for the booking `.xls` file;
- uses the bundled real FBT templates;
- creates the US and Non-US outputs automatically;
- processes the booking file locally — it does not upload it to GitHub;
- prefers Microsoft Excel already installed on the workstation, so no administrator
  install is normally required.

A company-managed Windows PC can still block an unsigned `.exe` through security policy.
If that happens, IT/security approval or code signing may be needed even though the
application itself does not require administrator rights.

To build it, open GitHub → **Actions** → **Build Windows FBT Generator** → **Run workflow**.
Download the `FBT-Manifest-Generator-Windows` artifact, extract it, and double-click the
EXE.

## Setup

**Local:**
```bash
git clone <this-repo-url>
cd fbt-manifest-split
pip install -r requirements.txt
```
You also need LibreOffice installed and on your `PATH` as `soffice`:
- Ubuntu/Debian: `sudo apt-get install -y libreoffice`
- macOS: `brew install --cask libreoffice`
- Windows: install LibreOffice normally; `soffice.exe` is added to PATH

**GitHub Actions (no local install needed):**
This repo includes `.github/workflows/generate.yml`. Push your booking
file and template files into the repo, then run the workflow from the
**Actions** tab with the file paths as inputs. It installs LibreOffice on
the runner automatically and uploads the two generated `.xls` files as a
downloadable artifact.

## Usage

```bash
python3 generate_fbt.py BOOKING_FILE.xls \
    --mode yes \
    --us-template FBT_US_NFEI_YES.xls \
    --nonus-template FBT_NON_US_NFEI_YES.xls \
    --outdir ./out
```

- `--mode yes` — NFEI YES rules (non-US rows leave Freight/Insurance/FOB
  blank; HS codes use the short 8-digit form for non-US).
- `--mode no` — NFEI NO rules (Freight/Insurance/FOB/Other charges are
  always calculated; HS codes use the full 10-digit form).
- `--us-template` / `--nonus-template` — your real FBT template files for
  this mode. These are edited in place (a copy is made; your originals
  are untouched) and drive the output's formatting and dropdowns.
- `--outdir` — where to write the two output files (default: current
  directory).

Only the sheet literally named `booking` (case-insensitive) is read from
the booking export — any other sheets are ignored.

## What gets generated

Two `.xls` files: `FBT_US_NFEI_<MODE>_<date>.xls` and
`FBT_NON_US_NFEI_<MODE>_<date>.xls`, built from your real templates:

1. **Recipient and Invoice Data** — old data rows cleared, one row written
   per booking row, with live formulas for `Other_charges`, `FOBValue`,
   `UNIT_VALUE1`, `UNIT_Weight1`. Dropdown lists on Currency,
   CountryofManufacture and UOM1 are preserved and stretched to cover
   every row you wrote.
2. **Importer_Info** — cleared back to header row only.
3. **MPS Dimension** — left as in your template (header only).
4. **StateCodes** — left untouched, exactly as in your template.

## Rule summary

| Field | Rule |
|---|---|
| SequenceNumber | 1, 2, 3… per output file |
| Recipient_ContactName / CompanyName | `cust_name` |
| Recipient_AddressLine1/2/3 | `cust_address`, word-wrapped at 35 chars/line; if the full address is over 105 chars, it goes entirely into Line1 |
| Recipient_Country | `destinationcountry` |
| Recipient_City / State / Postalcode | `cust_city`, `State`, `pin` |
| Recipient_PhoneNumber | `cust_mobile` |
| Reference_1 | `Ref. Number` |
| InvoiceNumber | `Invoice` |
| InvoiceDate | today, `ddmmyy` |
| TotalNoofPackage / TotalShipmentweight | 1 / 0.5 (constant) |
| Freight_charges / Insurance_charges | 13.5 / 0.5 — **except** NFEI YES + non-US, where these (and Other_charges, FOBValue, TotalGSTAmt, CarriageValue) are left blank |
| Other_charges | `=AG{r}-(AE{r}+AA{r}+AB{r})` (live formula) |
| FOBValue | `=(AG{r}-(AA{r}+AB{r}))/1.275` (live formula) |
| InvoiceValue | `Net Decleared amount` |
| Currency / CountryofManufacture | `US DOLLARS-USD` / `IN-INDIA` (constant) |
| Commodity | `Content` |
| HSCODE1 | Prescription Eyeglasses → `90049090` (NFEI YES+non-US) or `9004900090` (otherwise); Polarized Sunglasses → `90041000` or `9004100000` |
| StofOriginofgoods / DisOfOriginofgoods | `Gurugram` / `Haryana` (constant) |
| QUANTITY1 / QTY | `QTY` |
| UOM1 | `PIECE` (constant) |
| UNIT_VALUE1 | US: `=AE{r}/AN{r}`; **Non-US: `=AG{r}/AN{r}`** (live formula) |
| UNIT_Weight1 | `=W{r}/AN{r}` (live formula) |
| AdditionalShipment/Invoiceinfo | fixed FDA device-listing text (constant) |

Everything not listed above is left blank, per the rule books.


## Output cleanup

The generated **Recipient and Invoice Data** and **Importer_Info** sheets are physically
trimmed after generation: rows below the last populated field and columns to the right
of the last populated field are removed. The `StateCodes` support sheet is intentionally
not trimmed because its lookup data is required by the template's dropdowns.

## Important formula correction

For **Non-US**, `UNIT_VALUE1` is now:

`=AG{r}/AN{r}`

For **US**, the existing rule remains:

`=AE{r}/AN{r}`
