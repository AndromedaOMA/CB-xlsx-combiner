# Hotel workbook merger

The merger keeps both input workbooks as source tabs and creates a separate
merged view and reviewer checklist. Merged display cells may combine distinct
values, but the original rows and values remain available on the source tabs.

## Create a review workbook locally

Install the dependency and run the script with the two input files:

```powershell
python -m pip install -r requirements.txt
python main.py "Athens.xlsx" "Athens (1).xlsx" -o "Merged_Hotels_Review.xlsx"
```

The workbook contains:

- **Merged Hotels** — one row per automatically grouped hotel plus every
  unmerged record; orange rows need reviewer verification.
- **Match Review** — matched records and possible name/location variants with
  source row references and reviewer decision/notes columns.
- **Source tabs** — copies of the input worksheets, with matched source rows
  highlighted.
- **Unassigned Rows** — non-empty rows without an identifiable hotel name.
- **Read Me** — merge rules, color legend, and record counts.

Names are normalized for case, punctuation, and accents. Exact normalized names
are grouped unless available address/coordinate data conflict. A shared email
supports automatic grouping only when the names are highly similar and either
the meaningful name tokens agree or location data corroborate the match.
Conflicting locations and name-only near matches remain separate and are shown
as yellow review candidates; this avoids silently combining unrelated hotels
that share a chain contact or similar brand name.

## Use the static page

Open `index.html` locally or deploy this folder to Vercel with the **Other**
framework preset, no build command, and `.` as the output directory. The page
reads both selected workbooks in the browser and downloads the review workbook;
it has no upload endpoint or server-side processing. The spreadsheet library
loads from jsDelivr, so the page needs an internet connection to process files.
