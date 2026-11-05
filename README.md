# Hotel workbook merger

The merger combines two hotel lists into a list Operations can check and use
to select hotels. The browser interface and static Vercel deployment are unchanged.

## How to explain the list to Operations

1. **Select a category.** Hotels appear in 3*, 4*, then 5* order, alphabetically
   within each category. Other, unrated and conflicting categories follow, so
   no source records are discarded. Filter **Stars** for the client's budget.
2. **Read one hotel row.** The first columns are hotel name, stars, Google
   rating, five separate emails, address, phone number and website. Matching
   fields from both files are combined into these columns. Other populated
   source fields follow once each, rather than repeating columns for each file.
3. **Check green rows.** A green row combines duplicate records. Its original
   records appear directly below in blue, in the same column layout. All raw
   values also remain on the Source tabs, including emails beyond the first five.
4. **Choose contacts.** Emails are split, deduplicated without regard to case,
   and prioritized for Groups, Sales, Reservations, Meetings and Events. Equal
   priorities keep their original file/row/column order. Personal and general
   contacts follow, then no-reply/newsletter addresses. This does not verify an
   address or identify which contacts previously supplied an offer.
5. **Review before using.** Check **Review notes** for conflicting fields or
   extra emails. Differing star classifications stay **Conflicting / review**.
   Filter **Row type** to `MERGED - REVIEW` and `SINGLE RECORD` when selecting
   hotels after checking them; blue rows are supporting records.

Hotel stars (including `google_stars` and `Stele_Booking`) are separate from the
Google review score (`google_rating`). A category heading such as `3 Stele Athens`
supplies stars only when a row has no usable star classification. Unknown stars
are never inferred from the Google review score.

## Create a review workbook locally

Install the dependency and run the script with the two input files:

```powershell
python -m pip install -r requirements.txt
python main.py "Athens.xlsx" "Athens (1).xlsx" -o "Merged_Hotels_Review.xlsx"
```

The workbook contains:

- **Merged Hotels** — sorted hotel rows, with green combined rows followed
  immediately by their blue original duplicate records.
- **Match Review** — matched records and possible name/location variants with
  source row references and reviewer decision/notes columns.
- **Source tabs** — copies of the input worksheets, with matched source rows
  highlighted in blue and possible matches highlighted in red.
- **Unassigned Rows** — non-empty rows without an identifiable hotel name.
- **Read Me** — merge rules, color legend, and record counts.

Names are normalized for case, punctuation, and accents. Exact normalized names
are grouped unless available address/coordinate data conflict. A shared email
supports automatic grouping only when the names are highly similar and either
the meaningful name tokens agree or location data corroborate the match.
Conflicting locations and name-only near matches remain separate and are shown
as red review candidates; this avoids silently combining unrelated hotels
that share a chain contact or similar brand name.
Conflicting locations cannot be joined indirectly through a third record that
has no location. Input numbers distinguish sources even when filenames match.

## Regression checks

With Python, the requirements above and Node.js installed:

```powershell
python -m unittest discover -s tests -v
```

Checks cover both merge engines, email selection, alphabetical category order,
duplicate placement and colors, source preservation, ambiguous categories and
location conflicts, including the supplied Athens files and Amsterdam reference.

## Use the static page

Open `index.html` locally or deploy this folder to Vercel with the **Other**
framework preset, no build command, and `.` as the output directory. The page
reads both selected workbooks in the browser and downloads the review workbook;
it has no upload endpoint or server-side processing. The spreadsheet library
loads from jsDelivr, so the page needs an internet connection to process files.
