"""Merge two hotel workbooks into an auditable, non-destructive review workbook."""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from collections import defaultdict
from copy import copy
from difflib import SequenceMatcher
from math import asin, cos, radians, sin, sqrt
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ORANGE = PatternFill("solid", fgColor="FCE4D6")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", fgColor="F4CCCC")
HEADER = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NAME_HEADERS = {"name", "hotel", "hotelname", "propertyname"}
RATING_HEADERS = {"rating", "stars", "hotelrating", "googlestars"}
EMAIL_HEADERS = {"email", "emailaddress", "emailadress"}
COORDINATE_RADIUS_METERS = 30
AUTO_NAME_SIMILARITY = 0.86
POTENTIAL_NAME_SIMILARITY = 0.88
GENERIC_NAME_TOKENS = {
    "a", "an", "and", "athens", "by", "city", "hotel", "hotels", "in", "of",
    "the",
}


def normalized(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]", "", text)


def tokens(value: Any) -> set[str]:
    if value is None:
        return set()
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    return set(re.findall(r"[a-z0-9]+", text))


def name_similarity(left: Any, right: Any) -> float:
    left_tokens, right_tokens = tokens(left), tokens(right)
    union = left_tokens | right_tokens
    token_score = len(left_tokens & right_tokens) / len(union) if union else 0
    sequence_score = SequenceMatcher(
        None, normalized(left), normalized(right)
    ).ratio()
    return max(token_score, sequence_score)


def meaningful_name_tokens(value: Any) -> set[str]:
    return tokens(value) - GENERIC_NAME_TOKENS


def coordinate(value: Any) -> tuple[float, float] | None:
    try:
        latitude, longitude = float(value[0]), float(value[1])
    except (TypeError, ValueError, IndexError):
        return None
    if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
        return None
    return latitude, longitude


def distance_meters(
    left: tuple[float, float], right: tuple[float, float]
) -> float:
    lat1, lon1 = map(radians, left)
    lat2, lon2 = map(radians, right)
    haversine = (
        sin((lat2 - lat1) / 2) ** 2
        + cos(lat1) * cos(lat2) * sin((lon2 - lon1) / 2) ** 2
    )
    return 6_371_000 * 2 * asin(sqrt(haversine))


def cell_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def is_header_row(values: tuple[Any, ...] | list[Any]) -> bool:
    first = normalized(values[0]) if values else ""
    rest = {normalized(value) for value in values[1:] if value is not None}
    return first in NAME_HEADERS and bool(rest & (RATING_HEADERS | EMAIL_HEADERS))


def find_header_row(rows: list[tuple[Any, ...]]) -> int | None:
    for row_index, values in enumerate(rows):
        if is_header_row(values):
            return row_index
    return None


def unique_values(values: list[Any]) -> list[Any]:
    result: list[Any] = []
    seen: set[tuple[str, str]] = set()
    for value in values:
        if value is None or value == "":
            continue
        key = (type(value).__name__, str(value))
        if key not in seen:
            result.append(value)
            seen.add(key)
    return result


class DisjointSet:
    def __init__(self, size: int) -> None:
        self.parents = list(range(size))

    def find(self, item: int) -> int:
        if self.parents[item] != item:
            self.parents[item] = self.find(self.parents[item])
        return self.parents[item]

    def union(self, left: int, right: int) -> None:
        left_root, right_root = self.find(left), self.find(right)
        if left_root != right_root:
            self.parents[right_root] = left_root


def workbook_rows(path: Path) -> list[dict[str, Any]]:
    source_book = load_workbook(path, data_only=False)
    records: list[dict[str, Any]] = []
    for worksheet in source_book.worksheets:
        rows = [tuple(row) for row in worksheet.iter_rows(values_only=True)]
        header_index = find_header_row(rows)
        max_columns = max((len(row) for row in rows), default=0)
        if header_index is None:
            header_index = -1
            headers = ["name", "rating", "email"] + [
                f"field_{index}" for index in range(4, max_columns + 1)
            ]
        else:
            headers = [cell_text(value) for value in rows[header_index]]
            while len(headers) < max_columns:
                headers.append(f"column_{len(headers) + 1}")

        for row_index, values in enumerate(rows):
            if row_index == header_index or is_header_row(values):
                continue
            padded = tuple(values) + (None,) * (len(headers) - len(values))
            fields = {
                header or f"column_{column_index + 1}": value
                for column_index, (header, value) in enumerate(zip(headers, padded))
            }
            name = next(
                (
                    value
                    for header, value in fields.items()
                    if normalized(header) in NAME_HEADERS
                ),
                None,
            )
            if not name or not any(value is not None and value != "" for value in padded[1:]):
                continue
            emails = {
                normalized(value)
                for header, value in fields.items()
                if "email" in normalized(header) and value and "@" in str(value)
            }
            lat = next(
                (value for header, value in fields.items() if normalized(header) in {"latitude", "lat"}),
                None,
            )
            lon = next(
                (value for header, value in fields.items() if normalized(header) in {"longitude", "lon", "lng"}),
                None,
            )
            address = next(
                (value for header, value in fields.items() if normalized(header) in {"address", "hoteladdress"}),
                None,
            )
            records.append(
                {
                    "source": path.name,
                    "sheet": worksheet.title,
                    "row": row_index + 1,
                    "fields": fields,
                    "name": name,
                    "emails": {email for email in emails if email},
                    "address": address,
                    "coordinate": coordinate((lat, lon)),
                }
            )
    source_book.close()
    return records


def pair_evidence(left: dict[str, Any], right: dict[str, Any]) -> tuple[str | None, str | None]:
    similarity = name_similarity(left["name"], right["name"])
    same_name = normalized(left["name"]) == normalized(right["name"])
    shared_email = bool(left["emails"] & right["emails"])
    address_match = bool(
        left["address"]
        and right["address"]
        and normalized(left["address"]) == normalized(right["address"])
    )
    location_distance = (
        distance_meters(left["coordinate"], right["coordinate"])
        if left["coordinate"] and right["coordinate"]
        else None
    )
    nearby = location_distance is not None and location_distance <= COORDINATE_RADIUS_METERS
    address_differs = bool(left["address"] and right["address"] and not address_match)
    coordinates_differ = location_distance is not None and not nearby
    location_conflict = (address_differs or coordinates_differ) and not (address_match or nearby)
    if location_conflict:
        if same_name or shared_email or similarity >= POTENTIAL_NAME_SIMILARITY:
            return None, "Location fields disagree; keep separate pending reviewer confirmation"
        return None, None
    if same_name:
        return "Exact normalized hotel name", None
    if shared_email and similarity >= AUTO_NAME_SIMILARITY:
        location_confirmed = address_match or nearby
        same_core_name = meaningful_name_tokens(left["name"]) == meaningful_name_tokens(right["name"])
        if location_confirmed or same_core_name:
            return "Shared email, similar name, and no conflicting location", None
        return None, "Shared email/name similarity without location confirmation"
    if address_match and similarity >= 0.70:
        return "Same address and similar name", None
    if nearby and similarity >= AUTO_NAME_SIMILARITY:
        return "Coordinates within 30 m and highly similar name", None
    if shared_email:
        return None, "Shared email; confirm property identity"
    if address_match:
        return None, "Same address; confirm property identity"
    if nearby:
        return None, "Coordinates within 30 m; confirm property identity"
    if (
        left["source"] != right["source"]
        and similarity >= POTENTIAL_NAME_SIMILARITY
        and meaningful_name_tokens(left["name"])
        and meaningful_name_tokens(left["name"]) & meaningful_name_tokens(right["name"])
    ):
        return None, f"Similar hotel name ({similarity:.0%}); verify location"
    return None, None


def reference(record: dict[str, Any]) -> str:
    return f'{record["source"]} / {record["sheet"]} / row {record["row"]}'


def style_header(worksheet: Any) -> None:
    for cell in worksheet[1]:
        cell.fill = HEADER
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def copy_source_sheet(
    source_worksheet: Any,
    destination_workbook: Workbook,
    title: str,
    duplicate_rows: set[int],
    candidate_rows: set[int],
) -> Any:
    destination = destination_workbook.create_sheet(title=title)
    for row in source_worksheet.iter_rows():
        for source_cell in row:
            target = destination.cell(
                row=source_cell.row, column=source_cell.column, value=source_cell.value
            )
            if source_cell.has_style:
                target.font = copy(source_cell.font)
                target.fill = copy(source_cell.fill)
                target.border = copy(source_cell.border)
                target.alignment = copy(source_cell.alignment)
                target.protection = copy(source_cell.protection)
            if source_cell.number_format:
                target.number_format = source_cell.number_format
            if source_cell.hyperlink:
                target._hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target.comment = copy(source_cell.comment)
    for merged_range in source_worksheet.merged_cells.ranges:
        destination.merge_cells(str(merged_range))
    for key, dimension in source_worksheet.column_dimensions.items():
        destination.column_dimensions[key].width = dimension.width
        destination.column_dimensions[key].hidden = dimension.hidden
    for key, dimension in source_worksheet.row_dimensions.items():
        destination.row_dimensions[key].height = dimension.height
        destination.row_dimensions[key].hidden = dimension.hidden
    destination.freeze_panes = source_worksheet.freeze_panes
    destination.sheet_view.showGridLines = source_worksheet.sheet_view.showGridLines
    for row_number in duplicate_rows:
        for cell in destination[row_number]:
            cell.fill = ORANGE
    for row_number in candidate_rows - duplicate_rows:
        for cell in destination[row_number]:
            cell.fill = YELLOW
    return destination


def safe_sheet_title(title: str, used: set[str]) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "-", title)[:31] or "Source"
    candidate = cleaned
    suffix = 2
    while candidate in used:
        ending = f" ({suffix})"
        candidate = f"{cleaned[:31 - len(ending)]}{ending}"
        suffix += 1
    used.add(candidate)
    return candidate


def merge_workbooks(input_paths: list[Path], output_path: Path) -> dict[str, int]:
    if len(input_paths) != 2:
        raise ValueError("Provide exactly two input workbooks.")
    records = [record for path in input_paths for record in workbook_rows(path)]
    if not records:
        raise ValueError("No hotel rows were found in the input workbooks.")

    links: list[dict[str, Any]] = []
    disjoint_set = DisjointSet(len(records))
    for left_index, left in enumerate(records):
        for right_index in range(left_index + 1, len(records)):
            right = records[right_index]
            auto_reason, candidate_reason = pair_evidence(left, right)
            if auto_reason:
                disjoint_set.union(left_index, right_index)
                links.append(
                    {
                        "left": left_index,
                        "right": right_index,
                        "status": "MERGED - REVIEW",
                        "reason": auto_reason,
                        "similarity": name_similarity(left["name"], right["name"]),
                    }
                )
            elif candidate_reason:
                links.append(
                    {
                        "left": left_index,
                        "right": right_index,
                        "status": "POTENTIAL MATCH - REVIEW",
                        "reason": candidate_reason,
                        "similarity": name_similarity(left["name"], right["name"]),
                    }
                )

    groups: dict[int, list[int]] = defaultdict(list)
    for index in range(len(records)):
        groups[disjoint_set.find(index)].append(index)
    ordered_groups = sorted(groups.values(), key=lambda group: min(group))

    duplicate_rows: dict[tuple[str, str], set[int]] = defaultdict(set)
    candidate_rows: dict[tuple[str, str], set[int]] = defaultdict(set)
    for link in links:
        for index in (link["left"], link["right"]):
            record = records[index]
            key = (record["source"], record["sheet"])
            if link["status"] == "MERGED - REVIEW":
                duplicate_rows[key].add(record["row"])
            else:
                candidate_rows[key].add(record["row"])

    output = Workbook()
    output.remove(output.active)
    used_titles: set[str] = set()
    input_workbooks = [load_workbook(path, data_only=False) for path in input_paths]
    source_titles: dict[tuple[str, str], str] = {}
    for input_index, (path, source_book) in enumerate(zip(input_paths, input_workbooks), start=1):
        for source_worksheet in source_book.worksheets:
            title = safe_sheet_title(
                f"Source {input_index} - {path.stem} - {source_worksheet.title}", used_titles
            )
            key = (path.name, source_worksheet.title)
            source_titles[key] = title
            copy_source_sheet(
                source_worksheet,
                output,
                title,
                duplicate_rows[key],
                candidate_rows[key],
            )

    source_headers: list[tuple[str, str, str]] = []
    for record in records:
        for field in record["fields"]:
            column = (record["source"], record["sheet"], field)
            if column not in source_headers:
                source_headers.append(column)

    merged_sheet = output.create_sheet("Merged Hotels", 0)
    fixed_headers = [
        "Hotel name (merged)",
        "Merge status",
        "Source record count",
        "Source workbooks",
        "Match evidence",
        "Source references",
    ]
    field_headers = [
        f"{source_titles[(source, sheet)]} / {field}"
        for source, sheet, field in source_headers
    ]
    merged_sheet.append(fixed_headers + field_headers)
    conflicts: list[tuple[int, int]] = []
    for group in ordered_groups:
        group_records = [records[index] for index in group]
        names = unique_values([record["name"] for record in group_records])
        evidence = unique_values(
            [
                link["reason"]
                for link in links
                if link["status"] == "MERGED - REVIEW"
                and link["left"] in group
                and link["right"] in group
            ]
        )
        status = "MERGED - REVIEW" if len(group) > 1 else "SINGLE RECORD"
        values: list[Any] = [
            " | ".join(str(name) for name in names),
            status,
            len(group_records),
            " | ".join(unique_values([record["source"] for record in group_records])),
            " | ".join(str(value) for value in evidence),
            " | ".join(reference(record) for record in group_records),
        ]
        for column_index, (source, sheet, field) in enumerate(source_headers, start=7):
            field_values = unique_values(
                [
                    record["fields"].get(field)
                    for record in group_records
                    if record["source"] == source and record["sheet"] == sheet
                ]
            )
            if len(field_values) > 1:
                conflicts.append((merged_sheet.max_row + 1, column_index))
            values.append(
                field_values[0]
                if len(field_values) == 1
                else " | ".join(str(value) for value in field_values)
                if field_values
                else None
            )
        merged_sheet.append(values)
        if len(group_records) > 1:
            for cell in merged_sheet[merged_sheet.max_row]:
                cell.fill = ORANGE
    style_header(merged_sheet)
    for row_number, column_number in conflicts:
        merged_sheet.cell(row_number, column_number).fill = RED

    review_sheet = output.create_sheet("Match Review")
    review_sheet.append(
        [
            "Review status",
            "Evidence / reason",
            "Name similarity",
            "Hotel name (input A)",
            "Source reference (input A)",
            "Hotel name (input B)",
            "Source reference (input B)",
            "Reviewer decision",
            "Reviewer notes",
        ]
    )
    for link in links:
        left, right = records[link["left"]], records[link["right"]]
        review_sheet.append(
            [
                link["status"],
                link["reason"],
                round(link["similarity"], 3),
                left["name"],
                reference(left),
                right["name"],
                reference(right),
                None,
                None,
            ]
        )
        fill = ORANGE if link["status"] == "MERGED - REVIEW" else YELLOW
        for cell in review_sheet[review_sheet.max_row]:
            cell.fill = fill
    style_header(review_sheet)

    unassigned_sheet = output.create_sheet("Unassigned Rows")
    unassigned_sheet.append(["Source", "Sheet", "Row", "Raw non-empty values"])
    assigned_positions = {
        (record["source"], record["sheet"], record["row"]) for record in records
    }
    for path in input_paths:
        source_book = next(
            book for candidate_path, book in zip(input_paths, input_workbooks) if candidate_path == path
        )
        for source_worksheet in source_book.worksheets:
            header_index = find_header_row(
                [tuple(row) for row in source_worksheet.iter_rows(min_row=1, max_row=min(source_worksheet.max_row or 1, 1000), values_only=True)]
            )
            for row_number, row in enumerate(source_worksheet.iter_rows(values_only=True), start=1):
                values = list(row)
                if not any(value is not None and value != "" for value in values):
                    continue
                if header_index == row_number - 1 or is_header_row(values):
                    continue
                if (path.name, source_worksheet.title, row_number) in assigned_positions:
                    continue
                unassigned_sheet.append(
                    [
                        path.name,
                        source_worksheet.title,
                        row_number,
                        json.dumps(
                            {str(index + 1): value for index, value in enumerate(values) if value is not None},
                            ensure_ascii=False,
                            default=str,
                        ),
                    ]
                )
    style_header(unassigned_sheet)

    instructions = output.create_sheet("Read Me", 0)
    instructions.append(["Hotel workbook merge - reviewer guide"])
    instructions.append(["Both original workbooks are copied to the Source tabs; source values are not removed."])
    instructions.append(["Merged Hotels lists exact-name/strong-evidence groups plus all single-source records."])
    instructions.append(["Orange rows are merged groups or exact/high-confidence duplicate links; review their source references."])
    instructions.append(["Yellow rows in Match Review are possible variants only. They are NOT automatically merged."])
    instructions.append(["Red cells in Merged Hotels indicate differing values within that source field; all raw values remain on Source tabs."])
    instructions.append(["Location matching uses identical normalized addresses or coordinates within 30 m plus a similar name."])
    instructions.append(["This pair of sample files has no location fields in its contact-list workbook, so name-only variants remain review candidates."])
    instructions.append(["Blank-name rows and category/header rows are retained in Source tabs; non-header unassigned rows are listed separately."])
    instructions.append(["Merged rows combine distinct values only for display. Inspect the source references/tabs before accepting any merge."])
    instructions.append(["Input records processed", len(records)])
    instructions.append(["Merged groups", sum(len(group) > 1 for group in ordered_groups)])
    instructions.append(["Single records", sum(len(group) == 1 for group in ordered_groups)])
    instructions.append(["Potential matches", sum(link["status"] == "POTENTIAL MATCH - REVIEW" for link in links)])
    instructions.column_dimensions["A"].width = 125
    for row in instructions.iter_rows():
        instructions.row_dimensions[row[0].row].height = 30
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
    instructions["A1"].font = Font(bold=True, size=16, color="1F4E78")
    instructions.freeze_panes = "A2"

    for worksheet in output.worksheets:
        if worksheet.title not in {"Read Me"}:
            for column_cells in worksheet.columns:
                letter = get_column_letter(column_cells[0].column)
                max_length = max(
                    (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                    default=10,
                )
                worksheet.column_dimensions[letter].width = min(max(max_length + 2, 12), 55)
            worksheet.sheet_view.showGridLines = False

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output.save(output_path)
    output.close()
    for source_book in input_workbooks:
        source_book.close()
    return {
        "records": len(records),
        "merged_groups": sum(len(group) > 1 for group in ordered_groups),
        "single_records": sum(len(group) == 1 for group in ordered_groups),
        "potential_matches": sum(link["status"] == "POTENTIAL MATCH - REVIEW" for link in links),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "inputs",
        nargs="*",
        type=Path,
        help="exactly two .xlsx files (defaults to the first two in this folder)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("Merged_Hotels_Review.xlsx"),
        help="output review workbook",
    )
    args = parser.parse_args()
    input_paths = args.inputs or sorted(
        path for path in Path.cwd().glob("*.xlsx") if path.resolve() != args.output.resolve()
    )[:2]
    if len(input_paths) != 2:
        parser.error("Provide exactly two input .xlsx files, or place two .xlsx files in this folder.")
    for path in input_paths:
        if not path.is_file():
            parser.error(f"Input file not found: {path}")
        if path.suffix.lower() != ".xlsx":
            parser.error(f"Input must be an .xlsx file: {path}")
    try:
        summary = merge_workbooks(input_paths, args.output)
    except (OSError, ValueError) as error:
        parser.error(str(error))
    print(f"Created {args.output.resolve()}")
    print(
        f"Processed {summary['records']} hotel records; "
        f"{summary['merged_groups']} merged groups; "
        f"{summary['single_records']} single records; "
        f"{summary['potential_matches']} potential matches for review."
    )


if __name__ == "__main__":
    main()
