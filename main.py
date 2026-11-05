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


GREEN = PatternFill("solid", fgColor="C6EFCE")
BLUE = PatternFill("solid", fgColor="DDEBF7")
RED = PatternFill("solid", fgColor="F4CCCC")
HEADER = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NAME_HEADERS = {"name", "hotel", "hotelname", "propertyname", "hotelnamemerged"}
HOTEL_COLUMNS = [
    "Hotel name", "Stars", "Google rating", "Email 1", "Email 2", "Email 3",
    "Email 4", "Email 5", "Address", "Phone number", "Website",
]
FIELD_ALIASES = {
    "Stars": {"stars", "star", "starcategory", "hotelstars", "starclassification",
              "hotelclass", "hotelcategory", "category", "rating", "hotelrating", "googlestars",
              "stelebooking", "bookingstars"},
    "Google rating": {"googlerating", "googlescore", "googlereviewrating", "googlereviews", "googlereviewscore"},
    "Address": {"address", "hoteladdress", "streetaddress"},
    "Phone number": {"phone", "phonenumber", "telephone", "tel", "hotelphone"},
    "Website": {"website", "web", "url", "hotelwebsite", "websiteurl"},
}
EMAIL_PATTERN = re.compile(r"[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+", re.I)
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


def field_kind(header: str) -> str:
    key = normalized(header)
    if key in NAME_HEADERS:
        return "Hotel name"
    if "email" in key:
        return "Email"
    return next((name for name, aliases in FIELD_ALIASES.items() if key in aliases), key)


def star_number(value: Any) -> int | None:
    text = cell_text(value).lower()
    if re.fullmatch(r"[\u2605\u2606*]{1,5}", text):
        return len(text)
    match = re.fullmatch(r"([1-5])(?:[.,]0)?\s*(?:\*|stars?|stele|\u2605)?", text)
    return int(match[1]) if match else None


def category_number(value: Any) -> int | None:
    match = re.match(r"^\s*([1-5])\s*(?:\*|stars?\b|stele\b|\u2605)", cell_text(value), re.I)
    return int(match[1]) if match else None


def record_emails(fields: dict[str, Any]) -> list[str]:
    emails: list[str] = []
    seen: set[str] = set()
    for header, value in fields.items():
        if field_kind(header) != "Email" and not re.match(r"^(field|column)_", header):
            continue
        for match in EMAIL_PATTERN.finditer(cell_text(value)):
            email = match[0].strip(".")
            if email.lower() not in seen:
                seen.add(email.lower())
                emails.append(email)
    return emails


def email_priority(email: str) -> int:
    local = email.split("@", 1)[0].lower()
    if re.search(r"group|sales|reservation|reservierung|reserver|booking|meeting|event|banquet|convention|(^|[._-])rsv", local):
        return 0
    if re.search(r"no[._-]?reply|newsletter|unsubscribe", local):
        return 2
    return 1


def hotel_values(group_records: list[dict[str, Any]]) -> tuple[list[Any], list[str]]:
    """Combine aliases across every source; leave contradictory categories for review."""
    fields: dict[str, list[Any]] = defaultdict(list)
    emails: list[str] = []
    seen_emails: set[str] = set()
    for record in group_records:
        for header, value in record["fields"].items():
            if value is None or cell_text(value) == "":
                continue
            kind = field_kind(header)
            if kind not in {"Hotel name", "Email"}:
                fields[kind].append(value.strip() if isinstance(value, str) else value)
        if record.get("category") and not any(
            star_number(value) for header, value in record["fields"].items() if field_kind(header) == "Stars"
        ):
            fields["Stars"].append(record["category"])
        for email in record_emails(record["fields"]):
            if email.lower() not in seen_emails:
                emails.append(email)
                seen_emails.add(email.lower())
    emails.sort(key=email_priority)  # Stable: equal-priority contacts retain source order.
    notes: list[str] = []
    stars = sorted({number for value in fields["Stars"] if (number := star_number(value))})
    category = f"{stars[0]}*" if len(stars) == 1 else "Conflicting / review" if stars else "Unrated / review"
    if len(stars) > 1:
        notes.append("Conflicting stars: " + ", ".join(f"{star}*" for star in stars))
    result: list[Any] = [cell_text(group_records[0]["name"]), category]
    for kind in ["Google rating", "Address", "Phone number", "Website"]:
        values = unique_values(fields[kind])
        if kind == "Google rating":
            values = unique_values([numeric_rating(value) for value in values])
        if len(values) > 1:
            notes.append(f"Conflicting {kind.lower()}: " + " | ".join(map(str, values)))
        result.append(values[0] if len(values) == 1 else " | ".join(map(str, values)) if values else None)
    if len(emails) > 5:
        notes.append(f"{len(emails) - 5} additional email(s) retained on Source tabs")
    return result[:3] + (emails[:5] + [None] * max(0, 5 - len(emails))) + result[3:], notes


def numeric_rating(value: Any) -> Any:
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return value


def is_header_row(values: tuple[Any, ...] | list[Any]) -> bool:
    kinds = {field_kind(cell_text(value)) for value in values if value is not None}
    return "Hotel name" in kinds and bool(kinds & {"Stars", "Email", "Google rating", "Address", "Phone number", "Website"})


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


def workbook_rows(path: Path, source_number: int = 1) -> list[dict[str, Any]]:
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

        category = category_number(worksheet.title)
        for row_index, values in enumerate(rows):
            if row_index == header_index or is_header_row(values):
                continue
            padded = tuple(values) + (None,) * (len(headers) - len(values))
            fields = {}
            for column_index, (header, value) in enumerate(zip(headers, padded)):
                key = header or f"column_{column_index + 1}"
                while key in fields:
                    key += " "
                fields[key] = value
            name = next(
                (
                    value
                    for header, value in fields.items()
                    if normalized(header) in NAME_HEADERS
                ),
                None,
            )
            nonempty = [value for value in padded if cell_text(value)]
            section = category_number(name) if len(nonempty) == 1 else None
            if section:
                category = section
                continue
            if not name or (header_index == -1 and len(nonempty) < 2):
                continue
            emails = {email.lower() for email in record_emails(fields)}
            lat = next(
                (value for header, value in fields.items() if normalized(header) in {"latitude", "lat"}),
                None,
            )
            lon = next(
                (value for header, value in fields.items() if normalized(header) in {"longitude", "lon", "lng"}),
                None,
            )
            address = next(
                (value for header, value in fields.items() if field_kind(header) == "Address"),
                None,
            )
            records.append(
                {
                    "source": path.name,
                    "source_number": source_number,
                    "sheet": worksheet.title,
                    "row": row_index + 1,
                    "fields": fields,
                    "name": name,
                    "emails": {email for email in emails if email},
                    "address": address,
                    "coordinate": coordinate((lat, lon)),
                    "category": category,
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
        left.get("source_number", left["source"]) != right.get("source_number", right["source"])
        and similarity >= POTENTIAL_NAME_SIMILARITY
        and meaningful_name_tokens(left["name"])
        and meaningful_name_tokens(left["name"]) & meaningful_name_tokens(right["name"])
    ):
        return None, f"Similar hotel name ({similarity:.0%}); verify location"
    return None, None


def reference(record: dict[str, Any]) -> str:
    return f'Input {record["source_number"]} / {record["source"]} / {record["sheet"]} / row {record["row"]}'


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
            cell.fill = BLUE
    for row_number in candidate_rows - duplicate_rows:
        for cell in destination[row_number]:
            cell.fill = RED
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
    records = [record for number, path in enumerate(input_paths, 1) for record in workbook_rows(path, number)]
    if not records:
        raise ValueError("No hotel rows were found in the input workbooks.")

    links: list[dict[str, Any]] = []
    disjoint_set = DisjointSet(len(records))
    for left_index, left in enumerate(records):
        for right_index in range(left_index + 1, len(records)):
            right = records[right_index]
            auto_reason, candidate_reason = pair_evidence(left, right)
            if auto_reason and disjoint_set.find(left_index) != disjoint_set.find(right_index):
                left_group = [index for index in range(len(records)) if disjoint_set.find(index) == disjoint_set.find(left_index)]
                right_group = [index for index in range(len(records)) if disjoint_set.find(index) == disjoint_set.find(right_index)]
                if any(
                    (pair_evidence(records[a], records[b])[1] or "").startswith("Location fields disagree")
                    for a in left_group for b in right_group
                ):
                    auto_reason = None
                    candidate_reason = "Location fields disagree within proposed group; keep separate for review"
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
    def group_order(group: list[int]) -> tuple[int, str]:
        values, _ = hotel_values([records[index] for index in group])
        star = star_number(values[1])
        return ({3: 0, 4: 1, 5: 2}.get(star, 3), normalized(values[0]))

    ordered_groups = sorted(groups.values(), key=group_order)

    duplicate_rows: dict[tuple[int, str], set[int]] = defaultdict(set)
    candidate_rows: dict[tuple[int, str], set[int]] = defaultdict(set)
    for link in links:
        for index in (link["left"], link["right"]):
            record = records[index]
            key = (record["source_number"], record["sheet"])
            if link["status"] == "MERGED - REVIEW":
                duplicate_rows[key].add(record["row"])
            else:
                candidate_rows[key].add(record["row"])

    output = Workbook()
    output.remove(output.active)
    used_titles: set[str] = set()
    input_workbooks = [load_workbook(path, data_only=False) for path in input_paths]
    for input_index, (path, source_book) in enumerate(zip(input_paths, input_workbooks), start=1):
        for source_worksheet in source_book.worksheets:
            title = safe_sheet_title(
                f"Source {input_index} - {path.stem} - {source_worksheet.title}", used_titles
            )
            key = (input_index, source_worksheet.title)
            copy_source_sheet(
                source_worksheet,
                output,
                title,
                duplicate_rows[key],
                candidate_rows[key],
            )

    extra_headers: dict[str, str] = {}
    for record in records:
        for field, value in record["fields"].items():
            kind = field_kind(field)
            if re.match(r"^(field|column)_", field) and record_emails({field: value}):
                remainder = EMAIL_PATTERN.sub("", cell_text(value)).strip(" ;,|<>.\t\r\n")
                if not remainder:
                    continue  # Unlabelled contacts already belong in the five email columns.
            if kind not in {"Hotel name", "Email", *FIELD_ALIASES} and cell_text(value):
                extra_headers.setdefault(kind, field)

    merged_sheet = output.create_sheet("Merged Hotels", 0)
    merged_sheet.append(HOTEL_COLUMNS + list(extra_headers.values()) + [
        "Row type", "Source record count", "Source references", "Review notes",
    ])

    def append_hotel(hotels: list[dict[str, Any]], row_type: str, fill: Any = None) -> None:
        values, notes = hotel_values(hotels)
        for kind, label in extra_headers.items():
            alternatives = unique_values([
                value for hotel in hotels for header, value in hotel["fields"].items()
                if field_kind(header) == kind and cell_text(value)
            ])
            if len(alternatives) > 1:
                notes.append(f"Conflicting {label}: " + " | ".join(map(str, alternatives)))
            values.append(alternatives[0] if len(alternatives) == 1 else " | ".join(map(str, alternatives)) if alternatives else None)
        if row_type == "MERGED - REVIEW":
            notes.insert(0, "Check merged hotel and contact details before use")
        merged_sheet.append(values + [row_type, len(hotels), " | ".join(reference(hotel) for hotel in hotels), "; ".join(notes)])
        for cell in merged_sheet[merged_sheet.max_row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if fill:
                cell.fill = fill
        merged_sheet.cell(merged_sheet.max_row, 3).number_format = "0.0"
        merged_sheet.cell(merged_sheet.max_row, 10).number_format = "@"

    for group in ordered_groups:
        group_records = [records[index] for index in group]
        duplicate = len(group_records) > 1
        append_hotel(group_records, "MERGED - REVIEW" if duplicate else "SINGLE RECORD", GREEN if duplicate else None)
        if duplicate:
            for record in group_records:
                append_hotel([record], "SOURCE DUPLICATE", BLUE)
    style_header(merged_sheet)

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
        fill = GREEN if link["status"] == "MERGED - REVIEW" else RED
        for cell in review_sheet[review_sheet.max_row]:
            cell.fill = fill
    style_header(review_sheet)

    unassigned_sheet = output.create_sheet("Unassigned Rows")
    unassigned_sheet.append(["Source", "Sheet", "Row", "Raw non-empty values"])
    assigned_positions = {
        (record["source_number"], record["sheet"], record["row"]) for record in records
    }
    for input_index, (path, source_book) in enumerate(zip(input_paths, input_workbooks), 1):
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
                if (input_index, source_worksheet.title, row_number) in assigned_positions:
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

    instructions = output.create_sheet("Read Me")
    instructions.append(["Hotel workbook merge - reviewer guide"])
    instructions.append(["Both original workbooks are copied to the Source tabs; source values are not removed."])
    instructions.append(["Merged Hotels is sorted by 3*, 4*, 5*, then other/unrated/conflicting categories; hotels are alphabetical within each category."])
    instructions.append(["Green rows combine a hotel's data and require manual checking. Blue original duplicate records sit directly below each green row."])
    instructions.append(["Red rows in Match Review are possible variants only. They are NOT automatically merged."])
    instructions.append(["Five email columns show distinct contacts, prioritizing Groups, Sales, Reservations and Events; ties keep source order. Extra emails remain on Source tabs."])
    instructions.append(["Location matching uses identical normalized addresses or coordinates within 30 m plus a similar name."])
    instructions.append(["Hotel stars and Google review ratings are separate. Conflicting star categories remain unclassified for review; conflicting values are listed in Review notes."])
    instructions.append(["Blank-name rows and category/header rows are retained in Source tabs; non-header unassigned rows are listed separately."])
    instructions.append(["Contacts are taken from your files; their validity and whether they are verified must be checked manually. Filter Stars and Row type to select hotels."])
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
