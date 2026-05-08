#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import tempfile
from copy import deepcopy
from pathlib import Path
from typing import Any

try:
    import xlsxwriter
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Missing dependency: XlsxWriter. Run: python3 -m venv .venv && "
        ".venv/bin/pip install -r requirements.txt"
    ) from exc

try:
    from PIL import Image
except ModuleNotFoundError:
    Image = None

SCRIPT_DIR = Path(__file__).resolve().parent
GRAPHS_ROOT = SCRIPT_DIR.parent
REPO_ROOT = GRAPHS_ROOT.parent
DEFAULT_BANK_ORDER = ["RBC", "BMO", "Scotia", "CIBC", "National Bank", "TD"]
DEFAULT_OUTPUT_PATH = (
    REPO_ROOT
    / "outputs"
    / "attributed-growth-by-segment"
    / "attributed_growth_by_segment.xlsx"
)
DEFAULT_LOGO_DIR = GRAPHS_ROOT / "assets" / "bank-logos"
DEFAULT_LOGO_FILES = {
    "RBC": "ry.png",
    "BMO": "bmo.png",
    "Scotia": "bns.png",
    "CIBC": "cm.png",
    "National Bank": "na.png",
    "TD": "td.png",
}
NATIVE_ZERO_SLIVER_VALUE = 0.1
OVERLAY_LABEL_OFFSETS = [-0.1, 0.1, 0, 0.08, -0.08]
CENTER_COLS = [2, 4, 6, 8, 10, 12]
LAYOUT = {
    "absolute_label_row": 3,
    "total_label_row": 4,
    "chart_top_row": 5,
    "chart_top_col": 0,
    "logo_top_row": 29,
    "logo_bottom_row": 30,
    "source_table_start_row": 3,
}
CHART_GEOMETRY = {
    "axis_max_padding": 1.03,
    "zero_line_color": "#8C8C8C",
}
COLUMN_WIDTHS_PX = {
    0: 28,
    **{index: 84 for index in range(1, 14)},
    14: 21,
    15: 105,
    16: 118,
    17: 105,
    18: 70,
    19: 70,
    20: 70,
    21: 70,
    22: 85,
    23: 85,
    24: 85,
}
DEFAULT_CONFIG = {
    "sectionNumber": "1",
    "sectionTitle": "Attributed Growth by Segment",
    "sheetName": "Attributed Growth",
    "attributedLabel": "Attributed Growth/Contribution",
    "bankOrder": DEFAULT_BANK_ORDER,
    "requireCompleteBankOrder": True,
    "useLogos": True,
    "logoDir": str(DEFAULT_LOGO_DIR),
    "logoFiles": DEFAULT_LOGO_FILES,
    "segments": [
        {"key": "CB", "label": "CB", "color": "#40484A", "textColor": "#FFFFFF"},
        {"key": "IB", "label": "IB", "color": "#001A52", "textColor": "#FFFFFF"},
        {"key": "CM", "label": "CM", "color": "#2D5967", "textColor": "#FFFFFF"},
        {"key": "WM", "label": "WM", "color": "#F28C00", "textColor": "#000000"},
        {"key": "CS_INS", "label": "CS + Ins.", "color": "#000000", "textColor": "#FFFFFF"},
    ],
    "banks": [
        {
            "bank": "RBC",
            "absoluteGrowthBn": 7.4,
            "totalGrowthPct": 12,
            "brandColor": "#0B3A78",
            "segments": {"CB": 5, "IB": 1, "CM": 3, "WM": 3, "CS_INS": 0},
        },
        {
            "bank": "BMO",
            "absoluteGrowthBn": 3.0,
            "totalGrowthPct": 9,
            "brandColor": "#0072A8",
            "segments": {"CB": 2, "IB": 1, "CM": 2, "WM": 2, "CS_INS": 2},
        },
        {
            "bank": "Scotia",
            "absoluteGrowthBn": 3.7,
            "totalGrowthPct": 11,
            "brandColor": "#D71920",
            "segments": {"CB": 1, "IB": 0, "CM": 3, "WM": 2, "CS_INS": 4},
        },
        {
            "bank": "CIBC",
            "absoluteGrowthBn": 3.6,
            "totalGrowthPct": 13,
            "brandColor": "#8B1020",
            "segments": {"CB": 5, "IB": 1, "CM": 6, "WM": 2, "CS_INS": -1},
        },
        {
            "bank": "National Bank",
            "shortName": "National",
            "absoluteGrowthBn": 1.7,
            "totalGrowthPct": 15,
            "attributed": True,
            "brandColor": "#C6002B",
            "segments": {"CB": 2, "IB": 1, "CM": 7, "WM": 2, "CS_INS": 3},
        },
        {
            "bank": "TD",
            "absoluteGrowthBn": 5.9,
            "totalGrowthPct": 12,
            "attributed": True,
            "brandColor": "#2E7D32",
            "segments": {"CB": 2, "IB": 2, "CM": 3, "WM": 3, "CS_INS": 2},
        },
    ],
}

BANK_ALIASES = {
    "rbc": "RBC",
    "ry": "RBC",
    "royal bank of canada": "RBC",
    "bmo": "BMO",
    "bank of montreal": "BMO",
    "scotia": "Scotia",
    "scotiabank": "Scotia",
    "bns": "Scotia",
    "bank of nova scotia": "Scotia",
    "cibc": "CIBC",
    "cm": "CIBC",
    "canadian imperial bank of commerce": "CIBC",
    "national": "National Bank",
    "national bank": "National Bank",
    "national bank of canada": "National Bank",
    "na": "National Bank",
    "nbc": "National Bank",
    "td": "TD",
    "td bank": "TD",
    "toronto dominion": "TD",
    "toronto dominion bank": "TD",
}


def normalize_text(value: Any) -> str:
    text = str(value or "").strip().lower().replace("&", "and")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def canonical_bank_name(name: Any) -> str:
    raw = str(name or "").strip()
    return BANK_ALIASES.get(normalize_text(raw), raw)


def to_number(value: Any, fallback: float = 0) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.strip():
        cleaned = re.sub(r"[%$,]", "", value, flags=re.I)
        cleaned = re.sub(r"\bBN\b", "", cleaned, flags=re.I).strip()
        try:
            return float(cleaned)
        except ValueError:
            return fallback
    return fallback


def to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "yes", "y", "1"}
    return bool(value)


def row_value(row: dict[str, Any], keys: list[str]) -> Any:
    for key in keys:
        if key in row:
            return row[key]
    return None


def first_present(*values: Any) -> Any:
    for value in values:
        if value is not None:
            return value
    return None


def normalize_logo_files(logo_files: dict[str, str]) -> dict[str, str]:
    return {canonical_bank_name(bank): file_name for bank, file_name in logo_files.items()}


def normalize_bank_record(bank: dict[str, Any]) -> dict[str, Any]:
    canonical = canonical_bank_name(bank.get("bank"))
    defaults = next(
        (candidate for candidate in DEFAULT_CONFIG["banks"] if candidate["bank"] == canonical),
        {},
    )
    normalized = deepcopy(defaults)
    normalized.update(bank)
    normalized["bank"] = canonical
    normalized["absoluteGrowthBn"] = to_number(normalized.get("absoluteGrowthBn"))
    normalized["totalGrowthPct"] = to_number(normalized.get("totalGrowthPct"))
    normalized["shortName"] = bank.get("shortName", defaults.get("shortName"))
    normalized["brandColor"] = bank.get("brandColor", defaults.get("brandColor", "#111827"))
    normalized["segments"] = {
        **deepcopy(defaults.get("segments", {})),
        **{
            key: to_number(value)
            for key, value in deepcopy(bank.get("segments") or {}).items()
        },
    }
    return normalized


def normalize_data_rows(rows: list[dict[str, Any]], segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized_rows = []
    for row in rows:
        bank_name = row_value(row, ["bank", "Bank"])
        if not bank_name:
            continue
        segment_values = {}
        for segment in segments:
            segment_values[segment["key"]] = to_number(
                first_present(
                    (row.get("segments") or {}).get(segment["key"]),
                    (row.get("segments") or {}).get(segment["label"]),
                    row_value(row, [segment["key"], segment["label"]]),
                )
            )
        normalized = {
            "bank": bank_name,
            "absoluteGrowthBn": to_number(
                row_value(row, ["absoluteGrowthBn", "Absolute Growth ($BN)", "Absolute Growth"])
            ),
            "totalGrowthPct": to_number(
                row_value(row, ["totalGrowthPct", "Reported YoY Growth", "YoY Growth"])
            ),
            "segments": segment_values,
        }
        for target, keys in {
            "shortName": ["shortName", "Short Name", "ShortName"],
            "brandColor": ["brandColor", "Brand Color"],
            "logoPath": ["logoPath", "Logo Path"],
            "logoFile": ["logoFile", "Logo File"],
        }.items():
            value = row_value(row, keys)
            if value not in (None, ""):
                normalized[target] = value
        attributed = row_value(row, ["attributed", "Attributed"])
        if attributed not in (None, ""):
            normalized["attributed"] = to_bool(attributed)
        normalized_rows.append(normalized)
    return normalized_rows


def order_banks(
    banks: list[dict[str, Any]], bank_order: list[str], require_complete: bool = True
) -> list[dict[str, Any]]:
    normalized_banks = [normalize_bank_record(bank) for bank in banks]
    order = [canonical_bank_name(bank) for bank in bank_order]
    order_set = set(order)
    banks_by_name = {bank["bank"]: bank for bank in normalized_banks}
    missing = [bank for bank in order if bank not in banks_by_name]
    unknown = [bank["bank"] for bank in normalized_banks if bank["bank"] not in order_set]
    if missing and require_complete:
        raise ValueError(f"Missing required bank rows: {', '.join(missing)}")
    if unknown:
        raise ValueError(
            f"Unknown bank rows: {', '.join(unknown)}. Add them to bankOrder before rendering."
        )
    if len(order) > len(CENTER_COLS):
        raise ValueError(f"The fixed chart layout supports up to {len(CENTER_COLS)} banks.")
    return [banks_by_name[name] for name in order if name in banks_by_name]


def prepare_config(parsed: dict[str, Any]) -> dict[str, Any]:
    parsed = deepcopy(parsed)
    default_segments = DEFAULT_CONFIG["segments"]
    segments = []
    for segment in parsed.get("segments", default_segments):
        defaults = next(
            (
                default_segment
                for default_segment in default_segments
                if default_segment["key"] == segment.get("key")
                or default_segment["label"] == segment.get("label")
            ),
            {},
        )
        merged = deepcopy(defaults)
        merged.update(segment)
        segments.append(merged)

    raw_banks = (
        normalize_data_rows(parsed["data"], segments)
        if "data" in parsed
        else deepcopy(parsed.get("banks", DEFAULT_CONFIG["banks"]))
    )
    bank_order = parsed.get("bankOrder", DEFAULT_CONFIG["bankOrder"])
    config = deepcopy(DEFAULT_CONFIG)
    config.update(parsed)
    config["segments"] = segments
    config["bankOrder"] = bank_order
    config["banks"] = order_banks(
        raw_banks,
        bank_order,
        parsed.get("requireCompleteBankOrder", DEFAULT_CONFIG["requireCompleteBankOrder"]),
    )
    config["logoFiles"] = {
        **DEFAULT_LOGO_FILES,
        **normalize_logo_files(parsed.get("logoFiles", {})),
    }
    return config


def get_format_spec() -> str:
    """Return the JSON formatting instructions for the upstream LLM step."""
    return """
Format the calculation results into one JSON object for an attributed growth by segment chart.

Return only JSON. Use this structure:
{
  "section_number": "1",
  "section_title": "Attributed Growth by Segment",
  "sheet_name": "Attributed Growth",
  "segment_names": ["CB", "IB", "CM", "WM", "CS + Ins."],
  "banks": [
    {
      "bank": "RBC",
      "absolute_growth": 7.4,
      "yoy_pct": 12,
      "segments": {
        "CB": 5,
        "IB": 1,
        "CM": 3,
        "WM": 3,
        "CS + Ins.": 0
      }
    }
  ],
  "attributed_banks": ["National Bank", "TD"]
}

Field rules:
- `absolute_growth` is the displayed bank-level absolute growth in $BN.
- `yoy_pct` is the displayed bank-level YoY growth percentage.
- Segment values are contribution percentages, not dollar values.
- Keep bank order in the order the bars should render.
- Use `0` for zero contribution values so the chart can draw the 0% sliver and label.
- If the calculation payload stores segment metrics as objects, use the contribution or
  attribution percentage field for the segment value.
- `attributed_banks` controls which banks sit under the gray
  "Attributed Growth/Contribution" banner. Omit it to use the default banner placement.
""".strip()


def segment_key_from_name(name: Any) -> str:
    text = str(name or "").strip()
    default_lookup = {
        normalize_text(segment["key"]): segment["key"] for segment in DEFAULT_CONFIG["segments"]
    }
    default_lookup.update(
        {normalize_text(segment["label"]): segment["key"] for segment in DEFAULT_CONFIG["segments"]}
    )
    if normalize_text(text) in default_lookup:
        return default_lookup[normalize_text(text)]
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_").upper()


def normalize_formatted_segments(formatted: dict[str, Any]) -> list[dict[str, Any]]:
    raw_segments = formatted.get("segments") or formatted.get("segment_names")
    if not raw_segments:
        return deepcopy(DEFAULT_CONFIG["segments"])
    if isinstance(raw_segments, dict):
        raw_segments = [
            {"key": key, **value} if isinstance(value, dict) else {"key": key, "label": value}
            for key, value in raw_segments.items()
        ]

    segments = []
    for raw_segment in raw_segments:
        if isinstance(raw_segment, dict):
            label = first_present(
                raw_segment.get("label"),
                raw_segment.get("name"),
                raw_segment.get("segment"),
                raw_segment.get("key"),
            )
            key = raw_segment.get("key") or segment_key_from_name(label)
            override = raw_segment
        else:
            label = raw_segment
            key = segment_key_from_name(label)
            override = {}

        defaults = next(
            (
                segment
                for segment in DEFAULT_CONFIG["segments"]
                if segment["key"] == key or normalize_text(segment["label"]) == normalize_text(label)
            ),
            {},
        )
        merged = deepcopy(defaults)
        merged.update(override)
        merged["key"] = key
        merged["label"] = str(label or key)
        merged.setdefault("color", DEFAULT_CONFIG["segments"][len(segments) % len(DEFAULT_CONFIG["segments"])]["color"])
        merged.setdefault(
            "textColor",
            DEFAULT_CONFIG["segments"][len(segments) % len(DEFAULT_CONFIG["segments"])]["textColor"],
        )
        segments.append(merged)
    return segments


def segment_contribution_value(value: Any) -> Any:
    if not isinstance(value, dict):
        return value
    return first_present(
        value.get("contribution_pct"),
        value.get("attribution_pct"),
        value.get("attributed_pct"),
        value.get("contribution"),
        value.get("pct"),
        value.get("value"),
        value.get("yoy_pct"),
    )


def iter_formatted_bank_records(formatted: dict[str, Any]) -> list[dict[str, Any]]:
    banks = formatted.get("banks") or formatted.get("data") or []
    if isinstance(banks, dict):
        return [
            {"bank": bank_name, **(bank_data or {})}
            for bank_name, bank_data in banks.items()
            if bank_data is not None
        ]
    return [record for record in banks if isinstance(record, dict)]


def from_formatted(formatted: dict[str, Any]) -> dict[str, Any]:
    """Convert LLM-formatted JSON into the internal workbook config."""
    if not isinstance(formatted, dict):
        raise TypeError("formatted must be a dict")

    segments = normalize_formatted_segments(formatted)
    segment_lookup = {
        normalize_text(segment["key"]): segment["key"] for segment in segments
    }
    segment_lookup.update({normalize_text(segment["label"]): segment["key"] for segment in segments})
    raw_attributed_banks = formatted.get("attributed_banks", formatted.get("attributedBanks")) or []
    attributed_banks = {canonical_bank_name(bank) for bank in raw_attributed_banks}

    banks = []
    for record in iter_formatted_bank_records(formatted):
        bank_name = first_present(record.get("bank"), record.get("bank_name"), record.get("name"))
        if not bank_name:
            continue

        totals = record.get("totals") or record.get("total") or {}
        raw_segments = (
            record.get("segments")
            or record.get("segment_contributions")
            or record.get("segmentContributions")
            or {}
        )
        segment_values: dict[str, float] = {}
        if isinstance(raw_segments, dict):
            segment_items = raw_segments.items()
        else:
            segment_items = [
                (
                    first_present(item.get("segment"), item.get("segment_name"), item.get("name"), item.get("key")),
                    item,
                )
                for item in raw_segments
                if isinstance(item, dict)
            ]
        for raw_segment_name, raw_segment_value in segment_items:
            segment_key = segment_lookup.get(
                normalize_text(raw_segment_name), segment_key_from_name(raw_segment_name)
            )
            segment_values[segment_key] = to_number(segment_contribution_value(raw_segment_value))

        bank = {
            "bank": bank_name,
            "absoluteGrowthBn": to_number(
                first_present(
                    record.get("absoluteGrowthBn"),
                    record.get("absolute_growth"),
                    record.get("absolute_growth_bn"),
                    record.get("absoluteGrowth"),
                    totals.get("absolute_growth"),
                    totals.get("absolute_growth_bn"),
                    totals.get("absoluteGrowthBn"),
                )
            ),
            "totalGrowthPct": to_number(
                first_present(
                    record.get("totalGrowthPct"),
                    record.get("yoy_pct"),
                    record.get("yoy_growth"),
                    record.get("reported_yoy_growth"),
                    totals.get("yoy_pct"),
                    totals.get("yoy_growth"),
                    totals.get("totalGrowthPct"),
                )
            ),
            "segments": segment_values,
        }
        for optional_key in ("shortName", "brandColor", "logoPath", "logoFile"):
            snake_key = re.sub(r"([A-Z])", r"_\1", optional_key).lower()
            optional_value = first_present(record.get(optional_key), record.get(snake_key))
            if optional_value not in (None, ""):
                bank[optional_key] = optional_value
        if attributed_banks:
            bank["attributed"] = canonical_bank_name(bank_name) in attributed_banks
        elif "attributed" in record:
            bank["attributed"] = to_bool(record["attributed"])
        banks.append(bank)

    config = {
        "sectionNumber": str(
            first_present(
                formatted.get("sectionNumber"),
                formatted.get("section_number"),
                DEFAULT_CONFIG["sectionNumber"],
            )
        ),
        "sectionTitle": first_present(
            formatted.get("sectionTitle"),
            formatted.get("section_title"),
            DEFAULT_CONFIG["sectionTitle"],
        ),
        "sheetName": first_present(
            formatted.get("sheetName"),
            formatted.get("sheet_name"),
            DEFAULT_CONFIG["sheetName"],
        ),
        "attributedLabel": first_present(
            formatted.get("attributedLabel"),
            formatted.get("attributed_label"),
            DEFAULT_CONFIG["attributedLabel"],
        ),
        "segments": segments,
        "bankOrder": [
            canonical_bank_name(bank)
            for bank in first_present(
                formatted.get("bankOrder"),
                formatted.get("bank_order"),
                [bank["bank"] for bank in banks],
            )
        ],
        "banks": banks,
    }
    return prepare_config(config)


def load_config(input_path: Path | None) -> dict[str, Any]:
    if not input_path:
        return prepare_config({})
    parsed = json.loads(input_path.read_text())
    formatted_keys = {
        "section_number",
        "section_title",
        "sheet_name",
        "segment_names",
        "attributed_banks",
    }
    if formatted_keys.intersection(parsed):
        return from_formatted(parsed)
    return prepare_config(parsed)


def safe_sheet_name(config: dict[str, Any]) -> str:
    raw = str(config.get("sheetName") or config.get("sectionTitle") or "Segment Growth")
    cleaned = re.sub(r"[\[\]:*?/\\]", " ", raw)[:31].strip()
    return cleaned or "Segment Growth"


def section_header(config: dict[str, Any]) -> str:
    return (
        f'{config["sectionNumber"]}. {config["sectionTitle"]}'
        if config.get("sectionNumber")
        else config["sectionTitle"]
    )


def segment_value(bank: dict[str, Any], segment_key: str) -> float:
    value = bank.get("segments", {}).get(segment_key, 0)
    return float(value) if isinstance(value, (int, float)) else 0


def pct_label(value: float) -> str:
    return f"{round(value)}%"


def row_height_px(row_index: int) -> int:
    if row_index == 0:
        return 36
    if row_index == 1:
        return 24
    if row_index == 2:
        return 10
    return 20


def column_width_px(column_index: int) -> int:
    return COLUMN_WIDTHS_PX.get(column_index, 64)


def column_left_px(column_index: int) -> int:
    return sum(column_width_px(col) for col in range(column_index))


def category_center_px(bank_index: int) -> float:
    column = CENTER_COLS[bank_index]
    return column_left_px(column) + column_width_px(column) / 2


def pixel_to_col_anchor(x_px: float) -> tuple[int, int]:
    col = 0
    remaining = max(0, float(x_px))
    while remaining >= column_width_px(col):
        remaining -= column_width_px(col)
        col += 1
    return col, round(remaining)


def resolve_logo_path(bank: dict[str, Any], config: dict[str, Any]) -> Path | None:
    if config.get("useLogos") is False:
        return None
    if bank.get("logoPath"):
        explicit = Path(bank["logoPath"]).expanduser().resolve()
        return explicit if explicit.is_file() else None
    logo_file = bank.get("logoFile") or config.get("logoFiles", {}).get(bank["bank"])
    logo_dir = config.get("logoDir")
    if not logo_file or not logo_dir:
        return None
    candidate = Path(logo_dir).expanduser().resolve() / logo_file
    return candidate if candidate.is_file() else None


def png_dimensions(file_path: Path) -> tuple[int, int] | None:
    data = file_path.read_bytes()
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    return None


def normalized_logo_path(file_path: Path, temp_dir: Path) -> Path:
    if Image is None:
        return file_path

    image = Image.open(file_path).convert("RGBA")
    alpha_bbox = image.getchannel("A").getbbox()
    bbox = alpha_bbox or image.getbbox()
    if bbox:
        image = image.crop(bbox)

    output_path = temp_dir / file_path.name
    image.save(output_path)
    return output_path


def logo_layout(file_path: Path) -> tuple[int, int]:
    max_width, max_height = 126, 40
    dimensions = png_dimensions(file_path)
    if not dimensions:
        return 96, 34
    width, height = dimensions
    scale = min(max_width / width, max_height / height)
    return round(width * scale), round(height * scale)


def logo_position(logo_path: Path, bank_index: int) -> tuple[int, int, int, int, int, int]:
    center_x = category_center_px(bank_index)
    width_px, height_px = logo_layout(logo_path)
    total_label_height_px = 44
    x_px = center_x - width_px / 2
    col, x_offset = pixel_to_col_anchor(x_px)
    y_offset = max(0, round((total_label_height_px - height_px) / 2))
    return LAYOUT["logo_top_row"], col, x_offset, y_offset, width_px, height_px


def make_formats(workbook: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_size": 20,
                "font_color": "#FFFFFF",
                "bg_color": "#071734",
                "valign": "vcenter",
            }
        ),
        "top_abs": workbook.add_format(
            {
                "bold": True,
                "font_size": 17,
                "align": "center",
                "num_format": '+0.0"BN";-0.0"BN";0.0"BN"',
            }
        ),
        "top_pct": workbook.add_format(
            {"bold": True, "font_size": 15, "align": "center", "num_format": '0"%"'}
        ),
        "attributed": workbook.add_format(
            {
                "bold": True,
                "font_size": 15,
                "font_color": "#111827",
                "bg_color": "#BDBDBD",
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "source_header": workbook.add_format(
            {"bold": True, "font_size": 9, "font_color": "#FFFFFF", "bg_color": "#0D6786"}
        ),
        "source": workbook.add_format({"font_size": 9}),
        "source_bn": workbook.add_format({"font_size": 9, "num_format": '+0.0"BN";-0.0"BN";0.0"BN"'}),
        "source_pct": workbook.add_format({"font_size": 9, "num_format": '0"%"'}),
    }


def setup_common_sheet(worksheet: xlsxwriter.worksheet.Worksheet) -> None:
    worksheet.hide_gridlines(2)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 1)
    worksheet.set_margins(0.3, 0.3, 0.3, 0.3)
    worksheet.print_area(0, 0, 31, 13)
    worksheet.center_horizontally()
    for col in range(25):
        worksheet.set_column_pixels(col, col, COLUMN_WIDTHS_PX.get(col, 64))
    for row in range(32):
        worksheet.set_row_pixels(row, row_height_px(row))


def write_source_table(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.worksheet.Worksheet,
    config: dict[str, Any],
    formats: dict[str, Any],
) -> tuple[int, int, int, int]:
    header = [
        "Bank",
        "Absolute Growth ($BN)",
        "Reported YoY Growth",
        *[segment["label"] for segment in config["segments"]],
    ]
    rows = [
        [
            bank["bank"],
            bank["absoluteGrowthBn"],
            bank["totalGrowthPct"],
            *[
                segment_value(bank, segment["key"])
                for segment in config["segments"]
            ],
        ]
        for bank in config["banks"]
    ]
    start_row = LAYOUT["source_table_start_row"]
    start_col = 15
    end_row = start_row + len(rows)
    end_col = start_col + len(header) - 1

    worksheet.add_table(
        start_row,
        start_col,
        end_row,
        end_col,
        {
            "name": "AttributedGrowthBySegmentSource",
            "style": "Table Style Medium 2",
            "data": rows,
            "columns": [{"header": title} for title in header],
        },
    )
    for col_offset, title in enumerate(header):
        worksheet.write(start_row, start_col + col_offset, title, formats["source_header"])
    for row_offset, row in enumerate(rows, start=1):
        for col_offset, value in enumerate(row):
            fmt = formats["source"]
            if col_offset == 1:
                fmt = formats["source_bn"]
            elif 2 <= col_offset <= 2 + len(config["segments"]):
                fmt = formats["source_pct"]
            worksheet.write(start_row + row_offset, start_col + col_offset, value, fmt)
    return start_row, start_col, end_row, end_col


def write_native_chart_helper_values(
    worksheet: xlsxwriter.worksheet.Worksheet,
    config: dict[str, Any],
    source_bounds: tuple[int, int, int, int],
) -> int:
    start_row, start_col, end_row, end_col = source_bounds
    helper_start_col = end_col + 2

    for index, segment in enumerate(config["segments"]):
        helper_col = helper_start_col + index
        source_col = start_col + 3 + index
        worksheet.write(start_row, helper_col, f"{segment['label']} Chart Value")
        for row in range(start_row + 1, end_row + 1):
            source_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, source_col)
            source_value = segment_value(config["banks"][row - start_row - 1], segment["key"])
            helper_value = NATIVE_ZERO_SLIVER_VALUE if source_value == 0 else source_value
            worksheet.write_formula(
                row,
                helper_col,
                f'=IF({source_cell}="",NA(),IF({source_cell}=0,{NATIVE_ZERO_SLIVER_VALUE},{source_cell}))',
                None,
                helper_value,
            )

    worksheet.set_column(helper_start_col, helper_start_col + len(config["segments"]) - 1, None, None, {"hidden": True})
    return helper_start_col


def sheet_formula(sheet_name: str, row: int, col: int) -> str:
    cell = xlsxwriter.utility.xl_rowcol_to_cell(row, col, row_abs=True, col_abs=True)
    return f"'{sheet_name}'!{cell}"


def write_overlay_label_helper_values(
    worksheet: xlsxwriter.worksheet.Worksheet,
    sheet_name: str,
    config: dict[str, Any],
    source_bounds: tuple[int, int, int, int],
    chart_values_start_col: int,
) -> int:
    start_row, start_col, end_row, end_col = source_bounds
    helper_start_col = chart_values_start_col + len(config["segments"]) + 2
    source_segment_start_col = start_col + 3

    for segment_index, segment in enumerate(config["segments"]):
        x_col = helper_start_col + segment_index * 3
        y_col = x_col + 1
        label_col = x_col + 2
        worksheet.write(start_row, x_col, f"{segment['label']} Label X")
        worksheet.write(start_row, y_col, f"{segment['label']} Label Y")
        worksheet.write(start_row, label_col, f"{segment['label']} Label")

        chart_col = chart_values_start_col + segment_index
        source_col = source_segment_start_col + segment_index
        x_offset = OVERLAY_LABEL_OFFSETS[segment_index % len(OVERLAY_LABEL_OFFSETS)]

        for row in range(start_row + 1, end_row + 1):
            bank_number = row - start_row
            source_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, source_col)
            chart_cell = xlsxwriter.utility.xl_rowcol_to_cell(row, chart_col)

            if segment_index == 0:
                prev_pos = "0"
                prev_neg = "0"
            else:
                prev_first = xlsxwriter.utility.xl_rowcol_to_cell(row, chart_values_start_col)
                prev_last = xlsxwriter.utility.xl_rowcol_to_cell(
                    row, chart_values_start_col + segment_index - 1
                )
                prev_range = f"{prev_first}:{prev_last}"
                prev_pos = f'SUMIF({prev_range},">0",{prev_range})'
                prev_neg = f'SUMIF({prev_range},"<0",{prev_range})'

            source_value = segment_value(config["banks"][row - start_row - 1], segment["key"])
            chart_value = NATIVE_ZERO_SLIVER_VALUE if source_value == 0 else source_value
            cached_y = (
                sum(
                    max(
                        0,
                        NATIVE_ZERO_SLIVER_VALUE
                        if segment_value(config["banks"][row - start_row - 1], previous["key"]) == 0
                        else segment_value(config["banks"][row - start_row - 1], previous["key"]),
                    )
                    for previous in config["segments"][:segment_index]
                )
                + chart_value / 2
                if chart_value >= 0
                else sum(
                    min(
                        0,
                        NATIVE_ZERO_SLIVER_VALUE
                        if segment_value(config["banks"][row - start_row - 1], previous["key"]) == 0
                        else segment_value(config["banks"][row - start_row - 1], previous["key"]),
                    )
                    for previous in config["segments"][:segment_index]
                )
                + chart_value / 2
            )
            cached_x = bank_number + (x_offset if 0 <= abs(source_value) <= 1 else 0)
            cached_label = "" if source_value == "" else pct_label(source_value)

            worksheet.write_formula(
                row,
                x_col,
                f'=IF({source_cell}="",NA(),{bank_number}+IF(AND(ABS({source_cell})<=1,{source_cell}>=0),{x_offset},0))',
                None,
                cached_x,
            )
            worksheet.write_formula(
                row,
                y_col,
                f'=IF({source_cell}="",NA(),IF({chart_cell}>=0,{prev_pos}+{chart_cell}/2,{prev_neg}+{chart_cell}/2))',
                None,
                cached_y,
            )
            worksheet.write_formula(
                row,
                label_col,
                f'=IF({source_cell}="","",ROUND({source_cell},0)&"%")',
                None,
                cached_label,
            )

    worksheet.set_column(
        helper_start_col,
        helper_start_col + len(config["segments"]) * 3 - 1,
        None,
        None,
        {"hidden": True},
    )
    return helper_start_col


def add_native_chart(
    workbook: xlsxwriter.Workbook,
    worksheet: xlsxwriter.worksheet.Worksheet,
    sheet_name: str,
    config: dict[str, Any],
    source_bounds: tuple[int, int, int, int],
    chart_values_start_col: int,
    overlay_label_helper_start_col: int,
) -> None:
    start_row, start_col, end_row, _end_col = source_bounds
    chart = workbook.add_chart({"type": "column", "subtype": "stacked"})
    chart.show_hidden_data()
    category_range = [sheet_name, start_row + 1, start_col, end_row, start_col]
    for index, segment in enumerate(config["segments"]):
        name_col = start_col + 3 + index
        values_col = chart_values_start_col + index
        chart.add_series(
            {
                "name": [sheet_name, start_row, name_col],
                "categories": category_range,
                "values": [sheet_name, start_row + 1, values_col, end_row, values_col],
                "fill": {"color": segment["color"]},
                "border": {"color": segment["color"]},
            }
        )
    extents = segment_stack_extents(config)
    axis_max = max(1, extents["positive"] * CHART_GEOMETRY["axis_max_padding"])
    axis_min = -max(1.5, extents["negative"] * 1.5) if extents["negative"] > 0 else 0
    chart.set_title({"none": True})
    chart.set_legend({"none": True})
    chart.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
    chart.set_plotarea(
        {
            "fill": {"color": "#FFFFFF"},
            "border": {"none": True},
            "layout": {
                "x": 0.06135,
                "y": 0.07,
                "width": 0.8834,
                "height": 0.92,
            },
        }
    )
    chart.set_x_axis(
        {
            "major_gridlines": {"visible": False},
            "major_tick_mark": "none",
            "minor_tick_mark": "none",
            "line": {"color": CHART_GEOMETRY["zero_line_color"], "width": 1.25},
            "num_font": {"color": "#FFFFFF", "size": 1},
        }
    )
    chart.set_y_axis(
        {
            "visible": False,
            "min": axis_min,
            "max": axis_max,
            "major_unit": 5,
            "major_gridlines": {"visible": False},
            "line": {"none": True},
        }
    )
    chart.set_size({"width": column_left_px(15), "height": 470})
    chart.series_gap_1 = 70
    chart.series_gap_2 = 70
    chart.series_overlap_1 = 100
    chart.series_overlap_2 = 100

    label_chart = workbook.add_chart({"type": "scatter", "subtype": "marker_only"})
    label_chart.show_hidden_data()
    for segment_index, segment in enumerate(config["segments"]):
        x_col = overlay_label_helper_start_col + segment_index * 3
        y_col = x_col + 1
        label_col = x_col + 2
        custom_labels = []
        for row in range(start_row + 1, end_row + 1):
            custom_labels.append(
                {
                    "formula": sheet_formula(sheet_name, row, label_col),
                    "data": [
                        pct_label(
                            segment_value(config["banks"][row - start_row - 1], segment["key"])
                        )
                    ],
                }
            )
        label_chart.add_series(
            {
                "categories": [sheet_name, start_row + 1, x_col, end_row, x_col],
                "values": [sheet_name, start_row + 1, y_col, end_row, y_col],
                "x2_axis": True,
                "y2_axis": True,
                "marker": {"type": "none"},
                "line": {"none": True},
                "data_labels": {
                    "custom": custom_labels,
                    "position": "center",
                    "font": {
                        "color": segment.get("textColor", "#FFFFFF"),
                        "size": 12,
                        "bold": True,
                    },
                    "fill": {"color": segment["color"]},
                    "border": {"color": segment["color"]},
                },
            }
        )
    hidden_axis = {
        "visible": False,
        "major_gridlines": {"visible": False},
        "line": {"none": True},
        "num_font": {"color": "#FFFFFF", "size": 1},
    }
    label_chart.set_x2_axis(
        {
            **hidden_axis,
            "min": 0.5,
            "max": len(config["banks"]) + 0.5,
            "major_unit": 1,
        }
    )
    label_chart.set_y2_axis(
        {
            **hidden_axis,
            "min": axis_min,
            "max": axis_max,
            "major_unit": 5,
        }
    )
    chart.combine(label_chart)

    worksheet.insert_chart(LAYOUT["chart_top_row"], LAYOUT["chart_top_col"], chart)


def segment_stack_extents(config: dict[str, Any]) -> dict[str, float]:
    positive_max = 0.0
    negative_max = 0.0
    for bank in config["banks"]:
        positive = 0.0
        negative = 0.0
        for segment in config["segments"]:
            value = segment_value(bank, segment["key"])
            if value > 0:
                positive += value
            elif value < 0:
                negative += abs(value)
        positive_max = max(positive_max, positive)
        negative_max = max(negative_max, negative)
    return {"positive": positive_max, "negative": negative_max}


def write_attributed_growth_workbook(config: dict[str, Any], output_path: Path) -> None:
    sheet_name = safe_sheet_name(config)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet(sheet_name)
    formats = make_formats(workbook)
    setup_common_sheet(worksheet)

    worksheet.merge_range(0, 0, 0, 24, section_header(config), formats["title"])

    legend_cells = [1, 2, 3, 4, 5]
    for col, segment in zip(legend_cells, config["segments"]):
        legend_format = workbook.add_format({"font_size": 14, "font_color": segment["color"]})
        worksheet.write(1, col, f"{chr(9632)} {segment['label']}", legend_format)

    attributed = [(bank, index) for index, bank in enumerate(config["banks"]) if bank.get("attributed")]
    if attributed:
        first = CENTER_COLS[attributed[0][1]]
        last = CENTER_COLS[attributed[-1][1]] + 1
        worksheet.merge_range(1, first, 1, last, config["attributedLabel"], formats["attributed"])

    source_bounds = write_source_table(workbook, worksheet, config, formats)
    source_start_row, source_start_col, _source_end_row, _source_end_col = source_bounds

    for index, bank in enumerate(config["banks"]):
        if index >= len(CENTER_COLS):
            continue
        col = CENTER_COLS[index]
        source_row = source_start_row + 1 + index
        abs_cell = xlsxwriter.utility.xl_rowcol_to_cell(source_row, source_start_col + 1)
        pct_cell = xlsxwriter.utility.xl_rowcol_to_cell(source_row, source_start_col + 2)
        abs_value = bank["absoluteGrowthBn"]
        pct_value = bank["totalGrowthPct"]
        worksheet.write_formula(
            LAYOUT["absolute_label_row"],
            col,
            f'=IF({abs_cell}="","",{abs_cell})',
            formats["top_abs"],
            abs_value,
        )
        worksheet.write_formula(
            LAYOUT["total_label_row"],
            col,
            f'=IF({pct_cell}="","",{pct_cell})',
            formats["top_pct"],
            pct_value,
        )

    chart_values_start_col = write_native_chart_helper_values(worksheet, config, source_bounds)
    overlay_label_helper_start_col = write_overlay_label_helper_values(
        worksheet,
        sheet_name,
        config,
        source_bounds,
        chart_values_start_col,
    )
    add_native_chart(
        workbook,
        worksheet,
        sheet_name,
        config,
        source_bounds,
        chart_values_start_col,
        overlay_label_helper_start_col,
    )

    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        for index, bank in enumerate(config["banks"]):
            if index >= len(CENTER_COLS):
                continue
            logo_path = resolve_logo_path(bank, config)
            col = CENTER_COLS[index]
            if logo_path:
                display_logo = normalized_logo_path(logo_path, temp_dir)
                row, image_col, x_offset, y_offset, width_px, height_px = logo_position(
                    display_logo, index
                )
                dimensions = png_dimensions(display_logo) or (width_px, height_px)
                worksheet.insert_image(
                    row,
                    image_col,
                    str(display_logo),
                    {
                        "x_offset": x_offset,
                        "y_offset": y_offset,
                        "x_scale": width_px / dimensions[0],
                        "y_scale": height_px / dimensions[1],
                        "object_position": 2,
                    },
                )
            else:
                brand_format = workbook.add_format(
                    {
                        "bold": True,
                        "font_size": 14 if bank["bank"] == "National Bank" else 18,
                        "font_color": bank.get("brandColor", "#111827"),
                        "align": "center",
                        "valign": "vcenter",
                    }
                )
                worksheet.merge_range(
                    LAYOUT["logo_top_row"],
                    col,
                    LAYOUT["logo_bottom_row"],
                    col,
                    bank.get("shortName") or bank["bank"],
                    brand_format,
                )

        workbook.close()


def write_workbook(
    config: dict[str, Any], output_path: Path, supporting_data: Any | None = None
) -> None:
    """Pipeline entry point: render an XLSX workbook from an internal config dict."""
    _ = supporting_data
    write_attributed_growth_workbook(prepare_config(config), Path(output_path))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build attributed growth by segment workbook.")
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output = args.output or DEFAULT_OUTPUT_PATH
    config = load_config(args.input)
    write_workbook(config, output.resolve())
    print(output.resolve())


if __name__ == "__main__":
    main()
