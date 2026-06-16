import io
import re
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
FIXED_SHIFT_CODES = ["D", "N", "OFF", "POST"]


def clean_code(value: object) -> str:
    """Normalize shift codes so D/d, n/N, off/OFF all match."""
    if value is None:
        return ""
    return str(value).strip().upper()


def get_default_monday() -> date:
    today = date.today()
    days_until_monday = (7 - today.weekday()) % 7
    return today + timedelta(days=days_until_monday)


def get_schedule_date(start_date: date, week: int, day: str) -> date:
    """Return the actual date for a displayed week/day.

    start_date is the Monday date for Week 1.
    Therefore, Week 1 Sunday is the day before start_date.
    """
    monday = start_date + timedelta(days=(week - 1) * 7)

    offsets = {
        "Sun": -1,
        "Mon": 0,
        "Tue": 1,
        "Wed": 2,
        "Thu": 3,
        "Fri": 4,
        "Sat": 5,
    }

    return monday + timedelta(days=offsets[day])


def day_label(day: str, start_date: date, week: int) -> str:
    d = get_schedule_date(start_date, week, day)
    return f"{day}\n{d.month}/{d.day}"


def make_resident_names(n_residents: int) -> list[str]:
    default_names = [
        "Resident 1",
        "Resident 2",
        "Resident 3",
        "Resident 4",
        "Resident 5",
        "Resident 6",
    ]

    names = []

    for i in range(n_residents):
        default = default_names[i] if i < len(default_names) else f"Resident {i + 1}"

        names.append(
            st.sidebar.text_input(
                f"Resident {i + 1}",
                value=default,
                key=f"resident_{i}",
            ).strip()
            or default
        )

    return names


def parse_extra_shift_codes(extra_text: str) -> list[dict[str, float | str]]:
    rows = []

    for item in extra_text.split(","):
        if "=" not in item:
            continue

        code, hours = item.split("=", 1)
        code = clean_code(code)

        if not code or code in FIXED_SHIFT_CODES:
            continue

        try:
            hours_value = float(hours.strip())
        except ValueError:
            hours_value = 0.0

        rows.append({"Code": code, "Hours": hours_value})

    return rows


def make_shift_definitions() -> pd.DataFrame:
    """Sidebar shift-hour fields.

    D, N, OFF, and POST cannot be deleted. Only the hours can be edited.
    Optional extra shift codes can be added as CODE=hours.
    """
    st.sidebar.subheader("Shift hours")

    if "d_hours_widget" not in st.session_state:
        st.session_state.d_hours_widget = 12.0

    if "n_hours_widget" not in st.session_state:
        st.session_state.n_hours_widget = 13.0

    if "off_hours_widget" not in st.session_state:
        st.session_state.off_hours_widget = 0.0

    if "post_hours_widget" not in st.session_state:
        st.session_state.post_hours_widget = 0.0

    if "extra_shift_codes_widget" not in st.session_state:
        st.session_state.extra_shift_codes_widget = ""

    d_hours = st.sidebar.number_input(
        "D hours",
        min_value=0.0,
        step=0.5,
        key="d_hours_widget",
    )

    n_hours = st.sidebar.number_input(
        "N hours",
        min_value=0.0,
        step=0.5,
        key="n_hours_widget",
    )

    off_hours = st.sidebar.number_input(
        "OFF hours",
        min_value=0.0,
        step=0.5,
        key="off_hours_widget",
    )

    post_hours = st.sidebar.number_input(
        "POST hours",
        min_value=0.0,
        step=0.5,
        key="post_hours_widget",
    )

    rows = [
        {"Code": "D", "Hours": float(d_hours)},
        {"Code": "N", "Hours": float(n_hours)},
        {"Code": "OFF", "Hours": float(off_hours)},
        {"Code": "POST", "Hours": float(post_hours)},
    ]

    extra_text = st.sidebar.text_input(
        "Optional extra shift codes",
        help="Example: CLINIC=8, VAC=0",
        key="extra_shift_codes_widget",
    )

    rows.extend(parse_extra_shift_codes(extra_text))

    return pd.DataFrame(rows).drop_duplicates(subset=["Code"], keep="first").reset_index(drop=True)


def add_missing_schedule_codes_to_shift_defs(
    shift_defs: pd.DataFrame,
    schedule_df: pd.DataFrame | None,
) -> pd.DataFrame:
    """Keep current schedule values available as dropdown options."""
    if schedule_df is None or schedule_df.empty:
        return shift_defs

    known_codes = set(shift_defs["Code"].map(clean_code).tolist())
    schedule_codes: set[str] = set()

    for day in DAYS:
        if day in schedule_df.columns:
            schedule_codes.update(schedule_df[day].dropna().map(clean_code).tolist())

    missing_codes = sorted(code for code in schedule_codes if code and code not in known_codes)

    if not missing_codes:
        return shift_defs

    extra_rows = pd.DataFrame({"Code": missing_codes, "Hours": [0.0] * len(missing_codes)})

    return pd.concat([shift_defs, extra_rows], ignore_index=True)


def build_blank_schedule(residents: list[str], num_weeks: int, default_code: str) -> pd.DataFrame:
    rows = []

    for week in range(1, num_weeks + 1):
        for resident in residents:
            row = {"Week": week, "Resident": resident}

            for day in DAYS:
                row[day] = default_code

            rows.append(row)

    return pd.DataFrame(rows)


def reconcile_schedule(
    existing: pd.DataFrame | None,
    residents: list[str],
    num_weeks: int,
    default_code: str,
) -> pd.DataFrame:
    """Keep prior edits when controls change, but rebuild missing resident/week rows."""
    blank = build_blank_schedule(residents, num_weeks, default_code)

    if existing is None or existing.empty:
        return blank

    existing = existing.copy()

    for col in ["Week", "Resident", *DAYS]:
        if col not in existing.columns:
            existing[col] = default_code if col in DAYS else None

    merged = blank[["Week", "Resident"]].merge(
        existing[["Week", "Resident", *DAYS]],
        on=["Week", "Resident"],
        how="left",
    )

    for day in DAYS:
        merged[day] = merged[day].fillna(default_code).map(clean_code)

    return merged[["Week", "Resident", *DAYS]]


def clear_schedule_editor_state() -> None:
    """Clear Streamlit data_editor keys so imported/new data renders cleanly."""
    for key in list(st.session_state.keys()):
        if key.startswith("schedule_editor_week_") or key == "schedule_editor":
            del st.session_state[key]


def normalize_day_header(value: object) -> str:
    """Normalize Excel day headers.

    This accepts headers such as:
    - Sun
    - Sunday
    - Sun\\n6/30
    - Sun 6/30
    """
    if value is None:
        return ""

    text = str(value).strip().upper().replace("\n", " ")
    first_word = text.split()[0] if text.split() else ""

    aliases = {
        "SUNDAY": "Sun",
        "SUN": "Sun",
        "MONDAY": "Mon",
        "MON": "Mon",
        "TUESDAY": "Tue",
        "TUES": "Tue",
        "TUE": "Tue",
        "WEDNESDAY": "Wed",
        "WED": "Wed",
        "THURSDAY": "Thu",
        "THURS": "Thu",
        "THUR": "Thu",
        "THU": "Thu",
        "FRIDAY": "Fri",
        "FRI": "Fri",
        "SATURDAY": "Sat",
        "SAT": "Sat",
    }

    return aliases.get(first_word, "")


def row_has_day_headers(ws, row_num: int) -> tuple[bool, int | None, int | None]:
    """
    Return whether a worksheet row contains Sun-Sat headers.

    Returns:
        has_headers, resident_column, first_day_column
    """
    max_col_to_check = min(ws.max_column, 20)

    normalized = [
        normalize_day_header(ws.cell(row_num, col).value)
        for col in range(1, max_col_to_check + 1)
    ]

    for idx in range(0, len(normalized) - 6):
        if normalized[idx: idx + 7] == DAYS:
            first_day_col = idx + 1  # 1-based Excel column
            resident_col = first_day_col - 1

            if resident_col < 1:
                resident_col = 1

            return True, resident_col, first_day_col

    return False, None, None


def parse_week_number(value: object, fallback: int) -> int:
    """Read 'Week 1' as 1; otherwise use the fallback counter."""
    if value is None:
        return fallback

    match = re.search(r"\d+", str(value))

    if match:
        return int(match.group())

    return fallback


def default_shift_defs() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Code": ["D", "N", "OFF", "POST"],
            "Hours": [12.0, 13.0, 0.0, 0.0],
        }
    )


def import_shift_definitions_from_workbook(wb) -> pd.DataFrame:
    """Import Shift_Definitions tab if it exists; otherwise use defaults."""
    if "Shift_Definitions" not in wb.sheetnames:
        return default_shift_defs()

    ws = wb["Shift_Definitions"]
    rows = []

    # Expected layout:
    # A1 = Shift Code, B1 = Hours
    # A2/B2 onward = codes/hours
    for row_num in range(2, ws.max_row + 1):
        code = clean_code(ws.cell(row_num, 1).value)

        if not code:
            continue

        raw_hours = ws.cell(row_num, 2).value

        try:
            hours = float(raw_hours)
        except (TypeError, ValueError):
            hours = 0.0

        rows.append({"Code": code, "Hours": hours})

    if not rows:
        return default_shift_defs()

    out = pd.DataFrame(rows)
    out = out.drop_duplicates(subset=["Code"], keep="first").reset_index(drop=True)

    return out


def import_start_date_from_workbook(wb) -> date | None:
    """Import start date from Metadata sheet if available."""
    if "Metadata" not in wb.sheetnames:
        return None

    ws = wb["Metadata"]

    # Preferred layout created by this app:
    # A1 = Start_Date, B1 = YYYY-MM-DD
    raw_value = ws["B1"].value

    if raw_value is None:
        return None

    if isinstance(raw_value, datetime):
        return raw_value.date()

    if isinstance(raw_value, date):
        return raw_value

    try:
        return datetime.strptime(str(raw_value).strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def import_schedule_from_excel(uploaded_file) -> tuple[pd.DataFrame, list[str], int, pd.DataFrame, date | None]:
    """
    Import the Schedule tab from a workbook produced by this app.

    The importer looks for repeated Sun-Mon-Tue-Wed-Thu-Fri-Sat header rows,
    then reads resident names from the column immediately before Sun and shift
    codes from the seven day columns.
    """
    wb = load_workbook(uploaded_file, data_only=False)

    if "Schedule" not in wb.sheetnames:
        raise ValueError("I could not find a tab named 'Schedule' in that workbook.")

    ws = wb["Schedule"]
    imported_shift_defs = import_shift_definitions_from_workbook(wb)
    imported_start_date = import_start_date_from_workbook(wb)

    rows = []
    residents = []
    week_counter = 0
    row_num = 1

    while row_num <= ws.max_row:
        has_headers, resident_col, first_day_col = row_has_day_headers(ws, row_num)

        if not has_headers:
            row_num += 1
            continue

        week_counter += 1
        week_label = ws.cell(row_num, resident_col).value
        week_num = parse_week_number(week_label, week_counter)

        data_row = row_num + 1

        while data_row <= ws.max_row:
            next_has_headers, _, _ = row_has_day_headers(ws, data_row)

            if next_has_headers:
                break

            resident = ws.cell(data_row, resident_col).value
            resident = "" if resident is None else str(resident).strip()

            day_values = {
                day: clean_code(ws.cell(data_row, first_day_col + i).value)
                for i, day in enumerate(DAYS)
            }

            has_any_shift = any(value != "" for value in day_values.values())

            if resident and has_any_shift:
                if resident not in residents:
                    residents.append(resident)

                rows.append(
                    {
                        "Week": week_num,
                        "Resident": resident,
                        **day_values,
                    }
                )

            data_row += 1

        row_num = data_row

    if not rows:
        raise ValueError(
            "I found the Schedule tab, but I could not detect repeated Sun-Sat schedule rows."
        )

    schedule_df = pd.DataFrame(rows)

    # Fill blanks with OFF so the Streamlit dropdowns have valid values.
    for day in DAYS:
        schedule_df[day] = schedule_df[day].replace("", "OFF").map(clean_code)

    num_weeks = int(schedule_df["Week"].max())

    imported_shift_defs = add_missing_schedule_codes_to_shift_defs(imported_shift_defs, schedule_df)

    return schedule_df, residents, num_weeks, imported_shift_defs, imported_start_date


def apply_imported_shift_defs_to_widgets(imported_shift_defs: pd.DataFrame) -> None:
    """Push imported Shift_Definitions into fixed sidebar fields and optional extra field."""
    fixed_map = {"D": 12.0, "N": 13.0, "OFF": 0.0, "POST": 0.0}
    extra_parts = []

    for row in imported_shift_defs.itertuples():
        code = clean_code(row.Code)

        try:
            hours = float(row.Hours)
        except (TypeError, ValueError):
            hours = 0.0

        if code in fixed_map:
            fixed_map[code] = hours
        elif code:
            extra_parts.append(f"{code}={hours:g}")

    st.session_state.d_hours_widget = fixed_map["D"]
    st.session_state.n_hours_widget = fixed_map["N"]
    st.session_state.off_hours_widget = fixed_map["OFF"]
    st.session_state.post_hours_widget = fixed_map["POST"]
    st.session_state.extra_shift_codes_widget = ", ".join(extra_parts)


def sync_week_editor(week: int) -> None:
    """
    Save one week's data_editor dropdown edits immediately.

    This prevents the "clicked once but it snapped back" Streamlit behavior.
    """
    key = f"schedule_editor_week_{week}"

    if "schedule_df" not in st.session_state:
        return

    editor_state = st.session_state.get(key, {})
    df = st.session_state.schedule_df.copy()

    week_indices = df.index[df["Week"] == week].tolist()

    # Current Streamlit format:
    # {"edited_rows": {0: {"Mon": "D"}}}
    edited_rows = editor_state.get("edited_rows", {})

    for row_idx, updates in edited_rows.items():
        row_idx = int(row_idx)

        if row_idx >= len(week_indices):
            continue

        actual_idx = week_indices[row_idx]

        for col, value in updates.items():
            if col in DAYS:
                df.at[actual_idx, col] = clean_code(value)

    # Older Streamlit fallback:
    # {"edited_cells": {"0:Mon": "D"}}
    edited_cells = editor_state.get("edited_cells", {})

    for cell_key, value in edited_cells.items():
        try:
            row_text, col = str(cell_key).split(":", 1)
            row_idx = int(row_text)
        except ValueError:
            continue

        if row_idx >= len(week_indices):
            continue

        actual_idx = week_indices[row_idx]

        if col in DAYS:
            df.at[actual_idx, col] = clean_code(value)

    st.session_state.schedule_df = df


def summarize_schedule(
    schedule_df: pd.DataFrame,
    shift_defs: pd.DataFrame,
    residents: list[str],
    num_weeks: int,
) -> pd.DataFrame:
    shift_codes = shift_defs["Code"].tolist()
    hours_map = dict(zip(shift_defs["Code"], shift_defs["Hours"]))

    long_df = schedule_df.melt(
        id_vars=["Week", "Resident"],
        value_vars=DAYS,
        var_name="Day",
        value_name="Code",
    )

    long_df["Code"] = long_df["Code"].map(clean_code)
    long_df["Hours"] = long_df["Code"].map(hours_map).fillna(0)

    code_counts = (
        pd.crosstab(long_df["Resident"], long_df["Code"])
        .reindex(index=residents, columns=shift_codes, fill_value=0)
        .reset_index()
    )

    summary = code_counts.copy()
    count_cols = [c for c in summary.columns if c != "Resident"]

    summary["Worked Shifts"] = summary[count_cols].apply(
        lambda row: sum(row[code] for code in count_cols if hours_map.get(code, 0) > 0),
        axis=1,
    )

    summary["Total Hours"] = summary[count_cols].apply(
        lambda row: sum(row[code] * hours_map.get(code, 0) for code in count_cols),
        axis=1,
    )

    summary["Avg Hours / Week"] = summary["Total Hours"] / max(num_weeks, 1)
    summary["Avg Shifts / Week"] = summary["Worked Shifts"] / max(num_weeks, 1)
    summary["OFF Days / 7 Days"] = summary["OFF"] / max(num_weeks, 1) if "OFF" in summary.columns else 0
    summary["POST Days / 7 Days"] = summary["POST"] / max(num_weeks, 1) if "POST" in summary.columns else 0

    return summary


def sheet_ref(sheet_name: str) -> str:
    return f"'{sheet_name.replace(chr(39), chr(39) + chr(39))}'"


def count_formula_for_rows(
    sheet_name: str,
    schedule_rows: list[int],
    criteria_ref: str,
) -> str:
    pieces = [
        f"COUNTIF({sheet_ref(sheet_name)}!$B${row}:$H${row},{criteria_ref})"
        for row in schedule_rows
    ]

    return f"=SUM({','.join(pieces)})"


def build_coverage_tables(schedule_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Check coverage rules.

    Rules:
    - Exclude the first Sunday from all coverage checks.
    - Exclude Mon/Tue/Wed/Thu/Fri/Sat after the last Sunday, interpreted as
      Mon-Sat cells in the final displayed week.
    - Weekend day coverage: included Saturdays and Sundays need at least one D.
    - Night coverage: every included day needs at least one N.
    """
    if schedule_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    min_week = int(schedule_df["Week"].min())
    max_week = int(schedule_df["Week"].max())

    detail_rows = []
    miss_rows = []

    for week in sorted(schedule_df["Week"].unique()):
        week_df = schedule_df[schedule_df["Week"] == week]

        for day in DAYS:
            excluded = (
                (int(week) == min_week and day == "Sun")
                or (int(week) == max_week and day in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"])
            )

            day_series = week_df[["Resident", day]].copy()
            day_series[day] = day_series[day].map(clean_code)

            d_residents = day_series.loc[day_series[day] == "D", "Resident"].tolist()
            n_residents = day_series.loc[day_series[day] == "N", "Resident"].tolist()

            needs_weekend_d = day in ["Sun", "Sat"] and not excluded
            needs_night = not excluded

            weekend_day_covered = True if not needs_weekend_d else len(d_residents) >= 1
            night_covered = True if not needs_night else len(n_residents) >= 1

            detail_rows.append(
                {
                    "Week": int(week),
                    "Day": day,
                    "Excluded": "Yes" if excluded else "No",
                    "Needs Weekend D": "Yes" if needs_weekend_d else "No",
                    "Weekend D Covered": "Yes" if weekend_day_covered else "No",
                    "D Resident(s)": ", ".join(d_residents),
                    "Needs Night N": "Yes" if needs_night else "No",
                    "Night N Covered": "Yes" if night_covered else "No",
                    "N Resident(s)": ", ".join(n_residents),
                }
            )

            if needs_weekend_d and not weekend_day_covered:
                miss_rows.append(
                    {
                        "Week": int(week),
                        "Day": day,
                        "Issue": "Weekend day not covered",
                        "Need": "At least 1 D",
                    }
                )

            if needs_night and not night_covered:
                miss_rows.append(
                    {
                        "Week": int(week),
                        "Day": day,
                        "Issue": "Night not covered",
                        "Need": "At least 1 N",
                    }
                )

    detail_df = pd.DataFrame(detail_rows)
    miss_df = pd.DataFrame(miss_rows)

    return detail_df, miss_df


def show_coverage_checks(schedule_df: pd.DataFrame) -> None:
    """Display weekend and night coverage status in the Streamlit app."""
    detail_df, miss_df = build_coverage_tables(schedule_df)

    st.subheader("Coverage Checks")

    if detail_df.empty:
        st.warning("No schedule rows available to check.")
        return

    weekend_misses = (
        miss_df[miss_df["Issue"] == "Weekend day not covered"]
        if not miss_df.empty
        else pd.DataFrame()
    )

    night_misses = (
        miss_df[miss_df["Issue"] == "Night not covered"]
        if not miss_df.empty
        else pd.DataFrame()
    )

    col1, col2 = st.columns(2)

    with col1:
        if weekend_misses.empty:
            st.success("Weekend days covered: every included Sat/Sun has at least 1 D.")
        else:
            st.error(f"Weekend day gaps: {len(weekend_misses)}")

    with col2:
        if night_misses.empty:
            st.success("Nights covered: every included day has at least 1 N.")
        else:
            st.error(f"Night gaps: {len(night_misses)}")

    if not miss_df.empty:
        st.write("Coverage gaps:")
        st.dataframe(miss_df, use_container_width=True, hide_index=True)

    with st.expander("Coverage check details"):
        st.caption(
            "Excluded from checks: first Sunday, plus Mon/Tue/Wed/Thu/Fri/Sat in the final displayed week."
        )
        st.dataframe(detail_df, use_container_width=True, hide_index=True)


def make_excel(
    schedule_df: pd.DataFrame,
    shift_defs: pd.DataFrame,
    residents: list[str],
    num_weeks: int,
    start_date: date,
) -> bytes:
    shift_codes = shift_defs["Code"].tolist()
    n_residents = len(residents)
    block_height = n_residents + 1  # one header row plus resident rows

    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    summary_ws = wb.create_sheet("Summary")
    defs_ws = wb.create_sheet("Shift_Definitions")
    metadata_ws = wb.create_sheet("Metadata")
    metadata_ws.sheet_state = "hidden"

    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    week_fill = PatternFill("solid", fgColor="E2F0D9")
    thin_gray = Side(style="thin", color="B7B7B7")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    # Metadata sheet
    metadata_ws["A1"] = "Start_Date"
    metadata_ws["B1"] = start_date.strftime("%Y-%m-%d")

    # Shift definitions sheet
    defs_ws["A1"] = "Shift Code"
    defs_ws["B1"] = "Hours"

    for cell in defs_ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for i, row in shift_defs.iterrows():
        excel_row = i + 2
        defs_ws.cell(excel_row, 1).value = row["Code"]
        defs_ws.cell(excel_row, 2).value = float(row["Hours"])
        defs_ws.cell(excel_row, 1).border = border
        defs_ws.cell(excel_row, 2).border = border

    defs_ws.column_dimensions["A"].width = 16
    defs_ws.column_dimensions["B"].width = 10

    # Schedule sheet
    ws["A1"] = "Resident Schedule"
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)

    ws["A2"] = (
        f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}. "
        "Edit shift cells directly; formulas update in Excel."
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    dv_formula = f"={sheet_ref('Shift_Definitions')}!$A$2:$A${len(shift_codes) + 1}"
    shift_validation = DataValidation(type="list", formula1=dv_formula, allow_blank=False)

    for week in range(1, num_weeks + 1):
        header_row = 3 + (week - 1) * block_height
        data_start = header_row + 1

        dated_schedule_headers = [
            "Resident",
            *[day_label(day, start_date, week) for day in DAYS],
        ]

        ws.row_dimensions[header_row].height = 30

        for col_idx, header in enumerate(dated_schedule_headers, start=1):
            cell = ws.cell(header_row, col_idx)
            cell.value = header if col_idx != 1 else f"Week {week}"
            cell.font = Font(bold=True)
            cell.fill = week_fill if col_idx == 1 else header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for resident_idx, resident in enumerate(residents):
            row_num = data_start + resident_idx
            row_data = schedule_df[
                (schedule_df["Week"] == week) & (schedule_df["Resident"] == resident)
            ]

            ws.cell(row_num, 1).value = resident
            ws.cell(row_num, 1).font = Font(bold=True)
            ws.cell(row_num, 1).border = border

            for day_idx, day in enumerate(DAYS, start=2):
                val = clean_code(row_data.iloc[0][day]) if not row_data.empty else "OFF"
                ws.cell(row_num, day_idx).value = val
                ws.cell(row_num, day_idx).alignment = Alignment(horizontal="center")
                ws.cell(row_num, day_idx).border = border

        ws.add_data_validation(shift_validation)
        shift_validation.add(f"B{data_start}:H{data_start + n_residents - 1}")

    ws.freeze_panes = "B4"

    for col, width in {
        "A": 18,
        "B": 10,
        "C": 10,
        "D": 10,
        "E": 10,
        "F": 10,
        "G": 10,
        "H": 10,
    }.items():
        ws.column_dimensions[col].width = width

    # Summary sheet
    summary_headers = [
        "Resident",
        *shift_codes,
        "Worked Shifts",
        "Total Hours",
        "Avg Hours / Week",
        "Avg Shifts / Week",
        "OFF Days / 7 Days",
        "POST Days / 7 Days",
    ]

    summary_ws["A1"] = "Shift Counter and Weekly Averages"
    summary_ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    summary_ws["A1"].fill = title_fill
    summary_ws["A1"].alignment = Alignment(horizontal="center")
    summary_ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(summary_headers))

    header_row = 3

    for col_idx, header in enumerate(summary_headers, start=1):
        cell = summary_ws.cell(header_row, col_idx)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    code_start_col = 2
    code_end_col = code_start_col + len(shift_codes) - 1
    worked_col = code_end_col + 1
    hours_col = worked_col + 1
    avg_hours_col = hours_col + 1
    avg_shifts_col = avg_hours_col + 1
    off_col = avg_shifts_col + 1
    post_col = off_col + 1

    for resident_idx, resident in enumerate(residents):
        summary_row = header_row + 1 + resident_idx

        summary_ws.cell(summary_row, 1).value = resident
        summary_ws.cell(summary_row, 1).font = Font(bold=True)
        summary_ws.cell(summary_row, 1).border = border

        schedule_rows = [
            4 + resident_idx + (week_idx * block_height)
            for week_idx in range(num_weeks)
        ]

        for code_idx, _code in enumerate(shift_codes):
            col_idx = code_start_col + code_idx
            col_letter = get_column_letter(col_idx)
            criteria_ref = f"{col_letter}${header_row}"

            summary_ws.cell(summary_row, col_idx).value = count_formula_for_rows(
                "Schedule",
                schedule_rows,
                criteria_ref,
            )
            summary_ws.cell(summary_row, col_idx).border = border
            summary_ws.cell(summary_row, col_idx).alignment = Alignment(horizontal="center")

        count_terms_for_worked = []
        count_terms_for_hours = []

        for code_idx in range(len(shift_codes)):
            count_col = get_column_letter(code_start_col + code_idx)
            def_row = code_idx + 2
            hour_cell = f"{sheet_ref('Shift_Definitions')}!$B${def_row}"
            count_cell = f"{count_col}{summary_row}"

            count_terms_for_worked.append(f"{count_cell}*({hour_cell}>0)")
            count_terms_for_hours.append(f"{count_cell}*{hour_cell}")

        worked_letter = get_column_letter(worked_col)
        hours_letter = get_column_letter(hours_col)

        summary_ws.cell(summary_row, worked_col).value = f"=SUM({','.join(count_terms_for_worked)})"
        summary_ws.cell(summary_row, hours_col).value = f"=SUM({','.join(count_terms_for_hours)})"
        summary_ws.cell(summary_row, avg_hours_col).value = f"={hours_letter}{summary_row}/{num_weeks}"
        summary_ws.cell(summary_row, avg_shifts_col).value = f"={worked_letter}{summary_row}/{num_weeks}"

        code_header_range = (
            f"${get_column_letter(code_start_col)}${header_row}:"
            f"${get_column_letter(code_end_col)}${header_row}"
        )
        code_count_range = (
            f"{get_column_letter(code_start_col)}{summary_row}:"
            f"{get_column_letter(code_end_col)}{summary_row}"
        )

        summary_ws.cell(summary_row, off_col).value = (
            f'=IFERROR(INDEX({code_count_range},1,MATCH("OFF",{code_header_range},0))/{num_weeks},0)'
        )
        summary_ws.cell(summary_row, post_col).value = (
            f'=IFERROR(INDEX({code_count_range},1,MATCH("POST",{code_header_range},0))/{num_weeks},0)'
        )

        for col_idx in range(worked_col, post_col + 1):
            summary_ws.cell(summary_row, col_idx).border = border
            summary_ws.cell(summary_row, col_idx).alignment = Alignment(horizontal="center")

            if col_idx in [avg_hours_col, avg_shifts_col, off_col, post_col]:
                summary_ws.cell(summary_row, col_idx).number_format = "0.00"

    summary_ws.freeze_panes = "B4"

    for col_idx in range(1, len(summary_headers) + 1):
        col_letter = get_column_letter(col_idx)
        summary_ws.column_dimensions[col_letter].width = max(
            12,
            min(22, len(str(summary_headers[col_idx - 1])) + 2),
        )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Resident Schedule Builder", layout="wide")

    st.title("Resident Schedule Builder")

    st.sidebar.header("Setup")

    if "num_residents_widget" not in st.session_state:
        st.session_state.num_residents_widget = 3

    if "num_weeks_widget" not in st.session_state:
        st.session_state.num_weeks_widget = 5

    if "start_date_widget" not in st.session_state:
        st.session_state.start_date_widget = get_default_monday()

    uploaded_schedule = st.sidebar.file_uploader(
        "Upload existing Excel schedule",
        type=["xlsx"],
        help="Upload a workbook created by this app. It will read the Schedule tab and load it into the editor.",
    )

    if st.sidebar.button(
        "Load Excel into editor",
        disabled=uploaded_schedule is None,
        use_container_width=True,
    ):
        try:
            (
                imported_schedule_df,
                imported_residents,
                imported_num_weeks,
                imported_shift_defs,
                imported_start_date,
            ) = import_schedule_from_excel(uploaded_schedule)

            st.session_state.schedule_df = imported_schedule_df
            st.session_state.num_residents_widget = len(imported_residents)
            st.session_state.num_weeks_widget = imported_num_weeks

            if imported_start_date is not None:
                st.session_state.start_date_widget = imported_start_date

            for i, resident in enumerate(imported_residents):
                st.session_state[f"resident_{i}"] = resident

            # Clear old resident names beyond the imported count.
            for i in range(len(imported_residents), 20):
                key = f"resident_{i}"
                if key in st.session_state:
                    del st.session_state[key]

            apply_imported_shift_defs_to_widgets(imported_shift_defs)

            st.session_state.schedule_signature = None
            clear_schedule_editor_state()

            st.session_state.last_import_message = (
                f"Loaded {len(imported_residents)} residents across {imported_num_weeks} weeks."
            )

            st.rerun()

        except Exception as exc:
            st.sidebar.error(f"Could not import that workbook: {exc}")

    if "last_import_message" in st.session_state:
        st.sidebar.success(st.session_state.last_import_message)

    num_residents = int(
        st.sidebar.number_input(
            "Number of residents",
            min_value=1,
            max_value=20,
            key="num_residents_widget",
            step=1,
        )
    )

    num_weeks = int(
        st.sidebar.number_input(
            "Number of weeks",
            min_value=1,
            max_value=26,
            key="num_weeks_widget",
            step=1,
        )
    )

    start_date = st.sidebar.date_input(
        "Start date, Monday",
        key="start_date_widget",
        help="Enter the Monday start date. The first Sunday shown will be the day before this.",
    )

    if start_date.weekday() != 0:
        st.sidebar.warning("Start date should be a Monday.")

    residents = make_resident_names(num_residents)
    shift_defs = make_shift_definitions()

    # Keep any already-used schedule codes available as dropdown choices.
    if "schedule_df" in st.session_state:
        shift_defs = add_missing_schedule_codes_to_shift_defs(shift_defs, st.session_state.schedule_df)

    shift_codes = shift_defs["Code"].tolist()
    default_code = "OFF" if "OFF" in shift_codes else shift_codes[0]

    signature = (tuple(residents), num_weeks, tuple(shift_codes))

    if "schedule_df" not in st.session_state:
        st.session_state.schedule_df = build_blank_schedule(
            residents,
            num_weeks,
            default_code,
        )
        st.session_state.schedule_signature = signature

    if st.session_state.get("schedule_signature") != signature:
        st.session_state.schedule_df = reconcile_schedule(
            st.session_state.schedule_df,
            residents,
            num_weeks,
            default_code,
        )
        st.session_state.schedule_signature = signature
        clear_schedule_editor_state()

    disabled_cols = ["Resident"]

    for week in range(1, num_weeks + 1):
        st.markdown(f"### Week {week}")

        week_df = st.session_state.schedule_df[
            st.session_state.schedule_df["Week"] == week
        ][["Resident", *DAYS]].reset_index(drop=True)

        column_config = {
            day: st.column_config.SelectboxColumn(
                day_label(day, start_date, week),
                options=shift_codes,
                required=True,
            )
            for day in DAYS
        }

        st.data_editor(
            week_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            disabled=disabled_cols,
            column_config=column_config,
            key=f"schedule_editor_week_{week}",
            on_change=sync_week_editor,
            args=(week,),
            height=42 * (len(residents) + 1) + 20,
        )

    schedule_df = st.session_state.schedule_df.copy()

    for day in DAYS:
        schedule_df[day] = schedule_df[day].map(clean_code)

    st.subheader("Overall Shift Counter")
    summary_df = summarize_schedule(schedule_df, shift_defs, residents, num_weeks)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    show_coverage_checks(schedule_df)

    st.subheader("Export")
    excel_bytes = make_excel(schedule_df, shift_defs, residents, num_weeks, start_date)

    st.download_button(
        label="Download Excel schedule with formulas",
        data=excel_bytes,
        file_name="resident_schedule_with_formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Excel formula logic"):
        st.markdown(
            """
            The exported workbook has:

            - **Schedule** tab: editable schedule cells with dropdowns and dated headers.
            - **Summary** tab: formulas that count each shift code across repeated weekly rows.
            - **Shift_Definitions** tab: shift codes and hours.
            - **Metadata** tab: hidden start date used when you upload the file back into the app.

            The summary formulas are generated automatically and follow the same idea as:

            `=SUM(COUNTIF($B$4:$H$4,"D"),COUNTIF($B$8:$H$8,"D"),COUNTIF($B$12:$H$12,"D"))`
            """
        )


if __name__ == "__main__":
    main()
