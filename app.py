import os
import tempfile
import streamlit as st
import pandas as pd
import openpyxl
from xhtml2pdf import pisa

# --- Page Config ---
st.set_page_config(
    page_title="Seating Plan & Memo Generator",
    page_icon="🎓",
    layout="wide"
)

# --- Visitor Counter Logic ---
COUNTER_FILE = "visitor_count.txt"

def get_and_increment_visitor_count():
    """Increments and returns the total visitor count persistently using a local file."""
    count = 1
    try:
        if os.path.exists(COUNTER_FILE):
            with open(COUNTER_FILE, "r") as f:
                content = f.read().strip()
                if content.isdigit():
                    count = int(content)
            
            if 'has_counted' not in st.session_state:
                count += 1
                with open(COUNTER_FILE, "w") as f:
                    f.write(str(count))
                st.session_state['has_counted'] = True
        else:
            with open(COUNTER_FILE, "w") as f:
                f.write(str(count))
            st.session_state['has_counted'] = True
    except Exception:
        pass
    return count

# --- Backend Logic Functions ---

def is_student_roll(val):
    if val is None:
        return False
    val_str = str(val).strip()
    if any(char.isalpha() for char in val_str):
        return False
    return True

def is_student_absent(absent_val):
    if absent_val is None or pd.isna(absent_val):
        return False
    return str(absent_val).strip().lower() in ("yes", "y", "YES", "Yes", 'Y')

def get_sorted_rolls(roll_list):
    if not roll_list:
        return []
    cleaned = [str(r).strip() for r in roll_list if r is not None and str(r).strip() != "" and str(r).strip().lower() != "nan"]
    if not cleaned:
        return []
    try:
        sorted_rolls = sorted([int(r) for r in cleaned if r.isdigit()])
        str_rolls = sorted([r for r in cleaned if not r.isdigit()])
        return [str(r) for r in sorted_rolls] + str_rolls
    except Exception:
        return sorted(cleaned)

def format_absent_rolls(roll_list):
    sorted_list = get_sorted_rolls(roll_list)
    if not sorted_list:
        return "None"
    return ", ".join(sorted_list)

def allocate_room_columns(streams, num_cols, max_bench, header_sep="\n"):
    room_layout = {}
    room_class_counts = {}
    last_col_last_class = None
    col_idx_counter = 1

    while len(room_layout) < num_cols:
        active_streams = [s for s in streams if s['pointer'] < len(s['rolls'])]
        if not active_streams:
            break

        col_label = f"Row {col_idx_counter}"
        prev_col_label = f"Row {col_idx_counter - 1}"
        prev_col_cells = room_layout.get(prev_col_label, [])

        valid_streams = [s for s in active_streams if s['class'] != last_col_last_class]

        if not valid_streams:
            buf_label = f"Row {col_idx_counter} (Buffer)"
            room_layout[buf_label] = ["Vacant Buffer"] + [None] * max_bench
            last_col_last_class = "Buffer"
            col_idx_counter += 1
            continue

        chosen = valid_streams[0]
        hdr_text = f"{chosen['class']}{header_sep}({chosen['subject']})"
        col_cells = [hdr_text]

        chunk = chosen['rolls'][chosen['pointer'] : chosen['pointer'] + max_bench]
        chosen['pointer'] += len(chunk)
        col_cells.extend(chunk)

        key = (chosen['class'], chosen['subject'])
        room_class_counts[key] = room_class_counts.get(key, 0) + len(chunk)
        current_col_last_class = chosen['class']

        while len(col_cells) < max_bench:
            space_left = (max_bench + 1) - len(col_cells)
            if space_left < 2:
                break

            adj_classes = set()
            for r_pos in range(len(col_cells), max_bench + 1):
                if r_pos < len(prev_col_cells) and prev_col_cells[r_pos] is not None:
                    for back_idx in range(r_pos, -1, -1):
                        cell_val = str(prev_col_cells[back_idx])
                        if not is_student_roll(cell_val) and cell_val != "Vacant Buffer":
                            adj_classes.add(cell_val.split("\n")[0].split(" (")[0].strip())
                            break

            next_valid = [
                s for s in streams
                if s['pointer'] < len(s['rolls']) and s['class'] != current_col_last_class
            ]
            if not next_valid:
                break

            st_next = None
            for s in next_valid:
                if s['class'] not in adj_classes:
                    st_next = s
                    break
            if not st_next:
                st_next = next_valid[0]

            sub_hdr_text = f"{st_next['class']}{header_sep}({st_next['subject']})"
            col_cells.append(sub_hdr_text)

            next_chunk = st_next['rolls'][st_next['pointer'] : st_next['pointer'] + (space_left - 1)]
            st_next['pointer'] += len(next_chunk)
            col_cells.extend(next_chunk)

            key_next = (st_next['class'], st_next['subject'])
            room_class_counts[key_next] = room_class_counts.get(key_next, 0) + len(next_chunk)
            current_col_last_class = st_next['class']

        room_layout[col_label] = col_cells
        last_col_last_class = current_col_last_class
        col_idx_counter += 1

    return room_layout, room_class_counts

def generate_seating_plan_logic(input_path, rooms_path, output_path, exam_name, exam_center, exam_date):
    college_name = "SGTB Khalsa College, Sri Anandpur Sahib (Ropar)"
    df_students = pd.read_excel(input_path)
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_students['Class'] = df_students['Class'].astype(str).str.strip()
    df_students['Subject'] = df_students['Subject'].astype(str).str.strip()
    if 'Absent' not in df_students.columns:
        df_students['Absent'] = None

    df_rooms = pd.read_excel(rooms_path)
    total_students = len(df_students)
    df_rooms['Room_Capacity'] = df_rooms['Rows'] * df_rooms['Columns']
    total_capacity = df_rooms['Room_Capacity'].sum()

    if total_students > total_capacity:
        raise ValueError(f"INSUFFICIENT CAPACITY: Total student count ({total_students}) exceeds combined room capacity ({total_capacity}).")

    grouped = {}
    for (cls, subj), group in df_students.groupby(['Class', 'Subject']):
        grouped[(cls, subj)] = group['ROLL_NUMBER'].tolist()

    streams = [{'class': cls, 'subject': subj, 'rolls': list(rolls), 'pointer': 0} for (cls, subj), rolls in grouped.items()]
    streams.sort(key=lambda s: len(s['rolls']), reverse=True)
    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    for _, room_info in df_rooms.iterrows():
        active_streams = [s for s in streams if s['pointer'] < len(s['rolls'])]
        if not active_streams:
            break

        room_no = room_info['Room']
        num_cols = int(room_info['Rows'])
        max_bench = int(room_info['Columns'])

        room_layout, room_class_counts = allocate_room_columns(streams, num_cols, max_bench, header_sep="\n")
        if not room_layout:
            continue

        max_len = max(len(v) for v in room_layout.values()) if room_layout else 0
        s_no_list = ["S.No."] + list(range(1, max_len))
        upper_data = {"S.No.": s_no_list}

        for col_name, col_vals in room_layout.items():
            padded = col_vals + [None] * (max_len - len(col_vals))
            upper_data[col_name] = padded

        df_upper = pd.DataFrame(upper_data)
        total_row = ["Total"]
        for col_name in room_layout.keys():
            col_data = upper_data[col_name]
            student_count = sum(1 for x in col_data[1:] if is_student_roll(x))
            total_row.append(student_count)

        df_upper = df_upper.astype(object)
        df_upper.loc[len(df_upper)] = total_row
        total_cols = len(df_upper.columns)

        df_upper.to_excel(writer, sheet_name=f"Room {room_no}", index=False, startrow=5)
        worksheet = writer.sheets[f"Room {room_no}"]

        worksheet.cell(row=1, column=1, value=college_name)
        worksheet.cell(row=2, column=1, value=f"Examination: {exam_name}")
        worksheet.cell(row=3, column=1, value=f"Examination Center: {exam_center}")
        worksheet.cell(row=4, column=1, value=f"Room No: {room_no}    |    Date of Exam: {exam_date}")

        for r in range(1, 5):
            worksheet.merge_cells(start_row=r, start_column=1, end_row=r, end_column=total_cols)

        sub_header_fill = openpyxl.styles.PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        center_align = openpyxl.styles.Alignment(horizontal="center", vertical="center")
        wrap_center_align = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, total_cols + 1):
            worksheet.cell(row=6, column=col_idx).alignment = wrap_center_align
            worksheet.cell(row=7, column=col_idx).alignment = wrap_center_align

        for r_idx in range(8, 8 + max_len - 1):
            for c_idx in range(2, total_cols + 1):
                cell = worksheet.cell(row=r_idx, column=c_idx)
                if cell.value is not None and not is_student_roll(cell.value):
                    cell.fill = sub_header_fill
                cell.alignment = center_align

        summary_start_row = 6 + len(df_upper) + 2
        worksheet.merge_cells(start_row=summary_start_row, start_column=1, end_row=summary_start_row, end_column=5)
        worksheet.cell(row=summary_start_row, column=1, value="--- Class-wise Attendance Summary ---").alignment = center_align

        summary_headers = ["Class", "Subject", "Total Students", "Absent", "Present"]
        for col_idx, header_text in enumerate(summary_headers, start=1):
            worksheet.cell(row=summary_start_row + 1, column=col_idx, value=header_text).alignment = center_align

        curr_row = summary_start_row + 2
        total_room_students = 0
        for (cls, subj), count in room_class_counts.items():
            if count > 0:
                worksheet.cell(row=curr_row, column=1, value=cls)
                worksheet.cell(row=curr_row, column=2, value=subj)
                worksheet.cell(row=curr_row, column=3, value=count)
                total_room_students += count
                curr_row += 1

        worksheet.cell(row=curr_row, column=1, value="Total")
        worksheet.cell(row=curr_row, column=3, value=total_room_students)

        for col in worksheet.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            if col[0].column <= total_cols:
                worksheet.column_dimensions[col_letter].width = 14

    writer.close()
    return True

def generate_seating_plan_pdf_logic(input_path, rooms_path, pdf_output_path, exam_name, exam_center, exam_date):
    college_name = "SGTB Khalsa College, Sri Anandpur Sahib (Ropar)"
    if not exam_name:
        exam_name = "MST-II"
    if not exam_center:
        exam_center = "New Building"
    if not exam_date:
        exam_date = "20-Nov-2026"

    df_students = pd.read_excel(input_path)
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_students['Class'] = df_students['Class'].astype(str).str.strip()
    df_students['Subject'] = df_students['Subject'].astype(str).str.strip()
    if 'Absent' not in df_students.columns:
        df_students['Absent'] = None

    df_rooms = pd.read_excel(rooms_path)
    total_students = len(df_students)
    df_rooms['Room_Capacity'] = df_rooms['Rows'] * df_rooms['Columns']
    total_capacity = df_rooms['Room_Capacity'].sum()

    if total_students > total_capacity:
        raise ValueError(f"INSUFFICIENT CAPACITY: Total student count ({total_students}) exceeds room capacity ({total_capacity}).")

    grouped = {}
    for (cls, subj), group in df_students.groupby(['Class', 'Subject']):
        grouped[(cls, subj)] = group['ROLL_NUMBER'].tolist()

    streams = [{'class': cls, 'subject': subj, 'rolls': list(rolls), 'pointer': 0} for (cls, subj), rolls in grouped.items()]
    streams.sort(key=lambda s: len(s['rolls']), reverse=True)
    pages_html = []

    for _, room_info in df_rooms.iterrows():
        active_streams = [s for s in streams if s['pointer'] < len(s['rolls'])]
        if not active_streams:
            break

        room_no = room_info['Room']
        num_cols = int(room_info['Rows'])
        max_bench = int(room_info['Columns'])

        room_layout, room_class_counts = allocate_room_columns(streams, num_cols, max_bench, header_sep=" ")
        if not room_layout:
            continue

        max_len = max(len(v) for v in room_layout.values()) if room_layout else 0
        col_keys = list(room_layout.keys())
        
        th_r1 = "<th>S.No.</th>"
        for ck in col_keys:
            th_r1 += f"<th>{ck}</th>"

        th_r2 = "<th>-</th>"
        for ck in col_keys:
            c_vals = room_layout[ck]
            hdr_text = c_vals[0] if len(c_vals) > 0 and c_vals[0] is not None else ""
            th_r2 += f"<th>{str(hdr_text).replace(chr(10), '<br>')}</th>"

        table_headers = f"<tr>{th_r1}</tr><tr>{th_r2}</tr>"

        table_rows = ""
        for r_idx in range(1, max_len):
            row_cells = f"<td>{r_idx}</td>"
            for ck in col_keys:
                col_vals = room_layout[ck]
                val = col_vals[r_idx] if r_idx < len(col_vals) and col_vals[r_idx] is not None else ""
                if val != "" and not is_student_roll(val):
                    row_cells += f"<td class='sub-header-cell'>{str(val).replace(chr(10), '<br>')}</td>"
                else:
                    row_cells += f"<td>{val}</td>"
            table_rows += f"<tr>{row_cells}</tr>"

        total_cells = "<td>Total</td>"
        total_room_students = 0
        for ck in col_keys:
            col_data = room_layout[ck]
            student_count = sum(1 for x in col_data[1:] if is_student_roll(x))
            total_room_students += student_count
            total_cells += f"<td><b>{student_count}</b></td>"
        table_rows += f"<tr class='total-tr'>{total_cells}</tr>"

        summary_rows = ""
        for (cls, subj), count in room_class_counts.items():
            if count > 0:
                summary_rows += f"<tr><td style='white-space: nowrap;'>{cls}</td><td style='white-space: nowrap;'>{subj}</td><td>{count}</td><td></td><td></td></tr>"

        page_html = f"""
        <div class="seating-page">
            <div>
                <div class="college-header">
                    <div class="college-title">{college_name}</div>
                    <div class="exam-title">Examination: {exam_name} &nbsp;|&nbsp; Center: {exam_center} &nbsp;|&nbsp; Room No: {room_no} &nbsp;|&nbsp; Date: {exam_date}</div>
                </div>

                <div class="table-container">
                    <table class="seating-table">
                        <thead>
                            {table_headers}
                        </thead>
                        <tbody>
                            {table_rows}
                        </tbody>
                    </table>
                </div>

                <div class="summary-section">
                    <div class="summary-title">Class-wise Attendance Summary</div>
                    <table class="summary-table">
                        <thead>
                            <tr>
                                <th style="width: 35%;">Class</th>
                                <th style="width: 35%;">Subject</th>
                                <th style="width: 10%;">Total</th>
                                <th style="width: 10%;">Absent</th>
                                <th style="width: 10%;">Present</th>
                            </tr>
                        </thead>
                        <tbody>
                            {summary_rows}
                            <tr class="total-tr">
                                <td colspan="2" style="text-align: right; padding-right: 10px;"><b>Total Students:</b></td>
                                <td><b>{total_room_students}</b></td>
                                <td></td>
                                <td></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Absolute Pinned Footer at the Bottom Line -->
            <div class="footer-pinned">
                <table style="width: 100%; border-collapse: collapse; font-size: 15px; font-weight: 700;">
                    <tr>
                        <td style="width: 40%; text-align: left;"><b>Date:</b> {exam_date}</td>
                        <td style="width: 60%; text-align: right;">
                            <div>Signature of Center Superintendent: _______________________</div>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        """
        pages_html.append(page_html)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
    @page {{ size: A4 landscape; margin: 5mm; }}
    body {{ margin: 0; padding: 0; font-family: "Calibri", "Arial", sans-serif; color: #111; background: #fff; }}
    .seating-page {{ position: relative; page-break-after: always; page-break-inside: avoid; height: 189mm; box-sizing: border-box; padding: 2mm; }}
    .seating-page:last-child {{ page-break-after: auto; }}
    .college-header {{ text-align: center; margin-bottom: 1px; }}
    .college-title {{ font-size: 20px; font-weight: 800; }}
    .exam-title {{ font-size: 20px; font-weight: 700; margin-top: 1px; }}
    .table-container {{ margin-bottom: 3px; }}
    /* Font size increased by 1.7x (from 11px to ~18.7px) */
    .seating-table {{ width: 100%; border-collapse: collapse; font-size: 15px; text-align: center; }}
    .summary-table {{ width: 100%; border-collapse: collapse; font-size: 15px; text-align: center; }}
    .seating-table th, .seating-table td, .summary-table th, .summary-table td {{ border: 1px solid #333; padding: 2px 2px; }}
    .seating-table th {{ background-color: #2f855a; color: white; font-weight: 800; font-size: 15px; }}
    .sub-header-cell {{ background-color: #2f855a; color: white; font-weight: 800; font-size: 13px; line-height: 1.05; }}
    .summary-table th {{ background-color: #2f855a; color: white; font-weight: 800; font-size: 15px; }}
    .total-tr {{ background-color: #edf2f7; font-weight: 800; }}
    .summary-section {{ margin-top: 3px; width: 100%; }}
    .summary-title {{ font-size: 15px; font-weight: 800; text-align: center; margin-bottom: 1px; }}
    .footer-pinned {{ position: absolute; bottom: 0mm; left: 0; right: 0; width: 100%; }}
</style>
</head>
<body>
    {"".join(pages_html)}
</body>
</html>
"""
    with open(pdf_output_path, "wb") as pdf_file:
        pisaStatus = pisa.CreatePDF(full_html, dest=pdf_file)
    if pisaStatus.err:
        raise RuntimeError("xhtml2pdf encountered errors generating Seating Plan PDF.")
    return True

def generate_class_memo_logic(input_path, output_path, exam_name, exam_center, exam_date):
    center_start_memo_map = {
        "New Building": 101, "Hostel Building": 201, "Center A": 301, "Center B": 401, "Center C": 501
    }
    memo_counter = center_start_memo_map.get(exam_center, 101)
    df_students = pd.read_excel(input_path)
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_students['Class'] = df_students['Class'].astype(str).str.strip()
    df_students['Subject'] = df_students['Subject'].astype(str).str.strip()
    if 'Absent' not in df_students.columns:
        df_students['Absent'] = None

    wb = openpyxl.Workbook()
    default_sheet = wb.active
    grouped = df_students.groupby(['Class', 'Subject'])

    for (cls, subj), group in grouped:
        memo_no = f"{memo_counter:03d}/{exam_date}"
        memo_counter += 1
        sheet_name = f"{cls[:12]}_{subj[:10]}".replace("/", "-")
        present_rolls = []
        absent_rolls = []

        for _, row in group.iterrows():
            if is_student_absent(row['Absent']):
                absent_rolls.append(row['ROLL_NUMBER'])
            else:
                present_rolls.append(row['ROLL_NUMBER'])

        worksheet = wb.create_sheet(title=sheet_name)
        college_title = "SGTB Khalsa College, Sri Anandpur Sahib (Ropar)"

        worksheet.cell(row=1, column=1, value=college_title)
        worksheet.cell(row=2, column=1, value=f"Memo No: {memo_no}       Date: {exam_date}")
        worksheet.cell(row=3, column=1, value=f"Examination: {exam_name}")
        worksheet.cell(row=4, column=1, value=f"Class: {cls}")
        worksheet.cell(row=5, column=1, value=f"Subject: {subj}")

        worksheet.cell(row=7, column=1, value="Roll Numbers of Students Whose Answer Books are Sent:")
        curr_r = 8
        sorted_present = get_sorted_rolls(present_rolls)
        if sorted_present:
            for i in range(0, len(sorted_present), 10):
                worksheet.cell(row=curr_r, column=1, value=", ".join(sorted_present[i:i+10]))
                curr_r += 1

        worksheet.cell(row=curr_r, column=1, value=f"Total Present = {len(sorted_present)}")
        curr_r += 2
        worksheet.cell(row=curr_r, column=1, value="Absent Roll Numbers:")
        curr_r += 1

        sorted_absent = get_sorted_rolls(absent_rolls)
        if sorted_absent:
            for i in range(0, len(sorted_absent), 10):
                worksheet.cell(row=curr_r, column=1, value=", ".join(sorted_absent[i:i+10]))
                curr_r += 1
        else:
            worksheet.cell(row=curr_r, column=1, value="None")

        worksheet.column_dimensions['A'].width = 95

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    wb.save(output_path)
    return True

def generate_class_memo_pdf_logic(input_path, pdf_output_path, exam_name, exam_center, exam_date):
    center_start_memo_map = {
        "New Building": 101, "Hostel Building": 201, "Center A": 301, "Center B": 401, "Center C": 501
    }
    memo_counter = center_start_memo_map.get(exam_center, 101)
    if not exam_name:
        exam_name = "MST-II"
    if not exam_date:
        exam_date = "20-Nov-2026"

    df_students = pd.read_excel(input_path)
    df_students.columns = [str(c).strip() for c in df_students.columns]
    df_students['Class'] = df_students['Class'].astype(str).str.strip()
    df_students['Subject'] = df_students['Subject'].astype(str).str.strip()
    if 'Absent' not in df_students.columns:
        df_students['Absent'] = None

    grouped = df_students.groupby(['Class', 'Subject'])
    pages_html = []

    for (cls, subj), group in grouped:
        memo_no = f"{memo_counter:03d}/{exam_date}"
        memo_counter += 1

        present_rolls = []
        absent_rolls = []

        for _, row in group.iterrows():
            if is_student_absent(row['Absent']):
                absent_rolls.append(row['ROLL_NUMBER'])
            else:
                present_rolls.append(row['ROLL_NUMBER'])

        sorted_present = get_sorted_rolls(present_rolls)
        absent_text = format_absent_rolls(absent_rolls)
        total_present = len(sorted_present)
        total_absent = len(absent_rolls)

        present_chunks = [sorted_present[i:i + 10] for i in range(0, len(sorted_present), 10)]
        if not present_chunks:
            present_chunks = [["-"]]

        present_lines_html = ""
        for chunk in present_chunks:
            line_str = ", ".join(chunk)
            present_lines_html += f'<div style="margin-bottom: 3px;">{line_str}</div>'

        page_html = f"""
        <div class="memo-page">
            <div class="top-code-row">
                <span class="serial-left">Memo No: <b>{memo_no}</b></span>
            </div>

            <div class="header-section">
                <div class="college-name">SGTB Khalsa College</div>
                <div class="college-sub">Sri Anandpur Sahib (Ropar)</div>
                <div class="memo-badge-wrap">
                    <div class="memo-badge">MEMO</div>
                </div>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px; font-weight: 600;">
                <tr>
                    <td style="width: 42%; text-align: left; padding: 6px 0; font-weight: 700;">Examination:</td>
                    <td style="width: 58%; text-align: center; padding: 6px 0; border-bottom: 1.5px dotted #222; font-family: 'Calibri', sans-serif; font-size: 15px; font-weight: 700; color: #1e293b;">{exam_name}</td>
                </tr>
                <tr>
                    <td style="text-align: left; padding: 6px 0; font-weight: 700;">Class:</td>
                    <td style="text-align: center; padding: 6px 0; border-bottom: 1.5px dotted #222; font-family: 'Calibri', sans-serif; font-size: 15px; font-weight: 700; color: #1e293b;">{cls}</td>
                </tr>
                <tr>
                    <td style="text-align: left; padding: 6px 0; font-weight: 700;">Subject / Paper:</td>
                    <td style="text-align: center; padding: 6px 0; border-bottom: 1.5px dotted #222; font-family: 'Calibri', sans-serif; font-size: 15px; font-weight: 700; color: #1e293b;">{subj}</td>
                </tr>
            </table>

            <div class="present-section">
                <div class="section-prompt">
                    Roll Numbers of Students Whose Answer Books are Sent:
                </div>
                <div class="rolls-box" style="font-family: 'Calibri', sans-serif; font-size: 14px; font-weight: 700; color: #1e293b; padding: 4px 0;">
                    {present_lines_html}
                </div>
                <div class="total-row-text">
                    <b>Total = {total_present}</b>
                </div>
            </div>

            <div class="absent-section">
                <table style="width: 100%; border-collapse: collapse; font-size: 14px; font-weight: 700; margin-top: 10px;">
                    <tr>
                        <td style="width: 42%; text-align: left; padding: 4px 0;">Absent Student Roll Numbers:</td>
                        <td style="width: 58%; text-align: center; padding: 4px 0; border-bottom: 1.5px dotted #222; font-family: 'Calibri', sans-serif; font-size: 15px; font-weight: 700; color: #1e293b;">{absent_text}</td>
                    </tr>
                </table>
                <div class="absent-total-row">
                    <b>Total = {total_absent}</b>
                </div>
            </div>

            <div class="footer-section">
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
                    <tr>
                        <td style="width: 42%; text-align: left; font-weight: 700;">Name of Professor Sent To:</td>
                        <td style="width: 58%; border-bottom: 1.5px dotted #222;"></td>
                    </tr>
                </table>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="width: 40%; text-align: left; font-weight: 700;">
                            Date: <span style="border-bottom: 1.5px dotted #222; padding: 0 15px; font-family: 'Calibri', sans-serif;">{exam_date}</span>
                        </td>
                        <td style="width: 60%; text-align: center;">
                            <div class="sig-space"></div>
                            <div class="sig-label">Signature of Center Superintendent</div>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        """
        pages_html.append(page_html)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
    @page {{ size: A4 portrait; margin: 12mm; }}
    body {{ margin: 0; padding: 0; font-family: "Calibri", "Arial", sans-serif; color: #1a1a1a; background: #ffffff; }}
    .memo-page {{ background-color: #f9d5e2; border: 1.5px solid #c994a8; padding: 18px 26px 22px 26px; box-sizing: border-box; page-break-after: always; min-height: 268mm; position: relative; }}
    .memo-page:last-child {{ page-break-after: auto; }}
    .top-code-row {{ font-size: 12px; font-weight: 700; color: #333; margin-bottom: 4px; text-align: left; }}
    .serial-left {{ font-size: 15px; font-weight: 700; }}
    .header-section {{ text-align: center; margin-top: 2px; }}
    .college-name {{ font-size: 26px; font-weight: 800; color: #111; text-align: center; }}
    .college-sub {{ font-size: 18px; font-weight: 700; margin-top: 2px; text-align: center; }}
    .memo-badge-wrap {{ text-align: center; margin-top: 6px; }}
    .memo-badge {{ display: inline-block; border: 1.8px solid #111; border-radius: 6px; padding: 2px 22px; font-size: 16px; font-weight: 800; background: rgba(255,255,255,0.3); text-align: center; }}
    .present-section {{ margin-top: 15px; }}
    .section-prompt {{ font-size: 14px; font-weight: 700; margin-bottom: 8px; text-align: left; }}
    .total-row-text {{ text-align: right; margin-top: 8px; padding-right: 4px; font-size: 16px; font-weight: 800; color: #0f172a; }}
    .absent-section {{ margin-top: 15px; }}
    .absent-total-row {{ text-align: right; margin-top: 6px; padding-right: 4px; font-size: 15px; font-weight: 800; color: #0f172a; }}
    .footer-section {{ margin-top: 20px; font-size: 14px; font-weight: 700; }}
    .sig-space {{ height: 30px; }}
    .sig-label {{ font-size: 14px; font-weight: 700; text-align: center; }}
</style>
</head>
<body>
    {"".join(pages_html)}
</body>
</html>
"""
    with open(pdf_output_path, "wb") as pdf_file:
        pisaStatus = pisa.CreatePDF(full_html, dest=pdf_file)
    if pisaStatus.err:
        raise RuntimeError("xhtml2pdf encountered errors generating Class Memos PDF.")
    return True

# --- Streamlit Web UI ---
st.title("🎓 Examination Seating Plan & Memo Generator")
st.markdown("Upload your **Student Input File** and **Rooms Configuration File** below to generate seating plans and class-wise memos instantly.")

visitor_count = get_and_increment_visitor_count()

with st.sidebar:
    st.header("⚙️ Examination Settings")
    exam_name = st.text_input("Examination Name", value="MST-II")
    exam_center = st.selectbox("Examination Center", ["New Building", "Hostel Building", "Center A", "Center B", "Center C"])
    exam_date = st.text_input("Date of Examination", value="20-Nov-2026")
    
    st.divider()
    st.markdown(f"👥 **Total App Visitors:** `{visitor_count}`")
    
    st.divider()
    st.info("Files required:\n1. Student Input Excel (`input_data.xlsx`)\n2. Rooms Config Excel (`Rooms.xlsx`)")

col1, col2 = st.columns(2)
with col1:
    uploaded_input = st.file_uploader("Upload Student Input File (.xlsx)", type=["xlsx", "xls"])
with col2:
    uploaded_rooms = st.file_uploader("Upload Rooms Configuration File (.xlsx)", type=["xlsx", "xls"])

if uploaded_input and uploaded_rooms:
    st.success("Files uploaded successfully!")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_input.getvalue())
        input_path = tmp_input.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_rooms:
        tmp_rooms.write(uploaded_rooms.getvalue())
        rooms_path = tmp_rooms.name

    tab1, tab2 = st.tabs(["🪑 Seating Plan Generator", "📋 Class Memo Generator"])

    with tab1:
        st.subheader("Generate Seating Plan (Excel & PDF)")
        col_a, col_b = st.columns(2)
        
        with col_a:
            if st.button("Generate Seating Plan Excel", type="primary"):
                output_seating_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
                try:
                    generate_seating_plan_logic(input_path, rooms_path, output_seating_path, exam_name, exam_center, exam_date)
                    with open(output_seating_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Seating Plan Excel",
                            data=f,
                            file_name="Seating_Plan_Output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Seating Plan Excel generated successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")

        with col_b:
            if st.button("Generate Seating Plan PDF", type="primary"):
                output_seating_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                try:
                    generate_seating_plan_pdf_logic(input_path, rooms_path, output_seating_pdf, exam_name, exam_center, exam_date)
                    with open(output_seating_pdf, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Seating Plan PDF",
                            data=f,
                            file_name="Seating_Plan_Output.pdf",
                            mime="application/pdf"
                        )
                    st.success("Seating Plan PDF generated successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")

    with tab2:
        st.subheader("Generate Class Memos (Excel & PDF)")
        col_c, col_d = st.columns(2)

        with col_c:
            if st.button("Generate Class Memos Excel", type="primary"):
                output_memo_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
                try:
                    generate_class_memo_logic(input_path, output_memo_path, exam_name, exam_center, exam_date)
                    with open(output_memo_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Class Memos Excel",
                            data=f,
                            file_name="Class_Wise_Memos.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Class Memos Excel generated successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")

        with col_d:
            if st.button("Generate Class Memos PDF", type="primary"):
                output_memo_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
                try:
                    generate_class_memo_pdf_logic(input_path, output_memo_pdf, exam_name, exam_center, exam_date)
                    with open(output_memo_pdf, "rb") as f:
                        st.download_button(
                            label="⬇️ Download Class Memos PDF",
                            data=f,
                            file_name="Class_Wise_Memos.pdf",
                            mime="application/pdf"
                        )
                    st.success("Class Memos PDF generated successfully!")
                except Exception as z:
                    st.error(f"Error: {z}")
else:
    st.warning("Please upload both the Student Input file and Rooms configuration file to begin.")
