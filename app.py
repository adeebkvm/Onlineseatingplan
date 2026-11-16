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
            
            # Increment only once per session so page refreshes don't spam the count
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

def format_present_rolls_chunks(roll_list):
    if not roll_list:
        return [], 0
    cleaned_rolls = []
    for r in roll_list:
        if r is not None and str(r).strip() != "" and str(r).strip().lower() != "nan":
            try:
                cleaned_rolls.append(int(r))
            except ValueError:
                cleaned_rolls.append(str(r).strip())

    if not cleaned_rolls:
        return [], 0

    sorted_rolls = sorted([int(r) for r in cleaned_rolls if str(r).isdigit() or isinstance(r, (int, float))])
    str_rolls = sorted([str(r) for r in cleaned_rolls if not (str(r).isdigit() or isinstance(r, (int, float)))])

    chunks = []
    if sorted_rolls:
        current_chunk = [sorted_rolls[0]]
        for r in sorted_rolls[1:]:
            if r == current_chunk[-1] + 1:
                current_chunk.append(r)
            else:
                chunks.append(current_chunk)
                current_chunk = [r]
        chunks.append(current_chunk)

    formatted_chunks = []
    total_count = len(sorted_rolls) + len(str_rolls)

    for chunk in chunks:
        if len(chunk) == 1:
            range_str = f"{chunk[0]}"
        else:
            start_str = str(chunk[0])
            end_str = str(chunk[-1])
            if start_str[:-2] == end_str[:-2] and len(start_str) > 2:
                short_end = end_str[-2:]
                range_str = f"{start_str}-{short_end}"
            else:
                range_str = f"{start_str}-{end_str}"
        formatted_chunks.append({
            "range": range_str,
            "count": f"{len(chunk):02d}",
            "num_count": len(chunk)
        })

    for s in str_rolls:
        formatted_chunks.append({
            "range": str(s),
            "count": "01",
            "num_count": 1
        })

    return formatted_chunks, total_count

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
    college_name = "ਸ੍ਰੀ ਗੁਰੂ ਤੇਗ ਬਹਾਦਰ ਖਾਲਸਾ ਕਾਲਜ, ਸ੍ਰੀ ਅਨੰਦਪੁਰ  ਸਾਹਿਬ (ਰੋਪੜ)"
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
        college_title = "ਸ੍ਰੀ ਗੁਰੂ ਤੇਗ ਬਹਾਦਰ ਖਾਲਸਾ ਕਾਲਜ, ਸ੍ਰੀ ਅਨੰਦਪੁਰ  ਸਾਹਿਬ (ਰੋਪੜ)"

        worksheet.cell(row=1, column=1, value=college_title)
        worksheet.cell(row=2, column=1, value=f"ਮੀਮੋ ਨੰ: {memo_no}       ਮਿਤੀ / Date: {exam_date}")
        worksheet.cell(row=3, column=1, value=f"ਪ੍ਰੀਖਿਆ / Examination: {exam_name}")
        worksheet.cell(row=4, column=1, value=f"ਕਲਾਸ (Class): {cls}")
        worksheet.cell(row=5, column=1, value=f"ਵਿਸ਼ਾ (Subject): {subj}")

        worksheet.cell(row=7, column=1, value="ਪ੍ਰੀਖਿਆਰਥੀਆਂ ਦੇ ਰੋਲ ਨੰ: ਜਿੰਨ੍ਹਾਂ ਦੀਆਂ ਉੱਤਰ ਕਾਪੀਆਂ ਭੇਜੀਆਂ ਜਾ ਰਹੀਆਂ ਹਨ:")
        curr_r = 8
        sorted_present = get_sorted_rolls(present_rolls)
        if sorted_present:
            for i in range(0, len(sorted_present), 10):
                worksheet.cell(row=curr_r, column=1, value=", ".join(sorted_present[i:i+10]))
                curr_r += 1

        worksheet.cell(row=curr_r, column=1, value=f"ਕੁੱਲ ਹਾਜ਼ਰ (Total Present) = {len(sorted_present)}")
        curr_r += 2
        worksheet.cell(row=curr_r, column=1, value="ਗੈਰ ਹਾਜ਼ਰ ਪ੍ਰੀਖਿਆਰਥੀਆਂ ਦੇ ਰੋਲ ਨੰ:")
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

# --- Streamlit Web UI ---
st.title("🎓 Examination Seating Plan & Memo Generator")
st.markdown("Upload your **Student Input File** and **Rooms Configuration File** below to generate seating plans and class-wise memos instantly.")

# Retrieve visitor count
visitor_count = get_and_increment_visitor_count()

with st.sidebar:
    st.header("⚙️ Examination Settings")
    exam_name = st.text_input("Examination Name", value="MST-II")
    exam_center = st.selectbox("Examination Center", ["New Building", "Hostel Building", "Center A", "Center B", "Center C"])
    exam_date = st.text_input("Date of Examination", value="20-Nov-2026")
    
    st.divider()
    # Display Visitor Counter in Sidebar
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
    
    # Save uploaded files temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
        tmp_input.write(uploaded_input.getvalue())
        input_path = tmp_input.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_rooms:
        tmp_rooms.write(uploaded_rooms.getvalue())
        rooms_path = tmp_rooms.name

    tab1, tab2 = st.tabs(["🪑 Seating Plan Generator", "📋 Class Memo Generator"])

    with tab1:
        st.subheader("Generate Seating Plan (Excel)")
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
                st.success("Seating Plan generated successfully!")
            except Exception as e:
                st.error(f"Error generating seating plan: {e}")

    with tab2:
        st.subheader("Generate Class Memos (Excel)")
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
                st.success("Class Memos generated successfully!")
            except Exception as e:
                st.error(f"Error generating class memos: {e}")
else:
    st.warning("Please upload both the Student Input file and Rooms configuration file to begin.")
