import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import pyproj
from docx import Document
from docx.shared import Pt
from docx.shared import RGBColor

# Constants
TEMPLATE_FILE = "coordinations_template.docx"
HEBREW_LETTERS = ["א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט", "י", "יא", "יב", "יג", "יד", "טו", "טז", "יז", "יח", "יט", "כ"]
PERSONNEL_TRANSLATIONS = {
    "driver": "נהג",
    "assistant driver": "נהג משנה",
    "passenger": "נוסע",
    "patient": "מטופל",
    "patient's companion": "מלווה",
    "security escort": "נוסע - מאבטח",
    "suno": "בכיר",
    "south to north": "\u200F)מדרום לצפון(\u200F",
    "north to south": "\u200F)מצפון לדרום(\u200F",
    "leaving gaza": "\u200F)יוצא מרצע(\u200F",
    "entering gaza": "\u200F)נכנס לרצע(\u200F"
}
VEHICLE_TRANSLATIONS = {
    "truck": "משאית",
    "bus": "אוטובוס",
    "un": "רכב או\"ם", 
    "lc": "רכב או\"ם",
    "toyota land cruiser": "רכב או\"ם",
    "land cruiser": "רכב או\"ם"  
}

# Utility Functions
def extract_below_target(df, target, max_offset=3):
    """
    Extracts the value located below a given target keyword in the dataframe.
    Handles missing values, extra spaces, and case variations.
    """
    target = target.lower().strip()
    found = False  # Flag to indicate if we've found the target

    for row in range(len(df) - max_offset):  # Avoid out-of-bounds error
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()

            if target in cell_value:  # Allow partial match
                found = True
                for offset in range(1, max_offset + 1):  # Check multiple rows below in case of missing data
                    next_row = row + offset
                    if next_row < len(df):
                        next_cell = df.iloc[next_row, col]
                        if isinstance(next_cell, str) and next_cell.strip():
                            return next_cell.strip()
                        elif not pd.isna(next_cell):  # Handle numeric cases
                            return str(next_cell)

    if found:
        return "[MISSING DATA BELOW TARGET]"
    return "[TARGET NOT FOUND]"

def extract_hours(df):
    """
    Extracts start and end times from the table.
    Ensures correct parsing and formatting.
    """
    start_time, end_time = None, None
    time_values = []

    df = df.astype(str)

    for row in range(len(df) - 1):
        for col in range(len(df.columns)):
            cell_value = df.iloc[row, col].strip().lower()

            if "hrs/time" in cell_value:  # Ensure broader keyword match
                for r in range(row + 1, len(df)):  # Scan below the header
                    next_cell = df.iloc[r, col].strip()

                    if next_cell:  
                        try:
                            extracted_time = datetime.strptime(next_cell, "%H:%M:%S").time()
                        except ValueError:
                            try:
                                extracted_time = datetime.strptime(next_cell, "%H:%M").time()
                            except ValueError:
                                continue  # Skip if it's not a valid time format
                    
                        formatted_time = extracted_time.strftime("%H:%M")
                        time_values.append(formatted_time)

                if time_values:
                    start_time = time_values[0]
                    end_time = time_values[-1] 

                    try:
                        end_time_dt = datetime.strptime(end_time, "%H:%M") + timedelta(hours=2)
                        end_time = end_time_dt.strftime("%H:%M")
                    except ValueError:
                        end_time = "[INVALID TIME FORMAT]"

                return start_time or "[MISSING TIME]", end_time or "[MISSING TIME]"

    return "[MISSING TIME]", "[MISSING TIME]"

def format_and_validate_date(date_value):
    """Ensures that the extracted date is formatted correctly as dd.MM.yyyy"""
    date_formats = ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]

    if isinstance(date_value, float):
        extracted_date = datetime(1899, 12, 30) + timedelta(days=int(date_value))  # Excel numeric date conversion
    elif isinstance(date_value, datetime):  # Already a datetime object, just format it
        extracted_date = date_value
    elif isinstance(date_value, str):
        for fmt in date_formats:
            try:
                extracted_date = datetime.strptime(date_value, fmt)
                break
            except ValueError:
                continue
        else:
            return "[INVALID DATE]"
    else:
        return "[INVALID DATE]"
    
    return extracted_date.strftime("%d.%m.%Y")

# Define UTM conversion
transformer = pyproj.Transformer.from_crs("EPSG:4326", "EPSG:32636", always_xy=True)

def latlon_to_utm(lat, lon):
    """Converts latitude/longitude to UTM and shortens the northing value."""
    try:
        easting, northing = transformer.transform(lon, lat)
        short_northing = str(int(northing))[1:]  # Remove the leading '3'
        return f"{int(easting)}/{short_northing}"
    except Exception:
        return "[INVALID COORDS]"

def count_coordinates_levels(df):
    """Extracts location names and coordinates, ensuring correct RTL order."""
    location_col, coordinates_col = None, None
    start_row = None
    empty_row_count = 0
    max_empty_rows_allowed = 1  # Adjust if needed

    # Identify the header row dynamically
    for row in range(len(df)):  
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if "location name" in cell_value:
                location_col = col
                start_row = row + 1
            elif "coordinates" in cell_value:
                coordinates_col = col
                start_row = row + 1

    if location_col is None or coordinates_col is None or start_row is None:
        return ["[MISSING LOCATION/COORDINATES]"]

    counter = 0
    extracted_levels = []

    for row in range(start_row, len(df)):  
        location = df.iloc[row, location_col] if location_col is not None else None
        coords = df.iloc[row, coordinates_col] if coordinates_col is not None else None

        if isinstance(location, str) and location.strip().lower() == "personnel":
            break  # Stop at personnel section

        is_empty_row = not (isinstance(location, str) and location.strip()) or not (isinstance(coords, str) and coords.strip())

        if is_empty_row:
            empty_row_count += 1
            if empty_row_count > max_empty_rows_allowed:
                break  # Stop if too many consecutive empty rows appear
        else:
            empty_row_count = 0  # Reset empty row count if valid data appears
            try:
                lat, lon = map(float, coords.split(","))
                utm_coords = latlon_to_utm(lat, lon)
                extracted_levels.append((HEBREW_LETTERS[counter], location.strip(), utm_coords))
                counter += 1
            except ValueError:
                extracted_levels.append((HEBREW_LETTERS[counter], location.strip(), "[INVALID COORDS]"))
                counter += 1

    formatted_levels = []
    for i, (letter, location, utm) in enumerate(extracted_levels):
        if i == 0:
            point_type = "נקודת התחלה"
        elif i == len(extracted_levels) - 1:
            point_type = "נקודת סיום"
        else:
            point_type = "נקודת עצירה"

        formatted_levels.append(f"\u202B{letter}.\u00A0{point_type} -\u00A0\u200F{location}\u200F \u200E{utm}\u200E")

    return formatted_levels

def check_white_wheels(ids):
    """Ensures white wheels check is independent of other data."""
    for id_value in ids:
        if isinstance(id_value, str) and id_value.strip():
            first_digit = id_value.strip()[0]
            if first_digit not in {"3", "4", "8", "9"}:
                return "גלגלים לבנים"
    return ""

def translate_personnel_role(text):
    if not isinstance(text, str) or not text.strip():
        return "[MISSING DATA]"  # IMPORTANT: Capital letters

    original_text = text.strip()  # Save original casing
    text = original_text.lower()  # Use lowercase only for matching
    main_role = ""
    notes = []

    for eng_term, heb_term in PERSONNEL_TRANSLATIONS.items():
        if eng_term in text:
            if "(" in heb_term:  # Direction/note
                notes.append(heb_term)
            else:  # Main role
                main_role = heb_term

    if main_role:
        if notes:
            return f"\u202B{main_role} {' '.join(notes)}\u202C"
        else:
            return f"\u202B{main_role}\u202C"
    else:
        if notes:
            return f"\u202B{' '.join(notes)}\u202C"
        else:
            return original_text  # keep as is if nothing found

def extract_personnel_table_data(df):
    """
    Extracts personnel table data while ignoring 'Phone #' and flexibly finding 'NOTES (' column.
    """
    personnel_headers = ["name", "id", "notes ("]
    header_indices = {header: None for header in personnel_headers}
    personnel_start_row = None
    empty_row_count = 0
    max_empty_rows_allowed = 1

    # Find the row that contains all personnel headers
    for row in range(len(df)):
        row_values = [str(cell).strip().lower() for cell in df.iloc[row]]
        if all(any(header in cell for cell in row_values) for header in personnel_headers):
            personnel_start_row = row
            for col in range(len(df.columns)):
                cell_value = str(df.iloc[row, col]).strip().lower()
                for header in personnel_headers:
                    if header in cell_value and header_indices[header] is None:
                        header_indices[header] = col
            break

    if personnel_start_row is None or None in header_indices.values():
        return [], "[MISSING PERSONNEL DATA]"

    extracted_data = []
    id_list = []

    for row in range(personnel_start_row + 1, len(df)):
        row_data = {}
        is_empty_row = True

        for key in ["id", "name"]:
            col_index = header_indices[key]
            if col_index is not None:
                cell_value = df.iloc[row, col_index]
                if pd.isna(cell_value) or str(cell_value).strip() == "":
                    cell_value = "[MISSING DATA]"
                elif isinstance(cell_value, float):
                    cell_value = str(int(cell_value))
                else:
                    cell_value = str(cell_value).strip()

                if cell_value != "[MISSING DATA]":
                    is_empty_row = False

                row_data[key] = cell_value
                if key == "id":
                    id_list.append(cell_value)

        notes_col_index = header_indices["notes ("]
        notes_value = df.iloc[row, notes_col_index] if notes_col_index is not None else ""
        if pd.isna(notes_value) or str(notes_value).strip() == "":
            notes_value = "[MISSING DATA]"
        else:
            notes_value = str(notes_value).strip()

        if notes_value != "[MISSING DATA]":
            is_empty_row = False

        row_data["notes"] = notes_value

        if is_empty_row:
            empty_row_count += 1
            if empty_row_count > max_empty_rows_allowed:
                break
        else:
            empty_row_count = 0
            extracted_data.append(row_data)

    white_wheels_text = check_white_wheels(id_list)

    return extracted_data, white_wheels_text

def update_personnel_table(doc_path, extracted_data, white_wheels_text):
    """Updates the personnel table in the Word document with extracted Excel data."""
    doc = Document(doc_path)

    if not doc.tables:
        return "[NO TABLE FOUND IN DOCUMENT]"

    table = doc.tables[0]  # Assuming first table is the personnel table

    # Clear existing rows except for the header
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

    # Populate the table with extracted data
    for index, row_data in enumerate(extracted_data, start=1):
        new_row = table.add_row().cells
        new_row[0].text = f".{index}"  # Numbering
        new_row[1].text = translate_personnel_role(row_data.get("notes", ""))
        new_row[2].text = row_data.get("name", "[MISSING DATA]")  # Name
        new_row[3].text = row_data.get("id", "[MISSING DATA]")  # ID

        for cell in new_row:
            set_font(cell)

        # Apply RTL formatting to numbering (column 0) and role (column 1)
        for col_index in [0, 1]:  
            rtl_paragraph = new_row[col_index].paragraphs[0]
            rtl_paragraph.alignment = 2  # Right-align text
            for run in rtl_paragraph.runs:
                run.font.size = Pt(12)

    # Handle the "גלגלים לבנים" placeholder and remove text box if empty
    if white_wheels_text.strip():
        for para in doc.paragraphs:
            if "{ white_wheels }" in para.text:
                para.text = para.text.replace("{ white_wheels }", white_wheels_text)
    else:
        for shape in doc.element.xpath("//w:drawing"):  # Locate all drawing elements (text boxes)
            parent = shape.getparent()
            if parent is not None:
                parent.remove(shape)  # Remove the entire text box
                break

    doc.save("updated_document.docx")

    return "updated_document.docx"

def translate_vehicle_type(vehicle_type):
    """Ensures Hebrew appears first with a visible space before the English text, handling RTL issues in Word."""
    if not isinstance(vehicle_type, str) or not vehicle_type.strip():
        return "[MISSING DATA]"

    normalized_type = vehicle_type.lower().strip()

    for eng_type, hebrew_translation in VEHICLE_TRANSLATIONS.items():
        if eng_type in normalized_type:
            brand_name = normalized_type.replace(eng_type, "").strip().title()
            if brand_name:
                return f"{brand_name}\u00A0{hebrew_translation}"  # Hebrew first, non-breaking space, then English
            else:
                return f"{hebrew_translation}"  # No brand name, just Hebrew

    return vehicle_type  # If no match, return original input

def extract_vehicle_table_data(df):
    """
    Extracts vehicle type and plate number properly, ensuring no empty rows are added,
    while allowing rows to appear after a blank one.
    """
    headers = ["VEHICLE TYPE", "PLATE #"]
    header_indices = {header.lower(): None for header in headers}

    vehicle_start_row = None
    empty_row_count = 0  # Counter to track consecutive empty rows
    max_empty_rows_allowed = 1  # Adjust this if needed (e.g., allow 2 consecutive empty rows)

    # Identify the header row dynamically
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if cell_value in header_indices:
                header_indices[cell_value] = col
                if vehicle_start_row is None:
                    vehicle_start_row = row

    if None in header_indices.values() or vehicle_start_row is None:
        return [], "[MISSING VEHICLE DATA]"

    extracted_vehicles = []

    for row in range(vehicle_start_row + 1, len(df)):  
        row_data = {}
        is_empty_row = True  # Assume the row is empty

        for key, col_index in header_indices.items():
            cell_value = df.iloc[row, col_index] if col_index is not None else ""

            if pd.isna(cell_value) or str(cell_value).strip() == "":
                cell_value = "[MISSING DATA]"
            elif isinstance(cell_value, float):  # Ensure numbers are formatted as strings
                cell_value = str(int(cell_value))
            else:
                cell_value = str(cell_value).strip()

            if cell_value != "[MISSING DATA]":
                is_empty_row = False  # Row contains valid data

            row_data[key.lower()] = cell_value

        if is_empty_row:
            empty_row_count += 1
            if empty_row_count > max_empty_rows_allowed:
                break  # Stop extraction if too many consecutive empty rows appear
        else:
            empty_row_count = 0  # Reset empty row count if valid data appears
            extracted_vehicles.append(row_data)

    return extracted_vehicles

def update_vehicle_table(doc_path, extracted_vehicles):
    """Updates the vehicle table in the Word document correctly."""
    doc = Document(doc_path)

    if len(doc.tables) < 2:
        return "[SECOND TABLE NOT FOUND]"

    table = doc.tables[1]  # Second table assumed to be vehicle data

    # Clear existing rows except for the header
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

    # Populate the second table with vehicle data
    for index, row_data in enumerate(extracted_vehicles, start=1):
        new_row = table.add_row().cells
        new_row[0].text = f".{index}"  # Numbering
        new_row[1].text = translate_vehicle_type(row_data.get("vehicle type", "[MISSING DATA]"))  # Translated type
        new_row[2].text = row_data.get("plate #", "[MISSING DATA]")  # Plate number

        for cell in new_row:
            set_font(cell)

        # Apply RTL formatting to numbering (column 0) and role (column 1)
        for col_index in [0, 1]:  
            rtl_paragraph = new_row[col_index].paragraphs[0]
            rtl_paragraph.alignment = 2  # Right-align text
            for run in rtl_paragraph.runs:
                run.font.size = Pt(12)

    doc.save("updated_document.docx")

    return "updated_document.docx"

def extract_equipment_status(df):
    """
    Scans rows between 'Equipment' and two rows above 'NOTES' to determine if any cell contains 'yes'.
    Returns 'יש ציוד' if found, else 'אין ציוד'.
    """
    equipment_row, equipment_col = None, None
    notes_row = None

    # First: find 'Equipment' position
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if "equipment" in cell_value:
                equipment_row = row
                equipment_col = col
                break
        if equipment_row is not None:
            break

    # Second: find 'NOTES' position
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if "notes (" in cell_value:
                notes_row = row
                break
        if notes_row is not None:
            break

    # Validate we found both
    if equipment_row is None or notes_row is None:
        return "[MISSING EQUIPMENT OR NOTES]"

    # Scan between equipment_row+1 and notes_row-2
    for row in range(equipment_row + 1, notes_row - 1):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if cell_value == "yes":
                return "יש ציוד"

    # If no 'yes' found
    return "אין ציוד"

# Function to set font properties in Word tables
def set_font(cell, font_name="David", font_size=12):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)

def highlight_text(doc, text, highlight_color):
    """Highlights text in the document with the given color."""
    for para in doc.paragraphs:
        if text in para.text:
            for run in para.runs:
                if text in run.text:
                    run.font.highlight_color = highlight_color

def update_highlighting(doc_path, start_date, equipment_list):
    """Updates highlighting for specific conditions in the Word document."""
    doc = Document(doc_path)
    
    # Check if the date is tomorrow
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%d.%m.%Y")
    is_valid_date = start_date == tomorrow
    
    # Highlight start_date and end_date in red if not tomorrow's date
    if not is_valid_date:
        highlight_text(doc, start_date, 6)  # Red highlight
    
    # Highlight equipment status in yellow if it's "yes"
    if equipment_list == "יש":
        highlight_text(doc, "יש", 7)  # Yellow highlight
    
    # Highlight missing data errors in red
    missing_data_errors = ["[MISSING DATA]", "[MISSING TIME]", "[MISSING LOCATION/COORDINATES]", "[MISSING VEHICLE DATA]", "[TARGET NOT FOUND]", "[INVALID DATE]", "[INVALID COORDS]"]
    
    for para in doc.paragraphs:
        for error in missing_data_errors:
            if error in para.text:
                for run in para.runs:
                    if error in run.text:
                        run.font.highlight_color = 6  # Red highlight
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for error in missing_data_errors:
                        if error in para.text:
                            for run in para.runs:
                                if error in run.text:
                                    run.font.highlight_color = 6  # Red highlight
    
    doc.save("highlighted_document.docx")
    return "highlighted_document.docx"

# Streamlit UI Setup
st.title("Coordinations App")
uploaded_excel = st.file_uploader("Upload your Excel file (.xlsx, .xlsm)", type=["xlsx", "xlsm"])
user_name = st.text_input("Enter your name:")

if uploaded_excel and user_name:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl", header=None)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    st.dataframe(df)

    try:
        doc = DocxTemplate(TEMPLATE_FILE)
    except Exception as e:
        st.error(f"Error loading Word template: {e}")
        st.stop()

    organization_name = extract_below_target(df, "ORGANIZATION")
    raw_start_date = extract_below_target(df, "Date")
    coordination_header = extract_below_target(df, "Purpose")
    start_time, end_time = extract_hours(df)
    start_date = format_and_validate_date(raw_start_date)
    formatted_levels = "\n".join(count_coordinates_levels(df))
    mission = extract_below_target(df, "Execution")
    equipment_list = extract_equipment_status(df)
    table_data, white_wheels_text = extract_personnel_table_data(df)
    vehicle_table_data = extract_vehicle_table_data(df)

    data_dict = {
        "organization_name": organization_name,
        "start_date": start_date,
        "end_date": start_date,
        "start_time": start_time,
        "end_time": end_time,
        "levels": formatted_levels,
        "white_wheels": white_wheels_text,
        "soldier_name": user_name,
        "coordination_header": coordination_header,
        "mission": mission,
        "equipment_list": equipment_list,
    }
    doc.render(data_dict)

    doc.save("processed_document.docx")

    final_doc = update_personnel_table("processed_document.docx", table_data, white_wheels_text)
    final_doc = update_vehicle_table(final_doc, vehicle_table_data)
    final_doc = update_highlighting(final_doc, start_date, equipment_list)

    with open(final_doc, "rb") as f:
        output_stream = BytesIO(f.read())

    st.download_button(
        label="Download Processed Word Document",
        data=output_stream,
        file_name="processed_document.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
