from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from datetime import datetime
import xlsxwriter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Load RPN reference
RPN_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx"
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

# ---------------------
# Utility Functions
# ---------------------

def extract_component(observation):
    if pd.notna(observation):
        for component in known_components:
            if str(component).lower() in str(observation).lower():
                return component
    return "Unknown"

def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        s = int(row["Severity (S)"].values[0])
        o = int(row["Occurrence (O)"].values[0])
        d = int(row["Detection (D)"].values[0])
        return s, o, d
    return 1, 1, 10  # Default fallback

def determine_priority(rpn):
    if rpn >= 200:
        return "High"
    elif rpn >= 100:
        return "Moderate"
    else:
        return "Low"

def format_creation_date_auto(date_str):
    """
    Try parsing a date string as DD/MM/YYYY or MM/DD/YYYY and return the most logical one.
    """
    if pd.isnull(date_str):
        return None, None

    try:
        date_str = str(date_str).strip().replace("-", "/")
        parts = date_str.split("/")
        if len(parts) != 3:
            return None, None

        d1, d2, y = parts
        # Try DD/MM/YYYY
        try:
            dmy = f"{d1.zfill(2)}/{d2.zfill(2)}/{y}"
            date_dmy = datetime.strptime(dmy, "%d/%m/%Y")
            if date_dmy <= datetime.now():
                elapsed = (datetime.now() - date_dmy).days
                return dmy, elapsed
        except:
            pass

        # Try MM/DD/YYYY
        try:
            mdy = f"{d1.zfill(2)}/{d2.zfill(2)}/{y}"
            date_mdy = datetime.strptime(mdy, "%m/%d/%Y")
            if date_mdy <= datetime.now():
                corrected = date_mdy.strftime("%d/%m/%Y")
                elapsed = (datetime.now() - date_mdy).days
                return corrected, elapsed
        except:
            pass

    except:
        return None, None

    return None, None

def get_color(elapsed):
    if elapsed == 1:
        return '#ADD8E6'  # Light Blue
    elif elapsed == 2:
        return '#FFFF00'  # Yellow
    elif elapsed == 3:
        return '#FF1493'  # Deep Pink
    elif elapsed > 3:
        return '#FF0000'  # Red
    return None

# ---------------------
# Flask Routes
# ---------------------

@app.route('/')
def index():
    return render_template('frontNEW2.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No file part", 400

    file = request.files['complaint_file']
    if file.filename == '':
        return "No file selected", 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        return f"Error reading Excel file: {e}", 400

    if 'Observation' not in df.columns or 'Creation Date' not in df.columns or 'Incident no' not in df.columns:
        return "Required columns missing", 400

    # Format date + calculate elapsed
    formatted = df['Creation Date'].apply(lambda x: format_creation_date_auto(x))
    df['Creation Date'] = formatted.apply(lambda x: x[0])
    df['Elapsed Days'] = formatted.apply(lambda x: x[1])

    # Component extraction and RPN logic
    df["Component"] = df["Observation"].apply(extract_component)
    df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = df["Component"].apply(
        lambda c: pd.Series(get_rpn_values(c))
    )
    df["RPN"] = df["Severity (S)"] * df["Occurrence (O)"] * df["Detection (D)"]
    df["Priority"] = df["RPN"].apply(determine_priority)

    # Split by SPN/Non-SPN
    df_spn = df[df["Observation"].str.contains("spn", case=False, na=False)]
    df_nonspn = df[~df["Observation"].str.contains("spn", case=False, na=False)]

    # Split further into Open and Closed
    df_spn_open = df_spn[~df_spn["Incident Status"].str.lower().str.contains("closed|complete", na=False)]
    df_spn_closed = df_spn[df_spn["Incident Status"].str.lower().str.contains("closed|complete", na=False)]
    df_nonspn_open = df_nonspn[~df_nonspn["Incident Status"].str.lower().str.contains("closed|complete", na=False)]
    df_nonspn_closed = df_nonspn[df_nonspn["Incident Status"].str.lower().str.contains("closed|complete", na=False)]

    # Sort by priority
    priority_order = {"High": 1, "Moderate": 2, "Low": 3}
    for df_group in [df_spn_open, df_spn_closed, df_nonspn_open, df_nonspn_closed]:
        df_group.sort_values(by="Priority", key=lambda x: x.map(priority_order), inplace=True)

    processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)

    # Write to Excel with formatting
    with pd.ExcelWriter(processed_filepath, engine='xlsxwriter') as writer:
        grouped = [
            ("SPN Open", df_spn_open),
            ("SPN Closed", df_spn_closed),
            ("Non-SPN Open", df_nonspn_open),
            ("Non-SPN Closed", df_nonspn_closed)
        ]

        for sheet_name, sheet_df in grouped:
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            green_fmt = workbook.add_format({'bg_color': '#C6EFCE'})
            for idx, row_idx in enumerate(sheet_df.index):
                row = sheet_df.loc[row_idx]
                elapsed = row["Elapsed Days"]
                color = get_color(elapsed)
                incident_status = str(row["Incident Status"]).lower()

                # Apply green if Closed/Completed
                if "closed" in incident_status or "complete" in incident_status:
                    worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident Status"), row["Incident Status"], green_fmt)
                elif color:
                    fmt = workbook.add_format({'bg_color': color})
                    worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident no"), row["Incident no"], fmt)

    return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
