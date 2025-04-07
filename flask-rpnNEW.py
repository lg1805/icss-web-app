from flask import Flask, request, render_template, send_file
import pandas as pd
import os
from datetime import datetime
import xlsxwriter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/processed/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Paths for RPN file
RPN_FILE = r"D:\Lakshya\Project\ICSS_VSCode\PROJECT\ProcessedData\RPN.xlsx"

# Load RPN Data
rpn_data = pd.read_excel(RPN_FILE)
known_components = rpn_data["Component"].dropna().unique().tolist()

def extract_component(observation):
    if pd.notna(observation):
        for component in known_components:
            if str(component).lower() in observation.lower():
                return component
    return "Unknown"

def get_rpn_values(component):
    row = rpn_data[rpn_data["Component"] == component]
    if not row.empty:
        severity = int(row["Severity (S)"].values[0])
        occurrence = int(row["Occurrence (O)"].values[0])
        detection = int(row["Detection (D)"].values[0])
        return severity, occurrence, detection
    return 1, 1, 10

def determine_priority(rpn):
    if rpn >= 200:
        return "High"
    elif rpn >= 100:
        return "Moderate"
    else:
        return "Low"

def format_creation_date(date_str, month_hint):
    try:
        if pd.isna(date_str):
            return None, None

        dt = pd.to_datetime(str(date_str), dayfirst=True, errors='coerce')

        if pd.isna(dt):
            return None, None

        if month_hint.lower() == 'jan':
            day = dt.day
            year = dt.year

            parts = str(date_str).replace('-', '/').split('/')
            if len(parts) == 3:
                parsed_dd = int(parts[0])
                parsed_mm = int(parts[1])
                if parsed_dd == 1 and 1 <= parsed_mm <= 12:
                    day = parsed_mm

            fixed_date = datetime(year, 1, day)
            return fixed_date.strftime('%d/%m/%Y'), (datetime.now() - fixed_date).days

        return dt.strftime('%d/%m/%Y'), (datetime.now() - dt).days

    except Exception:
        return None, None

@app.route('/')
def index():
    return render_template('frontNEW.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'complaint_file' not in request.files:
        return "No complaint_file part", 400

    file = request.files['complaint_file']
    if file.filename == '':
        return "No selected file", 400

    month_hint = request.form.get('month_hint', 'default')

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
        except Exception as e:
            return f"Error reading file: {e}", 400

        if 'Observation' not in df.columns or 'Creation Date' not in df.columns or 'Incident no' not in df.columns:
            return "Error: Required columns are missing in the uploaded file.", 400

        formatted_dates = df['Creation Date'].apply(lambda x: format_creation_date(x, month_hint))
        df['Creation Date'] = formatted_dates.apply(lambda x: x[0])
        days_elapsed = formatted_dates.apply(lambda x: x[1])

        def get_color(elapsed):
            if elapsed == 1:
                return '#ADD8E6'
            elif elapsed == 2:
                return '#FFFF00'
            elif elapsed == 3:
                return '#FF1493'
            elif elapsed > 3:
                return '#FF0000'
            else:
                return None

        df["Component"] = df["Observation"].apply(extract_component)
        df[["Severity (S)", "Occurrence (O)", "Detection (D)"]] = df["Component"].apply(lambda comp: pd.Series(get_rpn_values(comp)))
        df["RPN"] = df["Severity (S)"] * df["Occurrence (O)"] * df["Detection (D)"]
        df["Priority"] = df["RPN"].apply(determine_priority)

        spn_df = df[df["Observation"].str.contains("spn", case=False, na=False)]
        non_spn_df = df[~df["Observation"].str.contains("spn", case=False, na=False)]

        priority_order = {"High": 1, "Moderate": 2, "Low": 3}
        spn_df = spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))
        non_spn_df = non_spn_df.sort_values(by="Priority", key=lambda x: x.map(priority_order))

        processed_filepath = os.path.join(UPLOAD_FOLDER, 'processed_' + file.filename)

        with pd.ExcelWriter(processed_filepath, engine='xlsxwriter') as writer:
            for sheet_name, sheet_df in zip(["SPN", "Non-SPN"], [spn_df, non_spn_df]):
                sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                for idx, row_idx in enumerate(sheet_df.index):
                    elapsed = days_elapsed.loc[row_idx]
                    color = get_color(elapsed)
                    if color:
                        fmt = workbook.add_format({'bg_color': color})
                        worksheet.write(idx + 1, sheet_df.columns.get_loc("Incident no"), sheet_df.iloc[idx]["Incident no"], fmt)

        return send_file(processed_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
