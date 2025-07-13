from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import os, tempfile, re, datetime
from google.cloud import documentai_v1 as documentai
from openpyxl import load_workbook
from copy import copy

# === CONFIG (cloud-ready) ===
PROJECT_ID = os.environ.get('PROJECT_ID', 'driver-schedule-ocr')
LOCATION = os.environ.get('LOCATION', 'us')
PROCESSOR_ID = os.environ.get('PROCESSOR_ID', '2acb7269aa33ccf9')
GOOGLE_KEY_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "/etc/secrets/driver-schedule-ocr.json")
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = GOOGLE_KEY_PATH
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", "Truck_Load_Record_Template.xlsx")
# ============================

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Hello, Flask! The server is running."

def extract_table_from_image(image_path):
    client = documentai.DocumentProcessorServiceClient()
    name = f"projects/{PROJECT_ID}/locations/{LOCATION}/processors/{PROCESSOR_ID}"
    with open(image_path, "rb") as image:
        image_content = image.read()
    raw_document = documentai.RawDocument(content=image_content, mime_type="image/jpeg")
    request = documentai.ProcessRequest(name=name, raw_document=raw_document)
    result = client.process_document(request=request)
    document = result.document

    # Extract tables
    tables = []
    headers = []
    for page in document.pages:
        for table in page.tables:
            header_cells = []
            for cell in table.header_rows[0].cells:
                cell_text = ""
                for segment in cell.layout.text_anchor.text_segments:
                    start = int(segment.start_index)
                    end = int(segment.end_index)
                    cell_text += document.text[start:end]
                header_cells.append(cell_text.strip())
            headers = header_cells
            for row in table.body_rows:
                row_values = []
                for cell in row.cells:
                    cell_text = ""
                    for segment in cell.layout.text_anchor.text_segments:
                        start = int(segment.start_index)
                        end = int(segment.end_index)
                        cell_text += document.text[start:end]
                    row_values.append(cell_text.strip())
                tables.append(row_values)
    if not tables:
        return None
    df = pd.DataFrame(tables, columns=headers)

    # Find columns by partial match
    def find_col(df, substring):
        for col in df.columns:
            if substring.lower() in col.lower():
                return col
        return None

    run_col = find_col(df, 'Run')
    driver1_col = find_col(df, 'Driver 1')
    driver2_col = find_col(df, 'Driver 2')
    truck_col = find_col(df, 'Truck')

    # Cleaners
    def extract_run_numbers(cell):
        nums = re.findall(r'\b\d{4}\b', str(cell))
        return " / ".join(nums)

    def clean_driver(cell):
        if pd.isnull(cell): return ""
        name = str(cell).split('\n')[0]
        return re.sub(r'[^A-Za-z ]+', '', name).strip()

    def clean_truck(cell):
        cell = str(cell).strip()
        return cell[:6]

    df_clean = pd.DataFrame()
    df_clean["Run#"] = df[run_col].apply(extract_run_numbers)
    df_clean["Driver1"] = df[driver1_col].apply(clean_driver)
    df_clean["Driver2"] = df[driver2_col].apply(clean_driver) if driver2_col else ""
    df_clean["Truck"] = df[truck_col].apply(clean_truck)
    return df_clean

def fill_template_per_truck(df_clean):
    # Load the template workbook (must only have ONE sheet)
    template_wb = load_workbook(TEMPLATE_PATH)
    template_sheet = template_wb.active

    # Prepare output workbook
    out_wb = load_workbook(TEMPLATE_PATH)  # Start from a real template copy
    out_wb.remove(out_wb.active)           # Remove template sheet after copying

    today = datetime.date.today() + datetime.timedelta(days=1)

    for idx, truck_row in enumerate(df_clean.to_dict(orient="records")):
        # Create new sheet by copying template sheet
        ws = out_wb.copy_worksheet(template_sheet)
        # Give a unique, valid name for Excel sheets
        sheet_name = (truck_row.get("Truck") or truck_row.get("Run#") or "Sheet")[:31]
        if sheet_name in out_wb.sheetnames:
            sheet_name = f"{sheet_name}_{idx+1}"
        ws.title = sheet_name

        # Fill fields
        ws["B3"] = truck_row.get("Run#", "")
        ws["I3"] = truck_row.get("Truck", "")
        driver1 = truck_row.get("Driver1", "")
        driver2 = truck_row.get("Driver2", "")
        ws["B4"] = " / ".join([d for d in [driver1, driver2] if d])
        ws["I4"] = today.strftime("%d/%m/%Y")

    # Save output workbook
    out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    out_wb.save(out.name)
    out.close()
    return out.name

@app.route('/parse_schedule_excel', methods=['POST'])
def parse_schedule_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp:
        file.save(temp.name)
        df_clean = extract_table_from_image(temp.name)
    if df_clean is None or df_clean.empty:
        return jsonify({"error": "No table detected"}), 400

    filled_path = fill_template_per_truck(df_clean)
    return send_file(filled_path, as_attachment=True, download_name="truck_load_records.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
