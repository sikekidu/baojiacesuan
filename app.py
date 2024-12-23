from googleapiclient.discovery import build
from google.oauth2 import service_account
from flask import Flask, request, jsonify, send_file, send_from_directory
import os
import io
from googleapiclient import http
import logging

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

SERVICE_ACCOUNT_FILE = 'plucky-portal-389210-4bb948748fd4.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file']
SPREADSHEET_ID = '1o4c7PUcp7Y5fhLxRISThiywpmsJFrF5bR3ssr8M-hTM'

creds = None
if os.path.exists(SERVICE_ACCOUNT_FILE):
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

service = build('sheets', 'v4', credentials=creds)
sheets = service.spreadsheets()
drive_service = build('drive', 'v3', credentials=creds)

@app.route('/', methods=['GET'])
def index():
    logging.debug("Serving index.html")
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), 'index.html')

@app.route('/process_sheets', methods=['POST'])
def process_sheets():
    try:
        logging.debug("Received POST request to /process_sheets")
        data = request.get_json()
        logging.debug(f"Request data: {data}")
        update_values(data)
        result_data, dropdown_options = get_sheet_data()
        logging.debug(f"Response data: {result_data}")
        return jsonify({'data': result_data, 'dropdown_options': dropdown_options})
    except Exception as e:
        logging.error(f"Error processing request: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/dropdown_options', methods=['GET'])
def dropdown_options():
    try:
        logging.debug("Received GET request to /dropdown_options")
        _, dropdown_options = get_sheet_data()
        logging.debug(f"Dropdown options: {dropdown_options}")
        return jsonify({'dropdown_options': dropdown_options})
    except Exception as e:
        logging.error(f"Error processing request: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download_pdf', methods=['GET'])
def download_pdf():
    try:
        logging.debug("Received GET request to /download_pdf")
        pdf_data = export_to_pdf()
        memory_file = io.BytesIO(pdf_data)
        logging.debug("PDF download successful")
        return send_file(memory_file, as_attachment=True, download_name='sheet_download.pdf', mimetype='application/pdf')
    except Exception as e:
        logging.error(f"Error processing PDF download: {e}")
        return jsonify({'success': False, 'error': str(e)})

def get_sheet_data():
    logging.debug("Fetching data from Google Sheet")

    # Fetch metadata to get the number of rows and columns
    spreadsheet = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
    num_rows = spreadsheet.get('sheets')[0].get('properties').get('gridProperties').get('rowCount')
    num_cols = spreadsheet.get('sheets')[0].get('properties').get('gridProperties').get('columnCount')

    range = f"A1:{chr(64+num_cols)}{num_rows}"
    logging.debug(f"Fetching data from sheet with range {range}")
    result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=range).execute()
    all_data = result.get('values', [])
    logging.debug(f"Data from sheet: {all_data}")

    # Get data validation rule for C4
    dropdown_options = []
    result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="基础物流价格信息!A2:A10").execute()
    values = result.get("values", [])
    dropdown_options = [value[0] for value in values]
    logging.debug(f"Dropdown options: {dropdown_options}")

    return all_data, dropdown_options

def update_values(data):
    logging.debug("Updating Google Sheet")
    values = []
    for key, value in data.items():
        if key == 'C4':  # 特殊处理 C4 下拉框
            value = value
        row_idx = int(key[1:]) - 1
        col_idx = ord(key[0]) - 65
        values.append({'range': key, 'values': [[value]]})
    body = {'value_input_option': 'USER_ENTERED', 'data': values}
    result = sheets.values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    logging.debug("Update operation successful")
    return result

def export_to_pdf():
    logging.debug("Exporting Google Sheet to PDF")
    request = drive_service.files().export_media(fileId=SPREADSHEET_ID, mimeType='application/pdf')
    fh = io.BytesIO()
    downloader = http.MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        logging.debug(f"Download {int(status.progress() * 100)}%.")

    fh.seek(0)
    pdf_data = fh.getvalue()
    logging.debug("PDF data downloaded successfully")
    return pdf_data

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
