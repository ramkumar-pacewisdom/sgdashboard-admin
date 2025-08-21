import openpyxl
import json
import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
from constants import PAGE_METADATA, TABS_METADATA

# === Google Drive API Setup ===
SERVICE_ACCOUNT_FILE = 'service_account.json'  # Path to your JSON key
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)


def extract_folder_id(drive_url):
    """Extract folder ID from Google Drive URL"""
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', drive_url)
    if not match:
        match = re.search(r'id=([a-zA-Z0-9_-]+)', drive_url)
    return match.group(1) if match else None


def download_file(file_id, output_dir):
    """Download single file from Google Drive using file ID"""
    try:
        file_metadata = drive_service.files().get(fileId=file_id).execute()
        filename = file_metadata.get('name')
        request = drive_service.files().get_media(fileId=file_id)
        os.makedirs(output_dir, exist_ok=True)
        file_path = os.path.join(output_dir, filename)
        fh = io.FileIO(file_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        return file_path
    except Exception as e:
        print(f"❌ Failed to download file {file_id}: {e}")
        return None


def download_folder_images(folder_id, output_dir):
    """List all files in a Google Drive folder and download them"""
    logo_urls = []
    page_token = None
    while True:
        response = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType contains 'image/'",
            spaces='drive',
            fields='nextPageToken, files(id, name)',
            pageToken=page_token
        ).execute()
        for file in response.get('files', []):
            file_id = file['id']
            local_file = download_file(file_id, output_dir)
            if local_file:
                logo_urls.append(local_file)
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    return logo_urls


def generate_program_reports(excel_file):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_base_dir = os.path.join(script_dir, '..', "pages", 'program-reports')
        base_images_dir = os.path.join(script_dir, 'tmp_images')
        os.makedirs(base_images_dir, exist_ok=True)

        # Load state codes
        state_code_path = os.path.join(script_dir, '..', "pages", 'state_code_details.json')
        with open(state_code_path, 'r', encoding='utf-8') as f:
            state_code_map = {k.strip().lower(): v for k, v in json.load(f).items()}

        # Load Excel
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook[PAGE_METADATA["PROGRAMS"]]
        headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]
        header_index_map = {col: i for i, col in enumerate(headers) if col}

        slc_data = {}
        wlc_data = {}

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            row_dict = {}
            for col in TABS_METADATA["PROGRAMS"]:
                idx = header_index_map.get(col)
                row_dict[col] = row[idx] if idx is not None and idx < len(row) else ''

            state = str(row_dict['State Name']).strip().lower()
            district = str(row_dict['District Name']).strip()
            program = str(row_dict['Name of the Program']).strip()
            program_type = str(row_dict.get('Program Type', '')).strip().upper()

            if not district or district.lower() == 'none':
                district = str(row_dict['State Name']).strip()
                row_dict['District Name'] = district

            if not (state and district and program):
                continue

            state_code = state_code_map.get(state)
            if not state_code:
                continue

            # === Handle Google Drive Folder Links ===
            folder_url = row_dict.get('Pictures from the program', '')
            logo_urls = []

            if folder_url:
                folder_id = extract_folder_id(folder_url)
                if folder_id:
                    program_folder = os.path.join(base_images_dir, f"{program.replace(' ', '_').lower()}")
                    os.makedirs(program_folder, exist_ok=True)
                    logo_urls = download_folder_images(folder_id, program_folder)

            row_dict['logo_urls'] = logo_urls

            # Save in target dict
            target_dict = wlc_data if program_type == "WLC" else slc_data
            target_dict.setdefault(state_code, {}).setdefault(district, {}).setdefault(program, []).append(row_dict)

        # Write JSON
        for category_name, data_dict in [('SLC', slc_data), ('WLC', wlc_data)]:
            category_dir = os.path.join(output_base_dir, category_name)
            os.makedirs(category_dir, exist_ok=True)
            for state_code, districts in data_dict.items():
                state_folder = os.path.join(category_dir, str(state_code))
                os.makedirs(state_folder, exist_ok=True)
                out_file = os.path.join(state_folder, f"{state_code}.json")
                with open(out_file, 'w', encoding='utf-8') as f:
                    json.dump(districts, f, indent=2, ensure_ascii=False)

        print("✅ Program reports generated successfully.")

    except Exception as e:
        print(f"❌ Fatal Error: {e}")
