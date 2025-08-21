import openpyxl
import json
import os
import re
import io
from difflib import get_close_matches
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from constants import PAGE_METADATA, TABS_METADATA
import importlib.util

# === Google Drive API Setup ===
SERVICE_ACCOUNT_FILE = 'service_account.json'  # Path to your JSON key
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
drive_service = build('drive', 'v3', credentials=credentials)


def normalize(text):
    """Normalize strings for matching (lowercase, remove non-alphanum)."""
    return re.sub(r'[^a-z0-9]', '', str(text).strip().lower())


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


def download_folder_images(folder_id, output_dir, program_type):
    """
    List all files in a Google Drive folder, download them locally,
    upload them to GCS, return public URLs.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Import gcp_access dynamically
    gcp_access_path = os.path.join(script_dir, '..', 'cloud-scripts', 'gcp_access.py')
    spec = importlib.util.spec_from_file_location('gcp_access', gcp_access_path)
    gcp_access = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gcp_access)

    bucket_name = os.environ.get("BUCKET_NAME")

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
                local_filename = os.path.basename(local_file)
                # === Upload to GCS ===
                destination_blob = f"sg-dashboard/partners/{program_type}/{local_filename}"
                folder_url = gcp_access.upload_file_to_gcs_and_get_directory(
                    bucket_name=bucket_name,
                    source_file_path=local_file,
                    destination_blob_name=destination_blob
                )

                if folder_url:
                    os.remove(local_file)  # clean up local copy
                    final_url = f"{folder_url.rstrip('/')}/{local_filename}"
                    logo_urls.append(final_url)
                    print(f"✅ Uploaded {local_filename} → {final_url}")
                else:
                    print(f"❌ Failed to upload {local_filename} to GCS")

        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

    return logo_urls


def build_lookup(state_code_map):
    """
    Build lookup: (norm_state, norm_district) → (state_code, district_code)
    """
    lookup = {}
    district_index = {}
    for state_name, state_info in state_code_map.items():
        state_code = str(state_info.get("id", "")).strip()
        for key, value in state_info.items():
            if key == "id":
                continue
            norm_key = (normalize(state_name), normalize(key))
            district_code = str(value).strip()
            lookup[norm_key] = (state_code, district_code)
            district_index.setdefault(normalize(state_name), {})[normalize(key)] = (state_code, district_code)
    return lookup, district_index


def resolve_codes(state, district, lookup, district_index):
    """Try exact → fuzzy → fail"""
    norm_key = (normalize(state), normalize(district))
    if norm_key in lookup:
        return lookup[norm_key]

    state_key = normalize(state)
    if state_key in district_index:
        possible_districts = list(district_index[state_key].keys())
        close = get_close_matches(normalize(district), possible_districts, n=1, cutoff=0.8)
        if close:
            print(f"ℹ️ Using fuzzy match: {district} → {close[0]}")
            return district_index[state_key][close[0]]

    return None, None


def generate_program_reports(excel_file):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_images_dir = os.path.join(script_dir, 'tmp_images')
        os.makedirs(base_images_dir, exist_ok=True)

        # Load state + district codes
        state_code_path = os.path.join(script_dir, '..', "pages", 'state_code_details.json')
        with open(state_code_path, 'r', encoding='utf-8') as f:
            state_code_map = json.load(f)

        # Build lookup
        district_lookup, district_index = build_lookup(state_code_map)

        # Load Excel
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook[PAGE_METADATA["PROGRAMS"]]

        headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]
        header_index_map = {str(col).strip(): i for i, col in enumerate(headers) if col}

        print("✅ Headers from Excel:", headers)

        slc_data = {}
        wlc_data = {}

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            row_dict = {}
            for col in TABS_METADATA["PROGRAMS"]:
                col = str(col).strip()
                idx = header_index_map.get(col)
                row_dict[col] = row[idx] if idx is not None and idx < len(row) else ''

            state = str(row_dict['State Name']).strip()
            district = str(row_dict['District Name']).strip()
            program = str(row_dict['Name of the Program']).strip()
            program_type = str(row_dict.get('Program Type', '')).strip().upper()

            if not district or district.lower() == 'none':
                district = state
                row_dict['District Name'] = district

            if not (state and district and program):
                continue

            # --- Get codes ---
            state_code, district_code = resolve_codes(state, district, district_lookup, district_index)
            if not district_code:
                print(f"⚠️ District code not found for {district} in {state}")
                continue

            # === Handle Google Drive Folder Links ===
            folder_url = row_dict.get('Pictures from the program', '')
            logo_urls = []
            if folder_url:
                folder_id = extract_folder_id(folder_url)
                if folder_id:
                    program_folder = os.path.join(base_images_dir, f"{program.replace(' ', '_').lower()}")
                    os.makedirs(program_folder, exist_ok=True)
                    logo_urls = download_folder_images(folder_id, program_folder, program_type)

            row_dict['logo_urls'] = logo_urls

            # --- Save using district id ---
            target_dict = wlc_data if program_type == "WLC" else slc_data
            state_dict = target_dict.setdefault(str(state_code), {})
            state_dict.setdefault(str(district_code), {}).setdefault(str(program), []).append(row_dict)

        # === Save SLC.json and WLC.json per state and upload to GCS ===
        states_dir = os.path.join(script_dir, '..', 'states')
        os.makedirs(states_dir, exist_ok=True)

        # Import gcp_access dynamically
        gcp_access_path = os.path.join(script_dir, '..', 'cloud-scripts', 'gcp_access.py')
        spec = importlib.util.spec_from_file_location('gcp_access', gcp_access_path)
        gcp_access = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(gcp_access)

        bucket_name = os.environ.get("BUCKET_NAME")

        for category_name, data_dict in [('SLC', slc_data), ('WLC', wlc_data)]:
            for state_code, districts in data_dict.items():
                state_folder = os.path.join(states_dir, str(state_code))
                os.makedirs(state_folder, exist_ok=True)
                out_file = os.path.join(state_folder, f"{category_name}.json")
                
                # Save locally
                with open(out_file, 'w', encoding='utf-8') as f:
                    json.dump(districts, f, indent=2, ensure_ascii=False)
                print(f"✅ Saved {category_name}.json for state {state_code} at {out_file}")

                # Upload to GCS
                gcs_path = f"sg-dashboard/states/{state_code}/{category_name}.json"
                folder_url = gcp_access.upload_file_to_gcs_and_get_directory(
                    bucket_name=bucket_name,
                    source_file_path=out_file,
                    destination_blob_name=gcs_path
                )
                if folder_url:
                    print(f"✅ Uploaded {category_name}.json for state {state_code} to {folder_url}")
                else:
                    print(f"❌ Failed to upload {category_name}.json for state {state_code}")

        print("✅ Program reports generated and uploaded successfully.")

    except Exception as e:
        print(f"❌ Fatal Error: {e}")


if __name__ == "__main__":
    excel_file = "programs.xlsx"  # replace with your Excel file path
    generate_program_reports(excel_file)
