import openpyxl
import json
import os
import importlib.util

from constants import PAGE_METADATA, TABS_METADATA

def generate_program_reports(excel_file):
    try:
        # === Setup Paths ===
        script_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"Script running from: {script_dir}")

        state_code_path = os.path.join(script_dir, '..', "pages", 'state_code_details.json')
        output_base_dir = os.path.join(script_dir, '..', "pages", 'program-reports')

        # === Load State Codes ===
        with open(state_code_path, 'r', encoding='utf-8') as f:
            state_code_map = json.load(f)
        state_code_map = {k.strip().lower(): v for k, v in state_code_map.items()}

        # === Load Excel ===
        workbook = openpyxl.load_workbook(excel_file, data_only=True)

        try:
            sheet = workbook[PAGE_METADATA["PROGRAMS"]]
        except KeyError:
            print("❌ Error: Sheet defined in PAGE_METADATA['PROGRAMS'] not found.")
            print(f"Available sheets: {workbook.sheetnames}")
            return

        # === Preserve full header row ===
        full_headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]

        # === Validate required columns ===
        expected_columns = TABS_METADATA["PROGRAMS"]
        if not all(col in full_headers for col in expected_columns):
            print(f"❌ Missing required columns. Expected: {expected_columns}")
            print(f"Found: {full_headers}")
            return

        # === Column index map for safe extraction ===
        header_index_map = {col: i for i, col in enumerate(full_headers) if col}

        # === Prepare Data Structures ===
        slc_data = {}  # Non-WLC
        wlc_data = {}  # Only WLC

        # === Parse Rows ===
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            try:
                row_dict = {}
                for col in expected_columns:
                    idx = header_index_map.get(col)
                    row_dict[col] = row[idx] if idx is not None and idx < len(row) else ''

                state = str(row_dict['State Name']).strip().lower()
                district = str(row_dict['District Name']).strip()
                program = str(row_dict['Name of the Program']).strip()
                program_type = str(row_dict.get('Program Type', '')).strip().upper()

                # ✅ Replace missing/empty district with state name
                if not district or district.lower() == 'none':
                    district = str(row_dict['State Name']).strip()
                    row_dict['District Name'] = district  # <-- ✅ UPDATE the field in data as well

                if not (state and district and program):
                    print(f"⚠️ Skipping row {row_idx}: Missing state/district/program name.")
                    continue

                state_code = state_code_map.get(state)
                if not state_code:
                    print(f"⚠️ Skipping row {row_idx}: Unknown state '{state}'")
                    continue

                # Choose target dict
                target_dict = wlc_data if program_type == "WLC" else slc_data

                if state_code not in target_dict:
                    target_dict[state_code] = {}

                if district not in target_dict[state_code]:
                    target_dict[state_code][district] = {}

                if program not in target_dict[state_code][district]:
                    target_dict[state_code][district][program] = []

                target_dict[state_code][district][program].append(row_dict)

            except Exception as e:
                print(f"❌ Error processing row {row_idx}: {str(e)}")
                continue

        # === Write JSONs ===
        for category_name, data_dict in [('SLC', slc_data), ('WLC', wlc_data)]:
            category_dir = os.path.join(output_base_dir, category_name)
            os.makedirs(category_dir, exist_ok=True)

            for state_code, districts in data_dict.items():
                state_folder = os.path.join(category_dir, str(state_code))
                os.makedirs(state_folder, exist_ok=True)

                out_file = os.path.join(state_folder, f"{state_code}.json")
                with open(out_file, 'w', encoding='utf-8') as f:
                    json.dump(districts, f, indent=2, ensure_ascii=False)

                print(f"✅ Saved: {out_file}")

    except Exception as e:
        print(f"❌ Fatal Error: {str(e)}")
