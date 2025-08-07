import openpyxl
import json
import os
from collections import defaultdict
from constants import PAGE_METADATA, TABS_METADATA
import importlib.util

try:
    from state_code_generator import state_code_generator
except ImportError:

    import importlib.util
    import os
    
    print("Could not import state_code_generator, loading manually...")
    
    # Get the correct path to state_code_generator.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    state_gen_path = os.path.join(script_dir, 'state_code_generator.py')
    
    if os.path.exists(state_gen_path):
        spec = importlib.util.spec_from_file_location('state_code_generator', state_gen_path)
        state_gen_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(state_gen_module)
        state_code_generator = state_gen_module.state_code_generator
        print("Successfully loaded state_code_generator function")
    else:
        print(f"state_code_generator.py not found at {state_gen_path}")
        def state_code_generator(excel_file):
            print("state_code_generator function not available")
            return False

def load_state_codes(excel_file):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        state_codes_file = os.path.join(script_dir, "..", "pages", "state_code_details.json")
        state_code_generator(excel_file)
        if not os.path.exists(state_codes_file):
            return None
        with open(state_codes_file, 'r', encoding='utf-8') as file:
            return json.load(file)
    except Exception:
        return None

def update_district_view_indicators(excel_file):
    try:
        state_codes = load_state_codes(excel_file)
        if not state_codes:
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        json_file_path = os.path.join(script_dir, "..", "pages", "district-view-indicators.json")

        # if not os.path.exists(excel_file_path):
        #     print(f"Error: Excel file not found at {excel_file_path}")
        #     return

        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        try:
            sheet = workbook[PAGE_METADATA["STATE_DETAILS"]]
        except KeyError:
            print(f"Sheet not found: {PAGE_METADATA['STATE_DETAILS']}")
            return

        expected_headers = TABS_METADATA["STATE_DETAILS"]
        column_indices = {}
        for cell in sheet[1]:
            if cell.value and cell.value.strip() in expected_headers:
                column_indices[cell.value.strip()] = cell.column

        missing_columns = [col for col in expected_headers if col not in column_indices]
        if missing_columns:
            print(f"Missing required columns: {missing_columns}")
            return

        states_data = {}
        states_mission_data = {}
        overview_aggregates = defaultdict(int)

        excluded_codes = ["categories", "state led missions", "district led missions", "community led missions"]

        row_num = 2
        while True:
            state_name = sheet.cell(row=row_num, column=column_indices["State Name"]).value
            if not state_name:
                break

            indicator = sheet.cell(row=row_num, column=column_indices["Indicator"]).value or ""
            data_value = sheet.cell(row=row_num, column=column_indices["Data"]).value

            indicator = str(indicator).strip()
            state_name = str(state_name).strip()
            code = indicator.lower().strip()

            if "categories" in code:
                row_num += 1
                continue

            if state_name not in state_codes:
                row_num += 1
                continue
            state_code = state_codes[state_name]

            try:
                if isinstance(data_value, (int, float)):
                    processed_value = int(data_value) if data_value == int(data_value) else data_value
                elif isinstance(data_value, str):
                    processed_value = int(float(data_value.strip('%'))) if '%' not in data_value else data_value
                else:
                    processed_value = 0
            except:
                processed_value = 0

            if state_code not in states_data:
                states_data[state_code] = {
                    "label": state_name,
                    "type": "category_4",
                    "details": []
                }
                states_mission_data[state_code] = {
                    "state_led_missions": 0,
                    "district_led_missions": 0
                }

            if code == "state led missions":
                states_mission_data[state_code]["state_led_missions"] = processed_value
            elif code == "district led missions":
                states_mission_data[state_code]["district_led_missions"] = processed_value

            if code not in excluded_codes:
                states_data[state_code]["details"].append({
                    "value": processed_value,
                    "code": indicator
                })
                overview_aggregates[indicator] += processed_value if isinstance(processed_value, int) else 0

            row_num += 1

        # Determine categories
        for code, data in states_data.items():
            state_led = states_mission_data[code]["state_led_missions"]
            district_led = states_mission_data[code]["district_led_missions"]

            if state_led > 0 and district_led > 0:
                data["type"] = "category_1"
            elif state_led > 0:
                data["type"] = "category_2"
            elif district_led > 0:
                data["type"] = "category_3"
            else:
                data["type"] = "category_4"

        workbook.close()

        # Load or create JSON
        if os.path.exists(json_file_path):
            with open(json_file_path, 'r', encoding='utf-8') as f:
                district_indicators = json.load(f)
        else:
            district_indicators = {
                "result": {
                    "states": {},
                    "overview": {
                        "label": "india",
                        "type": "category_4",
                        "details": []
                    },
                    "meta": {}
                }
            }

        # Update overview details
        overview_details = [
            {"code": key, "value": value} for key, value in overview_aggregates.items()
        ]
        district_indicators["result"]["overview"] = {
            "label": "india",
            "type": "category_4",
            "details": overview_details
        }

        district_indicators["result"]["states"] = states_data

        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(district_indicators, f, indent=2, ensure_ascii=False)

        gcp_access_path = os.path.join(script_dir, '..', 'cloud-scripts', 'gcp_access.py')
        spec = importlib.util.spec_from_file_location('gcp_access', gcp_access_path)
        gcp_access = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(gcp_access)

        private_key_path = os.path.join(script_dir, "..", "private-key.json")

        folder_url = gcp_access.upload_file_to_gcs_and_get_directory(
            bucket_name="dev-sg-dashboard",
            source_file_path=json_file_path,
            destination_blob_name="sg-dashboard/district-view-indicators.json"
        )

        if folder_url:
            print(f"Successfully uploaded and got public folder URL DIST: {folder_url}")
        else:
            print("Failed to upload file to GCS. Check logs for details.")

        print("✅ district-view-indicators.json updated successfully.")

    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    update_district_view_indicators()
