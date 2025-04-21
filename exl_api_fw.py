import os
import re
import json
import requests
import pandas as pd
from typing import Dict, List, Any, Optional, Union
import sys
# import openpyxl # Removed openpyxl
# from openpyxl.utils import get_column_letter # Removed openpyxl helper
import traceback  # Import traceback for detailed error info
from collections import OrderedDict, defaultdict  # Use OrderedDict to maintain column order

from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

class APITestFramework:
    def __init__(self, xlsx_path: str):
        """Initialize the API test framework with the path to an Excel file"""
        self.xlsx_path = xlsx_path
        self.environment_vars = {}
        # self.results will now store detailed results per test case (used for final summary)
        self.results: Dict[str, Dict[str, Any]] = {}
        self.verbose = False  # Initialize verbose flag
        self.load_environment()

    def load_environment(self) -> None:
        """Load environment variables from the first sheet of the Excel file"""
        try:
            # Use header=None to ensure it reads from the very first row
            env_df = pd.read_excel(self.xlsx_path, sheet_name=0, header=None)
            # Take only first two columns as key-value pairs
            # Filter out rows where the first column (key) is NaN or empty after stripping
            env_df = env_df.dropna(subset=[0])
            env_df[0] = env_df[0].astype(str).str.strip()  # Ensure key is string and strip whitespace
            env_df = env_df[env_df[0] != '']  # Remove rows where key is empty after strip

            for _, row in env_df.iterrows():
                key = row.iloc[0]
                # Handle potential NaN in the second column by treating it as empty string
                value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                self.environment_vars[key] = value
            print(f"Loaded {len(self.environment_vars)} environment variables")
        except FileNotFoundError:
            print(f"Error: Environment file not found at '{self.xlsx_path}' when loading environment.")
            # Decide if execution should stop here or try to continue without environment
            # sys.exit(1) # Uncomment to exit on file not found
        except Exception as e:
            print(f"Error loading environment variables from sheet 1: {e}")
            # sys.exit(1) # Uncomment to exit on environment loading failure

    def replace_env_vars(self, text: Union[str, Any]) -> Union[str, Any]:
        """Replace environment variables in text with their values"""
        if not isinstance(text, str):
            return text

        pattern = r'\$([a-zA-Z0-9_]+)'

        def replace_var(match):
            var_name = match.group(1)
            if var_name in self.environment_vars:
                return self.environment_vars[var_name]
            else:
                # print(f"Warning: Environment variable ${var_name} not found") # Suppress repetitive warnings
                return match.group(0)  # Return original string if not found

        return re.sub(pattern, replace_var, text)

    def parse_dict_list(self, text: str) -> List[Dict[str, str]]:
        """Parse a string representation of a list of dictionaries"""
        if not text or pd.isna(text):
            return []

        text = self.replace_env_vars(str(text))  # Ensure text is a string

        try:
            # Attempt to parse as JSON first, handling common variations
            # Replace single quotes with double quotes for JSON compatibility
            json_string = text.strip().replace("'", '"')
            return json.loads(json_string)

        except json.JSONDecodeError as e:
            # Fallback to a simpler parsing if JSON fails (less reliable)
            try:
                result = []
                # Attempt to handle [{'key', 'value'}, {'key2', 'value2'}] or [{'key': 'value'}]
                items = re.findall(r'\{(.*?)\}', text)
                for item in items:
                    item = item.strip()
                    if not item: continue

                    parts = [p.strip("'\" ") for p in item.split(',', 1)]
                    if len(parts) == 2:
                        key, value = parts
                        result.append({key: value})
                    elif ':' in item:
                        parts = [p.strip("'\" ") for p in item.split(':', 1)]
                        if len(parts) == 2:
                            key, value = parts
                            result.append({key: value})
                    # else: print(f"Warning: Fallback parsing found unhandled pair format in '{item}'")

                if result:
                    # print(f"Warning: Fell back to basic parsing for '{text[:100]}...'. JSON parsing preferred.")
                    return result
                else:
                    # print(f"Error: Fallback parsing failed for '{text[:100]}...'")
                    return []
            except Exception as e_fallback:
                print(f"Error during fallback parsing dictionary list '{text[:100]}...': {e_fallback}")
                return []
        except Exception as e:
            print(f"Unexpected error parsing dictionary list '{text[:100]}...': {e}")
            return []

    def parse_json_body(self, body_text: str) -> Any:
        """Parse the body text as JSON"""
        if not body_text or pd.isna(body_text):
            return None

        body_text = self.replace_env_vars(str(body_text))  # Ensure text is a string

        try:
            return json.loads(body_text)
        except json.JSONDecodeError as e:
            # print(f"Error parsing JSON body '{body_text[:100]}...': {e}")
            return None
        except Exception as e:
            print(f"Unexpected error parsing JSON body '{body_text[:100]}...': {e}")
            return None

    def parse_cookies(self, response) -> Dict[str, str]:
        """Extract cookies from response using requests' built-in cookiejar"""
        cookies = {}
        if response and hasattr(response, 'cookies'):
            for cookie in response.cookies:
                cookies[cookie.name] = cookie.value
        return cookies

    def _get_nested_value(self, obj: Any, path: str) -> Any:
        """
        Traverse an object (dict or list) using a dot-notation path
        including array indexing like 'key.list[index].nested_key'.
        Returns None if any part of the path is invalid or not found.
        """
        if not path:
            return obj

        current_value = obj

        segments = path.split('.')
        for segment in segments:
            if current_value is None:
                # print(f"Warning: Cannot access '{segment}', previous path segment returned None.")
                return None

            array_match = re.match(r'([a-zA-Z0-9_]+)\[(\d+)\]$', segment)

            if array_match:
                key_name = array_match.group(1)
                try:
                    index = int(array_match.group(2))
                except ValueError:
                    print(f"Warning: Invalid array index format in path segment '{segment}'")
                    return None

                if isinstance(current_value, dict) and key_name in current_value:
                    list_obj = current_value.get(key_name)
                    if isinstance(list_obj, list):
                        if 0 <= index < len(list_obj):
                            current_value = list_obj[index]
                        else:
                            # print(f"Warning: List index {index} out of bounds for key '{key_name}'")
                            return None
                    else:
                        # print(f"Warning: Value at key '{key_name}' is not a list (found {type(list_obj).__name__})")
                        return None
                else:
                    # print(f"Warning: Dictionary key '{key_name}' not found or current object is not a dictionary (found {type(current_value).__name__})")
                    return None
            elif segment.isdigit():
                try:
                    index = int(segment)
                    if isinstance(current_value, list):
                        if 0 <= index < len(current_value):
                            current_value = current_value[index]
                        else:
                            # print(f"Warning: List index {index} out of bounds.")
                            return None
                    else:
                        # print(f"Warning: Cannot access index '{segment}', current object is not a list (found {type(current_value).__name__})")
                        return None
                except ValueError:
                    # print(f"Warning: Invalid numeric segment in path '{segment}'")
                    return None
            else:
                if isinstance(current_value, dict) and segment in current_value:
                    current_value = current_value.get(segment)
                elif isinstance(current_value, list) and segment == 'length':  # Basic list length access
                    current_value = len(current_value) if isinstance(current_value, list) else None
                    if current_value is None:
                        # print(f"Warning: Cannot get '.length' on non-list object.")
                        return None
                else:
                    # print(f"Warning: Path segment '{segment}' not found or current object type is incorrect for access (found {type(current_value).__name__})")
                    return None

        return current_value

    def evaluate_condition(self, condition: str, result: Dict[str, Any]) -> bool:
        """Evaluate a condition against the result"""
        if not condition or pd.isna(condition):
            return True

        condition = self.replace_env_vars(str(condition))

        def contains(data, value):
            data_str = str(data) if data is not None else ""
            value_str = str(value) if value is not None else ""
            is_contained = value_str in data_str
            if not is_contained and self.verbose:
                print(f"  Condition Failed: Expected '{value_str}' to be contained in '{data_str[:200]}...'")
            return is_contained

        def equal(data, value):
            is_equal = data == value
            if not is_equal and self.verbose:
                print(
                    f"  Condition Failed: Expected '{value}' (type: {type(value).__name__}), Actual '{data}' (type: {type(data).__name__})")
            return is_equal

        def is_numeric(value):
            try:
                float(value)
                return True
            except (ValueError, TypeError):
                return False

        def greater_than(data, value):
            if not is_numeric(data) or not is_numeric(value):
                if self.verbose:
                    print(
                        f"  Condition Failed: Cannot perform greater_than comparison on non-numeric types: '{data}' and '{value}'")
                return False
            is_greater = float(data) > float(value)
            if not is_greater and self.verbose:
                print(f"  Condition Failed: Expected value > {value}, Actual {data}")
            return is_greater

        def less_than(data, value):
            if not is_numeric(data) or not is_numeric(value):
                if self.verbose:
                    print(
                        f"  Condition Failed: Cannot perform less_than comparison on non-numeric types: '{data}' and '{value}'")
                return False
            is_less = float(data) < float(value)
            if not is_less and self.verbose:
                print(f"  Condition Failed: Expected value < {value}, Actual {data}")
            return is_less

        true = True
        false = False
        null = None

        try:
            pattern = r'result\.([a-zA-Z0-9_\[\].]+)'

            def replace_result_ref(match):
                path = match.group(1)
                value = self._get_nested_value(result, path)

                if isinstance(value, (dict, list)):
                    return json.dumps(value, ensure_ascii=False)
                elif isinstance(value, str):
                    return repr(value)
                elif value is None:
                    return 'None'
                elif isinstance(value, bool):
                    return str(value)
                else:
                    return repr(value)

            eval_condition = re.sub(pattern, replace_result_ref, condition)

            if self.verbose:
                print(f"  Evaluating condition string: {eval_condition}")

            context = {
                'contains': contains,
                'equal': equal,
                'greatThan': greater_than,
                'lessThan': less_than,
                'result': result,
                'true': True,
                'false': False,
                'null': None,
            }

            eval_result = eval(eval_condition, {"__builtins__": {}}, context)

            if self.verbose:
                print(f"  Condition '{condition}' evaluated to: {eval_result}")

            return bool(eval_result)

        except Exception as e:
            print(f"Error evaluating condition '{condition}': {e}")
            # if self.verbose: traceback.print_exc()
            return False

    def execute_action(self, action: str, result: Dict[str, Any]) -> None:
        """Execute an action, such as setting an environment variable"""
        if not action or pd.isna(action):
            return

        actions = [act.strip() for act in re.split(r'[;\n]', str(action)) if act.strip()]

        for single_action in actions:
            if self.verbose:
                print(f"  Executing action: {single_action}")

            pattern = re.compile(r'\$([a-zA-Z0-9_]+)\s*=\s*result\.([a-zA-Z0-9_\[\].]+)')
            match = pattern.search(single_action)

            if match:
                var_name, result_path = match.groups()
                try:
                    value = self._get_nested_value(result, result_path)

                    if value is not None:
                        if isinstance(value, (dict, list)):
                            value_str = json.dumps(value, ensure_ascii=False)
                        elif isinstance(value, bool):
                            value_str = str(value).lower()
                        elif value is None:
                            value_str = "null"
                        else:
                            value_str = str(value)

                        self.environment_vars[var_name] = value_str
                        if self.verbose:
                            print_value = value_str if len(value_str) < 100 else value_str[:97] + "..."
                            print(f"  Set environment variable ${var_name} = '{print_value}'")
                    else:
                        if self.verbose:
                            print(
                                f"  Skipping assignment for ${var_name}: value not found at path '{result_path}' or was None.")

                except Exception as e:
                    print(f"Error executing action '{single_action}': {e}")
                    # if self.verbose: traceback.print_exc()
            # else: print(f"Warning: Unrecognized action format: '{single_action}'") # Suppress warning for blank lines etc.

    def parse_headers(self, header_text: str) -> Dict[str, str]:
        """Parse headers from various formats into a dictionary"""
        if not header_text or pd.isna(header_text):
            return {}

        header_text = self.replace_env_vars(str(header_text))
        headers = {}

        try:
            json_string = header_text.strip().replace("'", '"')
            parsed_data = json.loads(json_string)

            if isinstance(parsed_data, list):
                for item in parsed_data:
                    if isinstance(item, dict):
                        headers.update(item)
                    # else: if self.verbose: print(f"Warning: Unexpected item type in headers list: {type(item).__name__}")
            elif isinstance(parsed_data, dict):
                headers.update(parsed_data)
            # else: if self.verbose: print(f"Warning: Unexpected JSON format for headers: {type(parsed_data).__name__}")

            return {str(k): str(v) for k, v in headers.items()}

        except json.JSONDecodeError as e:
            try:
                headers_list = self.parse_dict_list(header_text)  # Reuse dict list parser as fallback
                for header_dict in headers_list:
                    headers.update(header_dict)
                if headers:
                    # if self.verbose: print(f"Warning: Fell back to basic parsing for headers '{header_text[:100]}...'. JSON format preferred.")
                    return headers
                else:
                    # print(f"Error: Fallback parsing failed for headers '{header_text[:100]}...'")
                    return {}

            except Exception as e_fallback:
                print(f"Error during fallback parsing headers '{header_text[:100]}...': {e_fallback}")
                return {}
        except Exception as e:
            print(f"Unexpected error parsing headers '{header_text[:100]}...': {e}")
            return {}

    def execute_test_case(self, test_case: pd.Series, excel_sheet_name: str) -> Dict[str, Any]:
        """Execute a single test case and return detailed results."""
        test_name = str(test_case.get('test_case_name',
                                      f'Unnamed Test Case Row {test_case.name + 2}'))  # Default name + Excel row number
        api_path_raw = test_case.get('api_path', None)

        # Store result by sheet::name for the global summary
        full_test_name = f"{excel_sheet_name}::{test_name}"
        detailed_result = {
            "test_name": test_name,  # Add test name to the detailed result
            "status": "Skipped",  # Default status
            "actual_code": "N/A",
            "body_validation": "N/A",
            "header_validation": "N/A",
            "details": "",
            "elapsed_time_ms": "N/A", # Initialize elapsed_time_ms here
            # Optional: store request/response snippets in detailed_result for rich printing later
            # "request": {},
            # "response": {}
        }
        self.results[full_test_name] = detailed_result

        if pd.isna(api_path_raw) or str(api_path_raw).strip() == '':
            print(f"\nSkipping test case '{test_name}' in sheet '{excel_sheet_name}': 'api_path' is missing or empty.")
            detailed_result["details"] = "'api_path' is missing or empty."
            detailed_result["status"] = "Skipped"
            return detailed_result

        # print(f"\nExecuting test case: '{test_name}' in sheet '{excel_sheet_name}'")

        # Check verbose flag specific to this test case row
        verbose_row = False
        if 'verbose' in test_case and not pd.isna(test_case['verbose']):
            verbose_value = str(test_case['verbose']).lower().strip()
            verbose_row = verbose_value in ('true', 'yes', '1')
        self.verbose = verbose_row  # Set class-level verbose for helpers

        try:
            # Extract test case parameters
            api_path = self.replace_env_vars(str(api_path_raw))
            method = str(test_case.get('method', 'GET')).upper()  # Default to GET

            # Parse query parameters
            query_params = {}
            if 'query_param' in test_case and not pd.isna(test_case['query_param']):
                query_param_list = self.parse_dict_list(test_case['query_param'])
                for param_dict in query_param_list:
                    query_params.update(param_dict)

            # Parse headers
            headers = self.parse_headers(test_case.get('inject_header', None))

            # Parse body
            body_input = test_case.get('body', None)
            body = self.parse_json_body(body_input)

            # Store request details (optional)
            # detailed_result["request"] = {"method": method, "url": api_path, "params": query_params, "headers": headers, "body": body}

            if self.verbose:
                print(f"  Request URL: {api_path}")
                print(f"  Request Method: {method}")
                if query_params: print(f"  Request Query Params: {query_params}")
                if headers: print(f"  Request Headers: {headers}")
                if body is not None:
                    body_print = json.dumps(body, indent=2, ensure_ascii=False)
                    print(f"  Request Body: {body_print[:500]}{'...' if len(body_print) > 500 else ''}")

            # Execute API request
            response = requests.request(
                method=method,
                url=api_path,
                params=query_params,
                headers=headers,
                json=body,  # Use json=body for automatic Content-Type: application/json
                timeout=15  # Add a default timeout (in seconds)
            )

            # Parse response
            response_json = None
            response_body_text = ""
            try:
                response_body_text = response.text
                content_type = response.headers.get('Content-Type', '').lower()
                if 'application/json' in content_type:
                    response_json = response.json()
                elif 'text/' in content_type or 'html' in content_type or 'xml' in content_type:
                    response_json = {"text": response.text}
                else:
                    response_json = {"content_type": content_type,
                                     "content_preview": response.text[:100] + "..." if len(
                                         response.text) > 100 else response.text}
            except json.JSONDecodeError:
                response_json = {"decoding_error": "Failed to decode JSON", "raw_response_text": response_body_text}
            except Exception as e:
                print(f"Error processing response body for {test_name}: {e}")
                response_json = {"processing_error": str(e), "raw_response_text": response_body_text}

            cookies = self.parse_cookies(response)

            api_result_data = {
                "code": response.status_code,
                "body": response_json,
                "headers": dict(response.headers),
                "cookies": cookies,
                "elapsed_time_ms": response.elapsed.total_seconds() * 1000
            }

            # Store response details (optional)
            # detailed_result["response"] = api_result_data

            detailed_result["actual_code"] = response.status_code  # Update actual code in result
            detailed_result["elapsed_time_ms"] = api_result_data["elapsed_time_ms"] # **Added: Copy elapsed time to detailed_result**

            # --- Validations ---
            body_validation_status = "N/A"
            header_validation_status = "N/A"
            test_passed_validations = True

            # Validate response code
            expected_code = test_case.get('expect_response_code', None)
            if pd.notna(expected_code):
                try:
                    expected_code = int(expected_code)
                    if api_result_data["code"] != expected_code:
                        detailed_result[
                            "details"] += f"Status Code Failed (Expected: {expected_code}, Actual: {api_result_data['code']}). "
                        test_passed_validations = False
                except ValueError:
                    print(f"Warning: Invalid value for 'expect_response_code': '{test_case['expect_response_code']}'")
                    detailed_result[
                        "details"] += f"Invalid 'expect_response_code' value: '{test_case['expect_response_code']}'. "

            # Validate response body
            expected_body = test_case.get('expect_response_body', None)
            if pd.notna(expected_body):
                if not self.evaluate_condition(expected_body, api_result_data):
                    body_validation_status = "Failed"
                    detailed_result["details"] += f"Body Validation Failed ('{expected_body}'). "
                    test_passed_validations = False
                else:
                    body_validation_status = "Passed"
            detailed_result["body_validation"] = body_validation_status

            # Validate response headers
            expected_headers = test_case.get('expect_response_header', None)
            if pd.notna(expected_headers):
                if not self.evaluate_condition(expected_headers, api_result_data):
                    header_validation_status = "Failed"
                    detailed_result["details"] += f"Header Validation Failed ('{expected_headers}'). "
                    test_passed_validations = False
                else:
                    header_validation_status = "Passed"
            detailed_result["header_validation"] = header_validation_status

            # Determine final test case status based on validations
            if test_passed_validations:
                detailed_result["status"] = "Passed"
                print(f"✅ Test case '{test_name}' PASSED")
            else:
                detailed_result["status"] = "Failed"
                print(f"❌ Test case '{test_name}' FAILED")

            # --- Actions ---
            action = test_case.get('action', None)
            if pd.notna(action):
                self.execute_action(action, api_result_data)

            return detailed_result

        except requests.exceptions.Timeout:
            print(f"❌ Timeout Error executing test case '{test_name}': Request timed out.")
            detailed_result["status"] = "Failed"
            detailed_result["details"] += "Request Timeout."
            return detailed_result
        except requests.exceptions.RequestException as e:
            print(f"❌ Request Error executing test case '{test_name}': {e}")
            detailed_result["status"] = "Failed"
            detailed_result["details"] += f"Request Error: {e}"
            return detailed_result
        except Exception as e:
            print(f"❌ Unexpected Error executing test case '{test_name}': {e}")
            detailed_result["status"] = "Error"
            detailed_result[
                "details"] += f"Unexpected Error: {e} - {traceback.format_exc()}"  # Add traceback to details
            traceback.print_exc()
            return detailed_result
        finally:
            self.verbose = False  # Reset verbose flag

    def _print_sheet_results_table(self, sheet_name: str, results_list: List[Dict[str, Any]]) -> None:
        """Prints the results for a single sheet in a formatted table, including Response Time."""
        if not results_list:
            print(f"\nNo test cases executed in sheet '{sheet_name}'.")
            return

        print(f"\n--- Results for Sheet: {sheet_name} ---")

        # Define columns and their corresponding keys in the result dictionary
        # Added 'Response Time' column
        columns = OrderedDict([
            ("Test Name", "test_name"),
            ("Response Time", "elapsed_time_ms"), # Added Response Time key
            ("Status", "status"),
            ("Code", "actual_code"),
            ("Body Val", "body_validation"),
            ("Header Val", "header_validation"),
            ("Details", "details"),
        ])

        # Define maximum width for the 'Details' column to keep the table manageable
        max_details_width = 80

        # Calculate column widths dynamically based on headers and content
        col_widths = {header: len(header) for header in columns.keys()}
        for result in results_list:
            for header, key in columns.items():
                value = result.get(key, '')
                # Format response time for display and width calculation
                if key == "elapsed_time_ms":
                    value_str = f"{value:.2f} ms" if isinstance(value, (int, float)) else str(value)
                else:
                    value_str = str(value)

                if header == "Details":
                    value_str = value_str[:max_details_width]  # Truncate for width calculation
                col_widths[header] = max(col_widths[header], len(value_str))

        # Add padding to widths
        col_padding = 2
        padded_widths = {header: width + col_padding for header, width in col_widths.items()}

        # Print Header Row
        header_row = "| " + " | ".join(
            header.ljust(padded_widths[header] - col_padding) for header in columns.keys()) + " |"
        print(header_row)

        # Print Separator Line
        separator_line = "|-" + "-|-".join(
            "-" * (padded_widths[header] - col_padding) for header in columns.keys()) + "-|"
        print(separator_line)

        # Print Data Rows
        for result in results_list:
            row_data = []
            for header, key in columns.items():
                value = result.get(key, '')
                # Format response time for display
                if key == "elapsed_time_ms":
                     value_str = f"{value:.2f} ms" if isinstance(value, (int, float)) else str(value)
                else:
                    value_str = str(value)


                # Truncate and pad details separately
                if header == "Details":
                    if len(value_str) > max_details_width:
                        value_str = value_str[:max_details_width - 3] + "..."  # Truncate and add ellipsis
                    row_data.append(value_str.ljust(padded_widths[header] - col_padding))
                else:
                    row_data.append(value_str.ljust(padded_widths[header] - col_padding))

            print("| " + " | ".join(row_data) + " |")

        print("-" * len(header_row))  # Match separator length to header row

    def run_tests(self) -> Dict[str, Dict[str, Any]]:
        """Run all test cases from the Excel file and print results as tables."""
        try:
            xl = pd.ExcelFile(self.xlsx_path)
            sheet_names = xl.sheet_names
        except FileNotFoundError:
            print(f"Error: Excel file not found at '{self.xlsx_path}'")
            return {}
        except Exception as e:
            print(f"Error reading Excel file '{self.xlsx_path}': {e}")
            return {}

        if len(sheet_names) < 2:
            print("Error: Excel file must have at least 2 sheets (Environment and at least one test sheet).")
            return {}

        # --- Execute Setup Sheet ---
        setup_sheet_name = sheet_names[1] if len(sheet_names) > 1 else None
        setup_success = True
        setup_results_list = []  # Collect results for printing table

        if setup_sheet_name:
            print(f"\n=== Running Setup Sheet: {setup_sheet_name} ===")
            try:
                setup_df = pd.read_excel(self.xlsx_path, sheet_name=setup_sheet_name)
                setup_df = setup_df.dropna(subset=['test_case_name'])

                for index, test_case in setup_df.iterrows():
                    detailed_result = self.execute_test_case(test_case, setup_sheet_name)
                    setup_results_list.append(detailed_result)  # Collect result

                    if detailed_result["status"] in ["Failed", "Error"]:
                        setup_success = False
                        print(
                            f"❌ Setup failed ('{test_case.get('test_case_name', 'Unnamed Setup Case')}'). Remaining setup tests and all main tests will be skipped.")
                        break  # Stop executing setup tests on first failure or error

            except Exception as e:
                print(f"Error processing Setup sheet '{setup_sheet_name}': {e}")
                setup_success = False
            finally:
                # Print table for the setup sheet results
                self._print_sheet_results_table(setup_sheet_name, setup_results_list)

        # --- Execute Main Test Sheets ---
        if setup_success:
            for sheet_name in sheet_names[2:]:  # Start from the 3rd sheet
                print(f"\n=== Running Test Sheet: {sheet_name} ===")
                sheet_results_list = []  # Collect results for printing table
                sheet_processing_error = None  # Track errors loading/processing sheet

                try:
                    test_df = pd.read_excel(self.xlsx_path, sheet_name=sheet_name)
                    test_df = test_df.dropna(subset=['test_case_name'])

                    sheet_has_failures = False
                    for index, test_case in test_df.iterrows():
                        detailed_result = self.execute_test_case(test_case, sheet_name)
                        sheet_results_list.append(detailed_result)  # Collect result

                        if detailed_result["status"] in ["Failed", "Error"]:
                            sheet_has_failures = True
                            # Decide if you want to stop the sheet on first failure:
                            # print(f"❌ Test case failed, skipping remaining tests in sheet '{sheet_name}'")
                            # break # Uncomment this line to stop sheet on first failure

                    if not sheet_has_failures and not test_df.empty:
                        print(f"✅ All executed tests in sheet '{sheet_name}' PASSED")
                    elif test_df.empty:
                        print(f"ℹ️ No test cases found with 'test_case_name' in sheet '{sheet_name}'")

                except Exception as e:
                    print(f"Error processing Test sheet '{sheet_name}': {e}")
                    sheet_processing_error = e  # Store error
                    # Continue to next sheet even if one sheet fails to load or run
                finally:
                    # Print table for the current sheet's results
                    self._print_sheet_results_table(sheet_name, sheet_results_list)
                    if sheet_processing_error:
                        print(f"‼️ Processing of sheet '{sheet_name}' encountered an error: {sheet_processing_error}")

        # --- Print Console Summary ---
        self._print_summary()

        return self.results  # Return the full results dictionary

    def _print_summary(self) -> None:
        """Prints the test execution summary based on self.results."""
        print("\n=== Overall Test Run Summary ===")
        total_attempted = len(self.results)
        passed_count = 0
        failed_count = 0
        error_count = 0
        skipped_count = 0
        unknown_count = 0

        if total_attempted == 0:
            print("No test cases were attempted.")
            return

        # Sort results by sheet and test name for consistent output
        sorted_test_names = sorted(self.results.keys())

        for full_test_name in sorted_test_names:
            result_data = self.results[full_test_name]
            status = result_data.get("status", "Unknown")

            if status == "Passed":
                passed_count += 1
            elif status == "Failed":
                failed_count += 1
            elif status == "Error":
                error_count += 1
            elif status == "Skipped":
                skipped_count += 1
            else:
                unknown_count += 1

        print(f"Total Test Cases Attempted: {total_attempted}")
        print(f"Passed: {passed_count}")
        print(f"Failed: {failed_count}")
        print(f"Errors: {error_count}")
        print(f"Skipped: {skipped_count}")
        print("-" * 30)

    def generate_pdf_report(self, output_path: str = "test_report.pdf"):
        """Generates a PDF report of the test results with per-sheet insights,
           failed/errored tests, and slowest tests."""
        # Ensure you have these imports at the top of your file:
        # from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, PageBreak
        # from reportlab.lib.pagesizes import letter
        # from reportlab.lib.styles import getSampleStyleSheet
        # from reportlab.lib import colors
        # from reportlab.lib.units import inch
        # from collections import defaultdict
        # import pandas as pd # Ensure pandas is imported for isnan checks

        doc = SimpleDocTemplate(output_path, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()

        # Add a main title
        elements.append(Paragraph("API Test Report", styles['Title']))
        elements.append(Spacer(1, 0.5 * inch))

        # --- Overall Summary ---
        elements.append(Paragraph("Overall Test Run Summary", styles['Heading1']))
        total_attempted_overall = len(self.results)
        passed_count_overall = 0
        failed_count_overall = 0
        error_count_overall = 0
        skipped_count_overall = 0

        for result_data in self.results.values():
            status = result_data.get("status", "Unknown")
            if status == "Passed":
                passed_count_overall += 1
            elif status == "Failed":
                failed_count_overall += 1
            elif status == "Error":
                error_count_overall += 1
            elif status == "Skipped":
                skipped_count_overall += 1

        summary_data_overall = [['Total Attempted', 'Passed', 'Failed', 'Errors', 'Skipped'],
                                [total_attempted_overall, passed_count_overall, failed_count_overall,
                                 error_count_overall, skipped_count_overall]]

        summary_table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]

        summary_table_overall = Table(summary_data_overall, colWidths=[1.2 * inch] * 5)
        summary_table_overall.setStyle(summary_table_style)
        elements.append(summary_table_overall)
        elements.append(Spacer(1, 0.75 * inch))

        # --- Group results by Sheet ---
        results_by_sheet = defaultdict(list)
        sorted_full_test_names = sorted(self.results.keys())
        for full_test_name in sorted_full_test_names:
            sheet_name, test_name = full_test_name.split("::", 1)
            results_by_sheet[sheet_name].append(self.results[full_test_name])

        # --- Add Section for Each Sheet ---
        sorted_sheet_names = sorted(results_by_sheet.keys())

        first_sheet = True

        for sheet_name in sorted_sheet_names:
            sheet_results_list = results_by_sheet[sheet_name]

            if not first_sheet:
                elements.append(PageBreak())
            else:
                first_sheet = False

            elements.append(Paragraph(f"Results for Sheet: {sheet_name}", styles['Heading2']))
            elements.append(Spacer(1, 0.25 * inch))

            # Calculate per-sheet summary
            total_attempted_sheet = len(sheet_results_list)
            passed_count_sheet = 0
            failed_count_sheet = 0
            error_count_sheet = 0
            skipped_count_sheet = 0

            for result_data in sheet_results_list:
                status = result_data.get("status", "Unknown")
                if status == "Passed":
                    passed_count_sheet += 1
                elif status == "Failed":
                    failed_count_sheet += 1
                elif status == "Error":
                    error_count_sheet += 1
                elif status == "Skipped":
                    skipped_count_sheet += 1

            summary_data_sheet = [['Total Attempted', 'Passed', 'Failed', 'Errors', 'Skipped'],
                                  [total_attempted_sheet, passed_count_sheet, failed_count_sheet, error_count_sheet,
                                   skipped_count_sheet]]

            summary_table_sheet = Table(summary_data_sheet, colWidths=[1.2 * inch] * 5)
            summary_table_sheet.setStyle(summary_table_style)
            elements.append(summary_table_sheet)
            elements.append(Spacer(1, 0.4 * inch))

            # Add Detailed Results for this Sheet
            elements.append(Paragraph("Detailed Results (This Sheet)", styles['Heading3']))
            elements.append(Spacer(1, 0.1 * inch))

            for result_data in sheet_results_list:
                test_name = result_data.get("test_name", "N/A")
                status = result_data.get("status", "Unknown")
                actual_code = result_data.get("actual_code", "N/A")
                elapsed_time_ms = result_data.get("elapsed_time_ms", "N/A")
                body_validation = result_data.get("body_validation", "N/A")
                header_validation = result_data.get("header_validation", "N/A")
                details = result_data.get("details", "")

                text_color = colors.black
                if status == "Passed":
                    text_color = colors.green
                elif status in ["Failed", "Error"]:
                    text_color = colors.red

                elements.append(Paragraph(f"<b>Test Case:</b> {test_name}", styles['Normal']))
                elements.append(
                    Paragraph(f"<b>Status:</b> <font color='{text_color}'>{status}</font>", styles['Normal']))
                elements.append(Paragraph(f"<b>Response Code:</b> {actual_code}", styles['Normal']))

                if isinstance(elapsed_time_ms, (int, float)):
                    elements.append(Paragraph(f"<b>Response Time:</b> {elapsed_time_ms:.2f} ms", styles['Normal']))
                else:
                    elements.append(Paragraph(f"<b>Response Time:</b> {elapsed_time_ms}", styles['Normal']))

                elements.append(Paragraph(f"<b>Body Validation:</b> {body_validation}", styles['Normal']))
                elements.append(Paragraph(f"<b>Header Validation:</b> {header_validation}", styles['Normal']))
                if details:
                    details_str = str(details) if not pd.isna(details) else ""
                    elements.append(Paragraph(f"<b>Details:</b> {details_str}", styles['Normal']))
                elements.append(Spacer(1, 0.25 * inch))

        # --- Section for Failed and Errored Test Cases ---
        # Iterate through items to get the full test name (key) and result data (value)
        failed_errored_tests_items = [
            (full_name, result) for full_name, result in self.results.items()
            if result.get("status") in ["Failed", "Error"]
        ]

        if failed_errored_tests_items:
            elements.append(PageBreak())
            elements.append(Paragraph("Failed and Errored Test Cases", styles['Heading1']))
            elements.append(Spacer(1, 0.25 * inch))

            for full_test_name, result_data in failed_errored_tests_items:
                # Extract sheet_name and test_name from the full_test_name key
                sheet_name, test_name = full_test_name.split("::", 1)

                status = result_data.get("status", "Unknown")
                actual_code = result_data.get("actual_code", "N/A")
                details = result_data.get("details", "")

                # Use the extracted sheet_name and test_name for the title
                elements.append(Paragraph(f"<b>Test Case:</b> {sheet_name}::{test_name}", styles['Normal']))
                elements.append(
                    Paragraph(f"<b>Status:</b> <font color='{colors.red}'>{status}</font>", styles['Normal']))
                elements.append(Paragraph(f"<b>Response Code:</b> {actual_code}", styles['Normal']))
                if details:
                    details_str = str(details) if not pd.isna(details) else ""
                    elements.append(Paragraph(f"<b>Details:</b> {details_str}", styles['Normal']))
                elements.append(Spacer(1, 0.25 * inch))

        # --- Section for Slowest Tests ---
        # Filter out tests that don't have a valid elapsed time before sorting
        # Iterate through items to keep the full test name (key) for sorting
        tests_with_time_items = [
            (full_name, result) for full_name, result in self.results.items()
            if isinstance(result.get("elapsed_time_ms"), (int, float))
        ]

        # Sort by elapsed time in descending order using the value from the dictionary
        slowest_tests_items = sorted(tests_with_time_items, key=lambda item: item[1].get("elapsed_time_ms", 0),
                                     reverse=True)

        # Define how many slowest tests to show (e.g., top 10)
        top_n_slowest = 10
        slowest_tests_to_show_items = slowest_tests_items[:top_n_slowest]

        if slowest_tests_to_show_items:
            elements.append(PageBreak())
            elements.append(Paragraph(f"Top {top_n_slowest} Slowest Test Cases", styles['Heading1']))
            elements.append(Spacer(1, 0.25 * inch))

            for full_test_name, result_data in slowest_tests_to_show_items:
                # Extract sheet_name and test_name from the full_test_name key
                sheet_name, test_name = full_test_name.split("::", 1)

                elapsed_time_ms = result_data.get("elapsed_time_ms", "N/A")
                status = result_data.get("status", "Unknown")

                # Use the extracted sheet_name and test_name for the title
                elements.append(Paragraph(f"<b>Test Case:</b> {sheet_name}::{test_name}", styles['Normal']))
                if isinstance(elapsed_time_ms, (int, float)):
                    elements.append(Paragraph(f"<b>Response Time:</b> {elapsed_time_ms:.2f} ms", styles['Normal']))
                else:
                    elements.append(Paragraph(f"<b>Response Time:</b> {elapsed_time_ms}", styles['Normal']))

                elements.append(Paragraph(f"<b>Status:</b> {status}", styles['Normal']))
                elements.append(Spacer(1, 0.25 * inch))

        # Build the PDF document
        try:
            doc.build(elements)
            print(f"PDF report generated successfully at {output_path}")
        except Exception as e:
            print(f"Error generating PDF report: {e}")