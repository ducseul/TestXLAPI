import os
import re
import json
import requests
import pandas as pd
from typing import Dict, List, Any, Optional, Union
import sys # Import sys for error handling

class APITestFramework:
    def __init__(self, xlsx_path: str):
        """Initialize the API test framework with the path to an Excel file"""
        self.xlsx_path = xlsx_path
        self.environment_vars = {}
        self.results = {}
        self.load_environment()

    def load_environment(self) -> None:
        """Load environment variables from the first sheet of the Excel file"""
        try:
            # Use header=None to ensure it reads from the very first row
            env_df = pd.read_excel(self.xlsx_path, sheet_name=0, header=None)
            # Take only first two columns as key-value pairs
            # Filter out rows where the first column (key) is NaN
            env_df = env_df.dropna(subset=[0])
            for _, row in env_df.iterrows():
                key = str(row.iloc[0]).strip()
                # Handle potential NaN in the second column by treating it as empty string
                value = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                self.environment_vars[key] = value
            print(f"Loaded {len(self.environment_vars)} environment variables")
        except Exception as e:
            print(f"Error loading environment variables: {e}")
            # Optionally, exit if environment loading fails
            # sys.exit(1)

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
                print(f"Warning: Environment variable ${var_name} not found")
                return match.group(0) # Return original string if not found

        return re.sub(pattern, replace_var, text)

    def parse_dict_list(self, text: str) -> List[Dict[str, str]]:
        """Parse a string representation of a list of dictionaries"""
        if not text or pd.isna(text):
            return []

        text = self.replace_env_vars(str(text)) # Ensure text is a string

        try:
            # Attempt to parse as JSON first, handling common variations
            # Replace single quotes with double quotes for JSON compatibility
            # Handle cases like [{'key', 'value'}] by converting to [{'key': 'value'}]
            # This is still risky, direct JSON/dict input is preferred
            # Let's stick to a more robust JSON parse after replacement
            # We assume the input, after env var replacement, is JSON-like
            # Allowing single quotes for keys and values in the input spreadsheet is common
            json_string = text.strip().replace("'", '"')
            # Basic attempt to fix {'a': 'b'} vs {"a": "b"} vs {"a": "b", }
            json_string = re.sub(r'(?<={|,)\s*(\w+)\s*:', r'"\1":', json_string) # Add quotes around unquoted keys
            # This regex replacement can be tricky and might break valid JSON
            # A safer approach is often manual parsing or assuming strict JSON

            # Let's stick to a simple JSON load after replacing single quotes
            return json.loads(json_string)

        except json.JSONDecodeError as e:
            print(f"JSON Decode Error parsing dictionary list '{text}': {e}")
            # Fallback to a simpler parsing if JSON fails
            try:
                result = []
                # This fallback is very basic and prone to errors for complex inputs
                items = text.strip('[]').split('}, {')
                for item in items:
                    item = item.strip('{}')
                    parts = item.split(', ')
                    if len(parts) == 2:
                        key = parts[0].strip("'\" ")
                        value = parts[1].strip("'\" ")
                        result.append({key: value})
                    elif ':' in item: # Also handle key:value format within {}
                         key, value = [p.strip("'\" ") for p in item.split(':', 1)]
                         result.append({key: value})
                if result: # Only return fallback result if something was parsed
                    print(f"Warning: Fell back to basic parsing for '{text}'. JSON parsing preferred.")
                    return result
                else:
                    print(f"Error: Fallback parsing failed for '{text}'")
                    return []
            except Exception as e_fallback:
                print(f"Error during fallback parsing dictionary list '{text}': {e_fallback}")
                return []
        except Exception as e:
             print(f"Unexpected error parsing dictionary list '{text}': {e}")
             return []


    def parse_json_body(self, body_text: str) -> Optional[Dict[str, Any]]:
        """Parse the body text as JSON"""
        if not body_text or pd.isna(body_text):
            return None

        body_text = self.replace_env_vars(str(body_text)) # Ensure text is a string

        try:
            return json.loads(body_text)
        except json.JSONDecodeError as e:
            print(f"Error parsing JSON body '{body_text[:100]}...': {e}") # Print truncated body for error
            return None
        except Exception as e:
             print(f"Unexpected error parsing JSON body '{body_text[:100]}...': {e}")
             return None


    def parse_cookies(self, response) -> Dict[str, str]:
        """Extract cookies from response headers into a dictionary"""
        cookies = {}
        if response.cookies:
            for cookie in response.cookies:
                cookies[cookie.name] = cookie.value

        # Also try to parse from Set-Cookie header if response.cookies is empty or incomplete
        # Note: requests.cookies should be the most reliable source
        # This manual parsing is a fallback and might not handle all complexities (expires, path, etc.)
        if not cookies and 'Set-Cookie' in response.headers:
             # Set-Cookie header can appear multiple times, requests usually combines them
             # getall might not be available on all header objects, use get with split
            set_cookie_header = response.headers.get('Set-Cookie', '')
            if set_cookie_header:
                # Split by comma, but be careful about commas within dates etc.
                # A robust parser would be needed for full compliance.
                # Basic split by primary delimiter (usually comma or just multiple headers)
                # requests library handles this better via response.cookies
                # Let's trust response.cookies primarily. If it's empty, this fallback is needed.
                # A simpler fallback: split by ';' for a single header value string
                 try:
                     # Example: "cookie1=value1; expires=...; path=/, cookie2=value2"
                     # This basic split won't handle quoted commas or semicolons well
                     cookie_strings = set_cookie_header.split(',') # Simple split by comma
                     for cookie_str in cookie_strings:
                         # Extract name=value part
                         name_value_part = cookie_str.split(';')[0].strip()
                         if '=' in name_value_part:
                             name, value = name_value_part.split('=', 1)
                             cookies[name] = value.strip() # Strip whitespace from value
                 except Exception as e:
                     print(f"Warning: Basic fallback parsing of Set-Cookie header failed: {e}")


        return cookies

    def _get_nested_value(self, obj: Any, path: str) -> Any:
        """
        Traverse an object (dict or list) using a dot-notation path
        including array indexing like 'key.list[index].nested_key'.
        """
        if not path:
            return obj # Return the object itself if path is empty

        # Use regex to find segments: either key[index] or just key
        # Example: 'body.documentProcesses[0].documentReceive.documentReceiveId'
        # Should yield segments: 'body', 'documentProcesses[0]', 'documentReceive', 'documentReceiveId'
        # Then parse each segment further.

        # Split path by '.' first, then handle [index] within parts
        parts = path.split('.')
        current_value = obj

        for part in parts:
            if current_value is None:
                # If we've already failed to find a parent object, stop.
                return None

            # Check if the part contains an array index like 'key[index]'
            array_match = re.match(r'([a-zA-Z0-9_]+)\[(\d+)\]$', part)

            if array_match:
                key_name = array_match.group(1)
                try:
                    index = int(array_match.group(2))
                except ValueError:
                    print(f"Warning: Invalid array index format in path part '{part}'")
                    return None # Invalid path

                # First, access the dictionary key
                if isinstance(current_value, dict) and key_name in current_value:
                    list_obj = current_value[key_name]
                    # Then, access the list index
                    if isinstance(list_obj, list):
                        if 0 <= index < len(list_obj):
                            current_value = list_obj[index]
                        else:
                            print(f"Warning: List index {index} out of bounds for key '{key_name}'")
                            return None # Index out of bounds
                    else:
                        print(f"Warning: Value at key '{key_name}' is not a list")
                        return None # Not a list
                else:
                    print(f"Warning: Key '{key_name}' not found or current object is not a dictionary")
                    return None # Key not found or wrong type
            else:
                # Standard dictionary key access or list index if the part is just a number
                if isinstance(current_value, dict) and part in current_value:
                    current_value = current_value[part]
                elif isinstance(current_value, list) and part.isdigit():
                    try:
                        index = int(part)
                        if 0 <= index < len(current_value):
                            current_value = current_value[index]
                        else:
                             print(f"Warning: List index {index} out of bounds")
                             return None # Index out of bounds
                    except ValueError:
                         print(f"Warning: Invalid list index format '{part}'")
                         return None # Invalid index format
                else:
                    # Path part not found in dict, not a valid list index for a list, or wrong type
                    print(f"Warning: Path part '{part}' not found or current object type is incorrect for access")
                    return None # Path part not found

        return current_value


    def evaluate_condition(self, condition: str, result: Dict[str, Any]) -> bool:
        """Evaluate a condition against the result"""
        if not condition or pd.isna(condition):
            return True

        condition = self.replace_env_vars(str(condition)) # Ensure condition is string

        # Create a safe evaluation context
        # Use lambda functions to allow access to result within the function scope if needed
        # However, the current approach replaces result.path references before eval, which is safer.
        def contains(data, value):
            # Convert data and value to string for comparison, handle None
            data_str = str(data) if data is not None else ""
            value_str = str(value) if value is not None else ""
            is_contained = value_str in data_str

            if not is_contained:
                print(f"  Expected: '{value_str}' to be contained in")
                # Print a potentially large data string, maybe truncate
                print(f"  Actual: '{data_str[:200]}...'")
            return is_contained

        def equal(data, value):
             # Perform equality check, handling None and type differences gracefully
             # Strict equality (==) is often fine, but be mindful of int vs float vs string "1" == 1
             # Let's keep simple == for now as it's the standard Python behavior
             is_equal = data == value
             if not is_equal:
                 print(f"  Expected: '{value}' (type: {type(value).__name__})")
                 print(f"  Actual: '{data}' (type: {type(data).__name__})")
             return is_equal

        def greater_than(data, value):
             # Try converting to float for comparison
             try:
                 data_f = float(data)
                 value_f = float(value)
                 is_greater = data_f > value_f
                 if not is_greater:
                      print(f"  Expected: value > {value_f}")
                      print(f"  Actual: {data_f}")
                 return is_greater
             except (ValueError, TypeError):
                 print(f"Error: Cannot perform greater_than comparison on non-numeric types: '{data}' and '{value}'")
                 return False

        def less_than(data, value):
            # Try converting to float for comparison
            try:
                 data_f = float(data)
                 value_f = float(value)
                 is_less = data_f < value_f
                 if not is_less:
                      print(f"  Expected: value < {value_f}")
                      print(f"  Actual: {data_f}")
                 return is_less
            except (ValueError, TypeError):
                 print(f"Error: Cannot perform less_than comparison on non-numeric types: '{data}' and '{value}'")
                 return False

        # Add Python boolean literals
        true = True
        false = False
        # Add None for checking nulls
        null = None # Use Python's None keyword

        try:
            # Replace result.x with actual values from result using the updated _get_nested_value
            # Pattern needs to allow for [ ] characters in the path
            pattern = r'result\.([a-zA-Z0-9_\[\].]+)'

            def replace_result_ref(match):
                path = match.group(1) # This is the path *after* 'result.'
                value = self._get_nested_value(result, path)

                # Represent fetched value safely for eval
                if isinstance(value, (dict, list)):
                    # Use json.dumps for dicts/lists
                    return json.dumps(value, ensure_ascii=False)
                elif isinstance(value, str):
                     # Quote strings
                     return repr(value) # repr handles internal quotes and escapes
                elif value is None:
                    # Represent None as 'None' keyword in Python eval context
                    return 'None'
                else:
                    # Use repr for other types (int, float, bool, etc.)
                    return repr(value)


            # Perform replacements iteratively to handle nested replacements if needed
            # (e.g., if a replaced value contains text that looks like another variable)
            # However, the current pattern only matches result.x, so a single pass is fine.
            eval_condition = re.sub(pattern, replace_result_ref, condition)

            if self.verbose:
                 print(f"Evaluating condition string: {eval_condition}")

            # Execute the condition within a limited context
            # __builtins__={} restricts access to built-in functions
            context = {
                'contains': contains,
                'equal': equal,
                'greatThan': greater_than, # Corrected function name
                'lessThan': less_than,     # Corrected function name
                'result': result,          # Still provide result, though path access is replaced
                'true': True,
                'false': False,
                'null': None,
                'and': lambda a, b: a and b, # Basic boolean operators if needed
                'or': lambda a, b: a or b,
                'not': lambda a: not a
            }

            # Use eval cautiously. The replacement logic is key to safety.
            # Restricting builtins and providing only necessary functions in context helps.
            return eval(eval_condition, {"__builtins__": {}}, context)

        except Exception as e:
            print(f"Error evaluating condition '{condition}': {e}")
            # print(f"Condition string after replacement: '{eval_condition}'") # Debug replacement
            return False

    def execute_action(self, action: str, result: Dict[str, Any]) -> None:
        """Execute an action, such as setting an environment variable"""
        if not action or pd.isna(action):
            return

        # Split actions by semicolon or newline, and handle empty lines
        actions = [act.strip() for act in re.split(r'[;\n]', action) if act.strip()]

        for single_action in actions:
            # print(f"Executing action: {single_action}")

            # Pattern for variable assignment: $varname = result.path.to.value
            # Allow [ ] in the path after 'result.'
            pattern = r'\$([a-zA-Z0-9_]+)\s*=\s*result\.([a-zA-Z0-9_\[\].]+)'
            match = re.search(pattern, single_action)

            if match:
                var_name, result_path = match.groups()
                try:
                    # Use the shared _get_nested_value helper
                    value = self._get_nested_value(result, result_path)

                    if value is not None:
                        # Convert complex objects (dicts, lists) to JSON string
                        # Convert booleans to string "true" or "false"
                        # Convert None to string "null" or empty string? Let's use "null"
                        if isinstance(value, (dict, list)):
                            value_str = json.dumps(value, ensure_ascii=False)
                        elif isinstance(value, bool):
                            value_str = str(value).lower() # "true" or "false"
                        elif value is None:
                            value_str = "null"
                        else:
                            # Convert numbers and other primitives directly to string
                            value_str = str(value)

                        self.environment_vars[var_name] = value_str
                        print(f"Set environment variable ${var_name} = '{value_str}'")
                    else:
                        # If _get_nested_value returned None, the path was invalid or value was None
                        # We already print a warning in _get_nested_value, so just skip assignment.
                        print(f"Skipping assignment for ${var_name}: value not found at path '{result_path}' or was None.")

                except Exception as e:
                    print(f"Error executing action '{single_action}': {e}")
            else:
                print(f"Warning: Unrecognized action format: '{single_action}'")


    def parse_headers(self, header_text: str) -> Dict[str, str]:
        """Parse headers from various formats into a dictionary"""
        if not header_text or pd.isna(header_text):
            return {}

        header_text = self.replace_env_vars(str(header_text)) # Ensure text is a string
        headers = {}

        try:
            # Try parsing as JSON first
            # Replace single quotes with double quotes for JSON compatibility
            json_string = header_text.strip().replace("'", '"')
            # Attempt to parse as JSON object or array of objects
            parsed_data = json.loads(json_string)

            if isinstance(parsed_data, list):
                for item in parsed_data:
                    if isinstance(item, dict):
                        headers.update(item)
                    else:
                        print(f"Warning: Unexpected item type in headers list: {type(item).__name__}")
            elif isinstance(parsed_data, dict):
                headers.update(parsed_data)
            else:
                 print(f"Warning: Unexpected JSON format for headers: {type(parsed_data).__name__}")

            # Ensure all header values are strings
            return {k: str(v) for k, v in headers.items()}

        except json.JSONDecodeError as e:
            print(f"JSON Decode Error parsing headers '{header_text[:100]}...': {e}")
            # Fallback to a very basic parsing if JSON fails. This is less reliable.
            try:
                # Look for key: value pairs separated by comma or newline
                # This is a very simple parser and won't handle complex values or structures
                pairs = re.findall(r'([a-zA-Z0-9_-]+)\s*:\s*(.+?)(?:,|\n|$)', header_text)
                for key, value in pairs:
                    headers[key.strip()] = value.strip()
                if headers:
                    print(f"Warning: Fell back to basic parsing for headers '{header_text[:100]}...'. JSON format preferred.")
                    return headers
                else:
                    print(f"Error: Fallback parsing failed for headers '{header_text[:100]}...'")
                    return {}

            except Exception as e_fallback:
                print(f"Error during fallback parsing headers '{header_text[:100]}...': {e_fallback}")
                return {}
        except Exception as e:
             print(f"Unexpected error parsing headers '{header_text[:100]}...': {e}")
             return {}


    def execute_test_case(self, test_case: pd.Series) -> bool:
        """Execute a single test case"""
        test_name = str(test_case.get('test_case_name', 'Unnamed Test Case')) # Default name
        if pd.isna(test_case.get('api_path', None)):
             print(f"Skipping test case '{test_name}': 'api_path' is missing or empty.")
             self.results[test_name] = {"status": "Skipped", "reason": "'api_path' is missing or empty"}
             return False # Consider skipping as a non-failure for framework flow, but report it

        print(f"\nExecuting test case: {test_name}")

        # Check verbose flag and store it in self for reuse
        self.verbose = False
        if 'verbose' in test_case and not pd.isna(test_case['verbose']):
            verbose_value = str(test_case['verbose']).lower().strip()
            self.verbose = verbose_value in ('true', 'yes', '1')


        try:
            # Extract test case parameters
            api_path = self.replace_env_vars(str(test_case['api_path']))
            method = str(test_case.get('method', 'GET')).upper() # Default to GET

            # Parse query parameters
            query_params = {}
            if 'query_param' in test_case and not pd.isna(test_case['query_param']):
                query_param_list = self.parse_dict_list(test_case['query_param'])
                # parse_dict_list returns List[Dict], need to flatten it
                for param_dict in query_param_list:
                    query_params.update(param_dict)

            # Parse headers
            headers = self.parse_headers(test_case.get('inject_header', None))


            if self.verbose:
                print(f"Request URL: {api_path}")
                print(f"Request Method: {method}")
                print(f"Request Query Params: {query_params}")
                print(f"Request Headers: {headers}")

            # Parse body
            body_input = test_case.get('body', None)
            body = self.parse_json_body(body_input)
            if self.verbose and body is not None:
                 print(f"Request Body: {json.dumps(body, indent=2, ensure_ascii=False)}")


            # Execute API request
            response = requests.request(
                method=method,
                url=api_path,
                params=query_params,
                headers=headers,
                json=body # Use json=body for automatic Content-Type: application/json
                # For other content types, use data=... or files=... and set header manually
            )

            # Parse response
            response_json = None
            try:
                # Attempt to parse JSON response only if content type is json
                content_type = response.headers.get('Content-Type', '')
                if 'application/json' in content_type:
                   response_json = response.json()
                elif 'text/' in content_type or 'html' in content_type:
                   # For text responses, put the text in a 'text' key
                   response_json = {"text": response.text}
                else:
                   # For other content types (e.g., images, files), response.text might be garbled
                   # Store raw content or a placeholder
                    response_json = {"content": "<binary or non-text content>"}
            except json.JSONDecodeError:
                # If content type was json but decoding failed
                print(f"Warning: Could not decode JSON response for {test_name}")
                response_json = {"raw_response_text": response.text} # Store raw text for debugging
            except Exception as e:
                 print(f"Error processing response body for {test_name}: {e}")
                 response_json = {"error_processing_body": str(e), "raw_response_text": response.text}


            # Parse cookies from response using requests' built-in cookie handling + fallback
            cookies = self.parse_cookies(response)

            result = {
                "code": response.status_code,
                "body": response_json,
                "headers": dict(response.headers), # Convert requests headers to a standard dict
                "cookies": cookies,
                "elapsed_time_ms": response.elapsed.total_seconds() * 1000 # Add response time
            }

            # Store result for potential actions/validation in subsequent tests
            self.results[test_name] = {"status": "Executed", "result": result}


            # Print verbose output if requested
            if self.verbose:
                print("\n=== RESPONSE DETAILS ===")
                print(f"Status Code: {response.status_code}")
                print(f"Elapsed Time: {result['elapsed_time_ms']:.2f} ms")
                print("\nHeaders:")
                for header, value in result['headers'].items():
                    print(f"  {header}: {value}")
                print("\nCookies:")
                if result['cookies']:
                    for cookie_name, cookie_value in result['cookies'].items():
                        print(f"  {cookie_name}: {cookie_value}")
                else:
                    print("  No cookies received.")
                print("\nBody:")
                # Print body carefully depending on content type or size
                if isinstance(result['body'], dict):
                    if "raw_response_text" in result['body']:
                         print(result['body'].get("raw_response_text", "<empty>"))
                    elif "content" in result['body'] and result['body']['content'] == "<binary or non-text content>":
                         print("<Binary or non-text response body>")
                    else:
                         print(json.dumps(result['body'], indent=2, ensure_ascii=False))
                else:
                     print(result['body']) # Print as is if not a dict
                print("======================\n")


            # --- Validations ---
            test_passed = True

            # Validate response code
            expected_code = test_case.get('expect_response_code', None)
            if pd.notna(expected_code):
                try:
                    expected_code = int(expected_code)
                    if response.status_code != expected_code:
                        print(f"❌ Status Code Validation Failed: Expected {expected_code}, got {response.status_code}")
                        test_passed = False
                    else:
                        print(f"✅ Status Code Validation Passed: Expected and got {expected_code}")
                except ValueError:
                    print(f"Warning: Invalid value for 'expect_response_code': '{test_case['expect_response_code']}'")


            # Validate response body
            expected_body = test_case.get('expect_response_body', None)
            if pd.notna(expected_body):
                print(f"Validating response body using condition: {expected_body}")
                if not self.evaluate_condition(expected_body, result):
                    print(f"❌ Response body validation failed.")
                    test_passed = False
                else:
                    print(f"✅ Response body validation passed.")

            # Validate response headers
            expected_headers = test_case.get('expect_response_header', None)
            if pd.notna(expected_headers):
                print(f"Validating response headers using condition: {expected_headers}")
                if not self.evaluate_condition(expected_headers, result):
                    print(f"❌ Response header validation failed.")
                    test_passed = False
                else:
                    print(f"✅ Response header validation passed.")

            # --- Actions ---
            # Execute any actions regardless of validation success (e.g., saving a token even if a check fails)
            action = test_case.get('action', None)
            if pd.notna(action):
                self.execute_action(action, result)


            # Final test case status
            self.results[test_name]["status"] = "Passed" if test_passed else "Failed"

            if test_passed:
                print(f"✅ Test case '{test_name}' PASSED")
            else:
                print(f"❌ Test case '{test_name}' FAILED")

            return test_passed

        except requests.exceptions.RequestException as e:
            print(f"❌ Request Error executing test case '{test_name}': {e}")
            self.results[test_name] = {"status": "Failed", "error": f"Request Error: {e}"}
            return False
        except Exception as e:
            print(f"❌ Unexpected Error executing test case '{test_name}': {e}")
            self.results[test_name] = {"status": "Failed", "error": f"Unexpected Error: {e}"}
            # Print traceback for unexpected errors
            import traceback
            traceback.print_exc()
            return False


    def run_tests(self) -> Dict[str, Dict[str, Any]]:
        """Run all test cases from the Excel file"""
        try:
            xl = pd.ExcelFile(self.xlsx_path)
            sheet_names = xl.sheet_names
        except FileNotFoundError:
            print(f"Error: Excel file not found at '{self.xlsx_path}'")
            return {}
        except Exception as e:
            print(f"Error reading Excel file '{self.xlsx_path}': {e}")
            return {}


        if len(sheet_names) < 3:
            print("Error: Excel file must have at least 3 sheets (Environment, Setup, Tests...)")
            # Still attempt to load env and setup if they exist
            # return {}

        # First sheet is environment (already loaded in __init__)
        # Second sheet is setup
        setup_sheet_name = sheet_names[1] if len(sheet_names) > 1 else None
        if setup_sheet_name and setup_sheet_name.lower() != "setup":
            print(f"Warning: Second sheet is typically named 'Setup', found '{setup_sheet_name}'")
        elif not setup_sheet_name:
             print("Warning: No second sheet found for Setup.")


        # Execute setup sheet if it exists
        setup_success = True
        if setup_sheet_name:
            print(f"\n=== Running Setup Sheet: {setup_sheet_name} ===")
            try:
                setup_df = pd.read_excel(self.xlsx_path, sheet_name=setup_sheet_name)
                # Filter out rows that don't have a test case name
                setup_df = setup_df.dropna(subset=['test_case_name'])

                for index, test_case in setup_df.iterrows():
                    # Add index to test case name if it's not unique? Or rely on user unique names.
                    # Let's rely on user providing unique names or accept potential overwrites in self.results
                    success = self.execute_test_case(test_case)
                    if not success:
                        setup_success = False
                        print(f"❌ Setup test case failed ('{test_case.get('test_case_name', 'Unnamed Setup Case')}'). Remaining setup tests will be skipped, and main tests will not run.")
                        break # Stop executing setup tests on first failure

            except Exception as e:
                print(f"Error running Setup sheet '{setup_sheet_name}': {e}")
                setup_success = False


        # Execute remaining sheets (test cases) only if setup was successful or skipped
        if setup_success:
            for sheet_name in sheet_names[2:]:
                print(f"\n=== Running Test Sheet: {sheet_name} ===")
                try:
                    test_df = pd.read_excel(self.xlsx_path, sheet_name=sheet_name)
                    # Filter out rows that don't have a test case name
                    test_df = test_df.dropna(subset=['test_case_name'])

                    sheet_success = True
                    for index, test_case in test_df.iterrows():
                        success = self.execute_test_case(test_case)
                        # Decide if sheet execution stops on first failure or continues
                        # Current logic continues through the sheet
                        # If you want to stop sheet on first fail, add: if not success: sheet_success = False; break

                    if sheet_success and not test_df.empty:
                         print(f"✅ All executed tests in sheet '{sheet_name}' PASSED")
                    elif test_df.empty:
                         print(f"ℹ️ No test cases found with 'test_case_name' in sheet '{sheet_name}'")
                    # else: A failure within the sheet was already reported by execute_test_case

                except Exception as e:
                    print(f"Error running Test sheet '{sheet_name}': {e}")
                    # Continue to next sheet even if one sheet fails to load or run

        # Generate and print summary
        self._print_summary()

        return self.results

    def _print_summary(self) -> None:
        """Prints the test execution summary."""
        print("\n=== Test Results Summary ===")
        total_executed = 0
        passed_count = 0
        failed_count = 0
        skipped_count = 0

        for test_name, result_data in self.results.items():
            status = result_data.get("status", "Unknown")
            if status == "Executed": # Cases that ran but maybe didn't have validations resulting in Pass/Fail yet
                 print(f"- {test_name}: Status Unknown (executed but no final status)")
                 total_executed += 1 # Count executed tests
            elif status == "Passed":
                print(f"✅ {test_name}: PASSED")
                passed_count += 1
                total_executed += 1
            elif status == "Failed":
                error_msg = result_data.get("error", "")
                print(f"❌ {test_name}: FAILED {error_msg}")
                failed_count += 1
                total_executed += 1
            elif status == "Skipped":
                 reason = result_data.get("reason", "")
                 print(f"⏭️ {test_name}: SKIPPED {reason}")
                 skipped_count += 1
            else:
                print(f"? {test_name}: Status Unknown ({status})")
                total_executed += 1 # Assume executed if status is weird


        print("-" * 30)
        print(f"Total Test Cases Defined/Attempted: {len(self.results)}")
        print(f"Total Executed: {total_executed}")
        print(f"Passed: {passed_count}")
        print(f"Failed: {failed_count}")
        print(f"Skipped: {skipped_count}")
        print("-" * 30)


def run_test_suite(xlsx_path: str) -> None:
    """Run a test suite from an Excel file"""
    test_framework = APITestFramework(xlsx_path)
    # run_tests already prints summary
    test_framework.run_tests()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python api_test_framework.py <path_to_excel_file>")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    if not os.path.exists(xlsx_path):
        print(f"Error: File '{xlsx_path}' not found")
        sys.exit(1)

    run_test_suite(xlsx_path)