import os
import re
import json
import requests
import pandas as pd
from typing import Dict, List, Any, Optional, Union


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
            env_df = pd.read_excel(self.xlsx_path, sheet_name=0)
            # Take only first two columns as key-value pairs
            for _, row in env_df.iterrows():
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
                    key = str(row.iloc[0]).strip()
                    value = str(row.iloc[1])
                    self.environment_vars[key] = value
            print(f"Loaded {len(self.environment_vars)} environment variables")
        except Exception as e:
            print(f"Error loading environment variables: {e}")

    def replace_env_vars(self, text: str) -> str:
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
                return match.group(0)

        return re.sub(pattern, replace_var, text)

    def parse_dict_list(self, text: str) -> List[Dict[str, str]]:
        """Parse a string representation of a list of dictionaries"""
        if not text or pd.isna(text):
            return []

        text = self.replace_env_vars(text)

        try:
            # For input like: [{'abc', 'xyz'}]
            # Convert to proper JSON format first
            text = text.replace("'", '"').replace("{", '{"').replace("}", '"}').replace(", ", '", "').replace('{"',
                                                                                                              '{').replace(
                '"}', '}')
            return json.loads(text)
        except json.JSONDecodeError:
            try:
                # Fall back to a simpler parsing for [{'key1', 'value1'}, {'key2', 'value2'}]
                result = []
                items = text.strip('[]').split('}, {')
                for item in items:
                    item = item.strip('{}')
                    parts = item.split(', ')
                    if len(parts) == 2:
                        key = parts[0].strip("'\"")
                        value = parts[1].strip("'\"")
                        result.append({key: value})
                return result
            except Exception as e:
                print(f"Error parsing dictionary list '{text}': {e}")
                return []

    def parse_json_body(self, body_text: str) -> Optional[Dict[str, Any]]:
        """Parse the body text as JSON"""
        if not body_text or pd.isna(body_text):
            return None

        body_text = self.replace_env_vars(body_text)

        try:
            return json.loads(body_text)
        except json.JSONDecodeError:
            print(f"Error parsing JSON body: {body_text}")
            return None

    def evaluate_condition(self, condition: str, result: Dict[str, Any]) -> bool:
        """Evaluate a condition against the result"""
        if not condition or pd.isna(condition):
            return True

        condition = self.replace_env_vars(condition)

        # Create a safe evaluation context
        def contains(data, value):
            is_contained = False
            if isinstance(data, dict):
                is_contained = value in str(data)
            elif isinstance(data, list):
                is_contained = value in data
            else:
                is_contained = value in str(data)

            if not is_contained:
                print(f"  Expected: '{value}' to be contained in")
                print(f"  Actual: {data}")
            return is_contained

        def equal(data, value):
            is_equal = data == value
            if not is_equal:
                print(f"  Expected: '{value}'")
                print(f"  Actual: '{data}'")
            return is_equal

        def greater_than(data, value):
            is_greater = data > value
            if not is_greater:
                print(f"  Expected: value > {value}")
                print(f"  Actual: {data}")
            return is_greater

        def less_than(data, value):
            is_less = data < value
            if not is_less:
                print(f"  Expected: value < {value}")
                print(f"  Actual: {data}")
            return is_less

        # Add JavaScript-like boolean literals for convenience
        true = True
        false = False

        try:
            # Replace result.x with actual values from result
            pattern = r'result\.([a-zA-Z0-9_\.]+)'

            def get_nested_value(obj, path):
                parts = path.split('.')
                for part in parts:
                    if isinstance(obj, dict) and part in obj:
                        obj = obj[part]
                    elif isinstance(obj, list) and part.isdigit():
                        obj = obj[int(part)]
                    else:
                        return None
                return obj

            def replace_result_ref(match):
                path = match.group(1)
                value = get_nested_value(result, path)
                if isinstance(value, (dict, list)):
                    return json.dumps(value)
                return repr(value)

            eval_condition = re.sub(pattern, replace_result_ref, condition)

            # Execute the condition
            context = {
                'contains': contains,
                'equal': equal,
                'greatThan': greater_than,
                'lessThan': less_than,
                'result': result,
                'true': True,
                'false': False
            }

            return eval(eval_condition, {"__builtins__": {}}, context)
        except Exception as e:
            print(f"Error evaluating condition '{condition}': {e}")
            return False

    def execute_action(self, action: str, result: Dict[str, Any]) -> None:
        """Execute an action, such as setting an environment variable"""
        if not action or pd.isna(action):
            return

        # Split actions by semicolon or newline
        actions = re.split(r';|\n', action)
        for single_action in actions:
            single_action = single_action.strip()
            if not single_action:
                continue

            # Pattern for variable assignment: $varname = result.path.to.value
            pattern = r'\$([a-zA-Z0-9_]+)\s*=\s*result\.([a-zA-Z0-9_\.]+)'
            match = re.search(pattern, single_action)

            if match:
                var_name, result_path = match.groups()
                try:
                    # Get value from result using the path
                    parts = result_path.split('.')
                    value = result
                    for part in parts:
                        if isinstance(value, dict) and part in value:
                            value = value[part]
                        elif isinstance(value, list) and part.isdigit():
                            value = value[int(part)]
                        else:
                            print(f"Warning: Path {result_path} not found in result")
                            value = None
                            break

                    if value is not None:
                        # Convert complex objects to string representation
                        if isinstance(value, (dict, list)):
                            value = json.dumps(value)
                        else:
                            value = str(value)

                        self.environment_vars[var_name] = value
                        print(f"Set environment variable ${var_name} = {value}")
                except Exception as e:
                    print(f"Error executing action '{single_action}': {e}")

    def execute_test_case(self, test_case: pd.Series) -> bool:
        """Execute a single test case"""
        try:
            test_name = test_case['test_case_name']
            print(f"\nExecuting test case: {test_name}")

            # Check verbose flag
            verbose = False
            if 'verbose' in test_case and not pd.isna(test_case['verbose']):
                verbose_value = str(test_case['verbose']).lower().strip()
                verbose = verbose_value in ('true', 'yes', '1')

            # Extract test case parameters
            api_path = self.replace_env_vars(str(test_case['api_path']))
            method = str(test_case['method']).upper()

            # Parse query parameters
            query_params = {}
            if 'query_param' in test_case and not pd.isna(test_case['query_param']):
                query_param_list = self.parse_dict_list(test_case['query_param'])
                for param in query_param_list:
                    for key, value in param.items():
                        query_params[key] = value

            # Parse headers
            headers = {}
            if 'inject_header' in test_case and not pd.isna(test_case['inject_header']):
                header_list = self.parse_dict_list(test_case['inject_header'])
                for header in header_list:
                    for key, value in header.items():
                        headers[key] = value

            # Parse body
            body = None
            if 'body' in test_case and not pd.isna(test_case['body']):
                body = self.parse_json_body(test_case['body'])

            # Execute API request
            response = requests.request(
                method=method,
                url=api_path,
                params=query_params,
                headers=headers,
                json=body
            )

            # Parse response
            try:
                response_json = response.json()
            except:
                response_json = {"text": response.text}

            result = {
                "code": response.status_code,
                "body": response_json,
                "headers": dict(response.headers)
            }

            self.results[test_name] = result

            # Print verbose output if requested
            if verbose:
                print("\n=== RESPONSE DETAILS ===")
                print(f"Status Code: {response.status_code}")
                print("\nHeaders:")
                for header, value in response.headers.items():
                    print(f"  {header}: {value}")
                print("\nBody:")
                print(json.dumps(response_json, indent=2, ensure_ascii=False))
                print("======================\n")

            # Validate response code
            expected_code = test_case.get('expect_response_code', None)
            if expected_code and not pd.isna(expected_code):
                expected_code = int(expected_code)
                if response.status_code != expected_code:
                    print(f"❌ Expected status code {expected_code}, got {response.status_code}")
                    return False

            # Validate response body
            expected_body = test_case.get('expect_response_body', None)
            if expected_body and not pd.isna(expected_body):
                print(f"Validating response body: {expected_body}")
                if not self.evaluate_condition(expected_body, result):
                    print(f"❌ Response body validation failed: {expected_body}")
                    print(f"Full response body: {json.dumps(result['body'], indent=2, ensure_ascii=False)}")
                    return False

            # Validate response headers
            expected_headers = test_case.get('expect_response_header', None)
            if expected_headers and not pd.isna(expected_headers):
                if not self.evaluate_condition(expected_headers, result):
                    print(f"❌ Response header validation failed: {expected_headers}")
                    return False

            # Execute any actions
            action = test_case.get('action', None)
            if action and not pd.isna(action):
                self.execute_action(action, result)

            print(f"✅ Test case {test_name} passed")
            return True

        except Exception as e:
            print(f"❌ Error executing test case: {e}")
            return False

    def run_tests(self) -> Dict[str, Dict[str, Any]]:
        """Run all test cases from the Excel file"""
        # Load all sheet names
        xl = pd.ExcelFile(self.xlsx_path)
        sheet_names = xl.sheet_names

        if len(sheet_names) < 3:
            print("Error: Excel file must have at least 3 sheets")
            return {}

        # First sheet is environment (already loaded)
        # Second sheet is setup
        setup_sheet = sheet_names[1]
        if setup_sheet.lower() != "setup":
            print(f"Warning: Second sheet should be named 'Setup', found '{setup_sheet}'")

        # Execute setup sheet
        print(f"\n=== Running Setup Sheet: {setup_sheet} ===")
        setup_df = pd.read_excel(self.xlsx_path, sheet_name=setup_sheet)
        setup_success = True

        for _, test_case in setup_df.iterrows():
            if pd.notna(test_case.get('test_case_name', None)):
                success = self.execute_test_case(test_case)
                if not success:
                    setup_success = False
                    print(f"❌ Setup failed, continuing with test cases. All test case halted.")
                    return {}

        # Execute remaining sheets (test cases)
        for sheet_name in sheet_names[2:]:
            print(f"\n=== Running Test Sheet: {sheet_name} ===")
            test_df = pd.read_excel(self.xlsx_path, sheet_name=sheet_name)

            sheet_success = True
            for _, test_case in test_df.iterrows():
                if pd.notna(test_case.get('test_case_name', None)):
                    success = self.execute_test_case(test_case)
                    if not success:
                        sheet_success = False
                        print(f"❌ Test case failed, skipping remaining tests in sheet '{sheet_name}'")
                        break

            if sheet_success:
                print(f"✅ All tests in sheet '{sheet_name}' passed")

        return self.results


def run_test_suite(xlsx_path: str) -> None:
    """Run a test suite from an Excel file"""
    test_framework = APITestFramework(xlsx_path)
    results = test_framework.run_tests()

    # Print summary
    print("\n=== Test Results Summary ===")
    total_tests = len(results)
    passed_tests = sum(1 for result in results.values() if result.get('code', 0) == result.get('expected_code', 0))
    print(f"Total tests: {total_tests}")
    print(f"Passed tests: {passed_tests}")
    print(f"Failed tests: {total_tests - passed_tests}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python api_test_framework.py <path_to_excel_file>")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    if not os.path.exists(xlsx_path):
        print(f"Error: File '{xlsx_path}' not found")
        sys.exit(1)

    run_test_suite(xlsx_path)