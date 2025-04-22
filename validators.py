import re
import json
import pandas as pd
from typing import Dict, List, Any, Union


class Validator:
    """Handles validation of API responses and executing actions"""

    def __init__(self, environment_vars: Dict[str, str]):
        self.environment_vars = environment_vars

    def validate_response(self, test_case: pd.Series, api_result_data: Dict[str, Any],
                          verbose: bool) -> Dict[str, Any]:
        """Validate API response against expected values"""
        validation_result = {
            "body_validation": "N/A",
            "header_validation": "N/A",
            "test_passed_validations": True,
            "details": ""
        }

        # Validate response code
        expected_code = test_case.get('expect_response_code', None)
        if pd.notna(expected_code):
            try:
                expected_code = int(expected_code)
                if api_result_data["code"] != expected_code:
                    validation_result[
                        "details"] += f"Status Code Failed (Expected: {expected_code}, Actual: {api_result_data['code']}). "
                    validation_result["test_passed_validations"] = False
            except ValueError:
                print(f"Warning: Invalid value for 'expect_response_code': '{test_case['expect_response_code']}'")
                validation_result[
                    "details"] += f"Invalid 'expect_response_code' value: '{test_case['expect_response_code']}'. "

        # Validate response body
        expected_body = test_case.get('expect_response_body', None)
        if pd.notna(expected_body):
            if not self.evaluate_condition(expected_body, api_result_data, verbose):
                validation_result["body_validation"] = "Failed"
                validation_result["details"] += f"Body Validation Failed ('{expected_body}'). "
                validation_result["test_passed_validations"] = False
            else:
                validation_result["body_validation"] = "Passed"

        # Validate response headers
        expected_headers = test_case.get('expect_response_header', None)
        if pd.notna(expected_headers):
            if not self.evaluate_condition(expected_headers, api_result_data, verbose):
                validation_result["header_validation"] = "Failed"
                validation_result["details"] += f"Header Validation Failed ('{expected_headers}'). "
                validation_result["test_passed_validations"] = False
            else:
                validation_result["header_validation"] = "Passed"

        return validation_result

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
                            return None
                    else:
                        return None
                else:
                    return None
            elif segment.isdigit():
                try:
                    index = int(segment)
                    if isinstance(current_value, list):
                        if 0 <= index < len(current_value):
                            current_value = current_value[index]
                        else:
                            return None
                    else:
                        return None
                except ValueError:
                    return None
            else:
                if isinstance(current_value, dict) and segment in current_value:
                    current_value = current_value.get(segment)
                elif isinstance(current_value, list) and segment == 'length':  # Basic list length access
                    current_value = len(current_value) if isinstance(current_value, list) else None
                    if current_value is None:
                        return None
                else:
                    return None

        return current_value

    def evaluate_condition(self, condition: str, result: Dict[str, Any], verbose: bool) -> bool:
        """Evaluate a condition against the result"""
        if not condition or pd.isna(condition):
            return True

        condition = self._replace_env_vars(str(condition))

        def contains(data, value):
            data_str = str(data) if data is not None else ""
            value_str = str(value) if value is not None else ""
            is_contained = value_str in data_str
            if not is_contained and verbose:
                print(f"  Condition Failed: Expected '{value_str}' to be contained in '{data_str[:200]}...'")
            return is_contained

        def equal(data, value):
            is_equal = data == value
            if not is_equal and verbose:
                print(
                    f"  Condition Failed: Expected '{value}' (type: {type(value).__name__}), "
                    f"Actual '{data}' (type: {type(data).__name__})")
            return is_equal

        def is_numeric(value):
            try:
                float(value)
                return True
            except (ValueError, TypeError):
                return False

        def greater_than(data, value):
            if not is_numeric(data) or not is_numeric(value):
                if verbose:
                    print(
                        f"  Condition Failed: Cannot perform greater_than comparison on non-numeric types: "
                        f"'{data}' and '{value}'")
                return False
            is_greater = float(data) > float(value)
            if not is_greater and verbose:
                print(f"  Condition Failed: Expected value > {value}, Actual {data}")
            return is_greater

        def less_than(data, value):
            if not is_numeric(data) or not is_numeric(value):
                if verbose:
                    print(
                        f"  Condition Failed: Cannot perform less_than comparison on non-numeric types: "
                        f"'{data}' and '{value}'")
                return False
            is_less = float(data) < float(value)
            if not is_less and verbose:
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

            if verbose:
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

            if verbose:
                print(f"  Condition '{condition}' evaluated to: {eval_result}")

            return bool(eval_result)

        except Exception as e:
            print(f"Error evaluating condition '{condition}': {e}")
            return False

    def _replace_env_vars(self, text: Union[str, Any]) -> Union[str, Any]:
        """Replace environment variables in text with their values"""
        if not isinstance(text, str):
            return text

        pattern = r'\$([a-zA-Z0-9_]+)'

        def replace_var(match):
            var_name = match.group(1)
            if var_name in self.environment_vars:
                return self.environment_vars[var_name]
            else:
                return match.group(0)  # Return original string if not found

        return re.sub(pattern, replace_var, text)

    def execute_action(self, action: str, result: Dict[str, Any]) -> None:
        """Execute an action, such as setting an environment variable"""
        if not action or pd.isna(action):
            return

        actions = [act.strip() for act in re.split(r'[;\n]', str(action)) if act.strip()]

        for single_action in actions:
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

                except Exception as e:
                    print(f"Error executing action '{single_action}': {e}")