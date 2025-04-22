import re
import json
import pandas as pd
from typing import Dict, List, Any, Union, Optional
import json


class RequestParser:
    """Handles parsing of request components and variable replacements"""

    def __init__(self, environment_vars: Dict[str, str]):
        self.environment_vars = environment_vars

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
                return match.group(0)  # Return original string if not found

        return re.sub(pattern, replace_var, text)

    def parse_dict_list(self, text: str) -> Dict[str, str]:
        """Parse a string representation of a list of dictionaries"""
        if not text or pd.isna(text):
            return {}

        text = self.replace_env_vars(str(text))  # Ensure text is a string
        result_dict = {}

        try:
            # Attempt to parse as JSON first
            json_string = text.strip().replace("'", '"')
            parsed_list = json.loads(json_string)

            # Convert list of dictionaries to a single dict
            if isinstance(parsed_list, list):
                for item in parsed_list:
                    if isinstance(item, dict):
                        result_dict.update(item)
            elif isinstance(parsed_list, dict):
                result_dict = parsed_list

            return result_dict

        except json.JSONDecodeError:
            # Fallback to a simpler parsing if JSON fails
            try:
                # Attempt to handle [{'key', 'value'}, {'key2', 'value2'}] or [{'key': 'value'}]
                items = re.findall(r'\{(.*?)\}', text)
                for item in items:
                    item = item.strip()
                    if not item: continue

                    parts = [p.strip("'\" ") for p in item.split(',', 1)]
                    if len(parts) == 2:
                        key, value = parts
                        result_dict[key] = value
                    elif ':' in item:
                        parts = [p.strip("'\" ") for p in item.split(':', 1)]
                        if len(parts) == 2:
                            key, value = parts
                            result_dict[key] = value

                return result_dict
            except Exception as e_fallback:
                print(f"Error during fallback parsing dictionary list '{text[:100]}...': {e_fallback}")
                return {}
        except Exception as e:
            print(f"Unexpected error parsing dictionary list '{text[:100]}...': {e}")
            return {}

    def parse_json_body(self, body_text: str) -> Any:
        """Parse the body text as JSON"""
        if not body_text or pd.isna(body_text):
            return None

        body_text = self.replace_env_vars(str(body_text))  # Ensure text is a string

        try:
            return json.loads(body_text)
        except json.JSONDecodeError:
            return None
        except Exception as e:
            print(f"Unexpected error parsing JSON body '{body_text[:100]}...': {e}")
            return None

    def parse_headers(self, header_text: str) -> Dict[str, str]:
        """Parse headers from various formats into a dictionary"""
        if not header_text or pd.isna(header_text):
            return {}

        header_text = self.replace_env_vars(str(header_text))

        # Use the parse_dict_list method for consistency
        return self.parse_dict_list(header_text)

    def print_body_preview(self, body: Any) -> None:
        """Print a preview of the request body"""
        if body is not None:
            body_print = json.dumps(body, indent=2, ensure_ascii=False)
            print(f"  Request Body: {body_print[:500]}{'...' if len(body_print) > 500 else ''}")