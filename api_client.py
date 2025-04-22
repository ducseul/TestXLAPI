import requests
import json
from typing import Dict, Any, List, Optional
import traceback


class APIClient:
    """Handles API requests and response processing"""

    def __init__(self):
        self.timeout = 15  # Default timeout in seconds

    def execute_request(self, method: str, url: str, params: Dict[str, str],
                        headers: Dict[str, str], body: Any) -> Dict[str, Any]:
        """Execute an API request and return processed response data"""
        try:
            response = requests.request(
                method=method,
                url=url,
                params=params,
                headers=headers,
                json=body,  # Use json=body for automatic Content-Type: application/json
                timeout=self.timeout
            )

            return self._process_response(response)

        except requests.exceptions.Timeout:
            raise requests.exceptions.Timeout("Request timed out")
        except requests.exceptions.RequestException as e:
            raise e
        except Exception as e:
            traceback.print_exc()
            raise e

    def _process_response(self, response) -> Dict[str, Any]:
        """Process the API response into a standardized format"""
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
                response_json = {
                    "content_type": content_type,
                    "content_preview": response.text[:100] + "..." if len(response.text) > 100 else response.text
                }
        except json.JSONDecodeError:
            response_json = {"decoding_error": "Failed to decode JSON", "raw_response_text": response_body_text}
        except Exception as e:
            response_json = {"processing_error": str(e), "raw_response_text": response_body_text}

        cookies = self._parse_cookies(response)

        return {
            "code": response.status_code,
            "body": response_json,
            "headers": dict(response.headers),
            "cookies": cookies,
            "elapsed_time_ms": response.elapsed.total_seconds() * 1000
        }

    def _parse_cookies(self, response) -> Dict[str, str]:
        """Extract cookies from response using requests' built-in cookiejar"""
        cookies = {}
        if response and hasattr(response, 'cookies'):
            for cookie in response.cookies:
                cookies[cookie.name] = cookie.value
        return cookies