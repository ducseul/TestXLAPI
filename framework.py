import pandas as pd
from typing import Dict, List, Any, Optional, Union
import traceback
import sys
import os

from config import ConfigLoader
from api_client import APIClient
from validators import Validator
from parsers import RequestParser
from reporters import ConsoleReporter, PDFReporter


class APITestFramework:
    def __init__(self, xlsx_path: str):
        """Initialize the API test framework with the path to an Excel file"""
        self.xlsx_path = xlsx_path
        self.environment_vars = {}
        # self.results will store detailed results per test case (used for final summary)
        self.results: Dict[str, Dict[str, Any]] = {}
        self.verbose = False  # Initialize verbose flag

        # Load configuration and environment variables
        self.config_loader = ConfigLoader(xlsx_path)
        self.environment_vars = self.config_loader.load_environment()

        # Initialize components
        self.parser = RequestParser(self.environment_vars)
        self.api_client = APIClient()
        self.validator = Validator(self.environment_vars)

        # Initialize reporters
        self.console_reporter = ConsoleReporter()
        self.pdf_reporter = PDFReporter()

    def execute_test_case(self, test_case: pd.Series, excel_sheet_name: str) -> Dict[str, Any]:
        """Execute a single test case and return detailed results."""
        test_name = str(test_case.get('test_case_name',
                                      f'Unnamed Test Case Row {test_case.name + 2}'))

        # Store result by sheet::name for the global summary
        full_test_name = f"{excel_sheet_name}::{test_name}"
        detailed_result = {
            "test_name": test_name,
            "status": "Skipped",
            "actual_code": "N/A",
            "body_validation": "N/A",
            "header_validation": "N/A",
            "details": "",
            "elapsed_time_ms": "N/A",
        }
        self.results[full_test_name] = detailed_result

        # Check if test case has api_path
        api_path_raw = test_case.get('api_path', None)
        if pd.isna(api_path_raw) or str(api_path_raw).strip() == '':
            print(f"\nSkipping test case '{test_name}' in sheet '{excel_sheet_name}': 'api_path' is missing or empty.")
            detailed_result["details"] = "'api_path' is missing or empty."
            return detailed_result

        # Check verbose flag specific to this test case row
        verbose_row = False
        if 'verbose' in test_case and not pd.isna(test_case['verbose']):
            verbose_value = str(test_case['verbose']).lower().strip()
            verbose_row = verbose_value in ('true', 'yes', '1')
        self.verbose = verbose_row

        try:
            # Parse request data
            api_path = self.parser.replace_env_vars(str(api_path_raw))
            method = str(test_case.get('method', 'GET')).upper()
            query_params = self.parser.parse_dict_list(test_case.get('query_param', None))
            headers = self.parser.parse_headers(test_case.get('inject_header', None))
            body = self.parser.parse_json_body(test_case.get('body', None))

            # Debug output if verbose
            if self.verbose:
                print(f"  Request URL: {api_path}")
                print(f"  Request Method: {method}")
                if query_params: print(f"  Request Query Params: {query_params}")
                if headers: print(f"  Request Headers: {headers}")
                if body is not None:
                    self.parser.print_body_preview(body)

            # Execute API request
            api_result_data = self.api_client.execute_request(method, api_path, query_params, headers, body)

            # Update result details
            detailed_result["actual_code"] = api_result_data["code"]
            detailed_result["elapsed_time_ms"] = api_result_data["elapsed_time_ms"]

            # Validate response
            validation_results = self.validator.validate_response(
                test_case, api_result_data, self.verbose
            )

            detailed_result.update(validation_results)

            # Determine final test status
            if validation_results["test_passed_validations"]:
                detailed_result["status"] = "Passed"
                print(f"✅ Test case '{test_name}' PASSED")
            else:
                detailed_result["status"] = "Failed"
                print(f"❌ Test case '{test_name}' FAILED")

            # Execute actions
            action = test_case.get('action', None)
            if pd.notna(action):
                self.validator.execute_action(action, api_result_data)

            return detailed_result

        except Exception as e:
            # Handle various errors
            if hasattr(e, '__module__') and e.__module__ == 'requests.exceptions':
                detailed_result["status"] = "Failed"
                detailed_result["details"] += f"Request Error: {e}"
                print(f"❌ Request Error executing test case '{test_name}': {e}")
            else:
                detailed_result["status"] = "Error"
                detailed_result["details"] += f"Unexpected Error: {e} - {traceback.format_exc()}"
                print(f"❌ Unexpected Error executing test case '{test_name}': {e}")
                traceback.print_exc()

            return detailed_result
        finally:
            self.verbose = False  # Reset verbose flag

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
                    setup_results_list.append(detailed_result)

                    if detailed_result["status"] in ["Failed", "Error"]:
                        setup_success = False
                        print(f"❌ Setup failed ('{test_case.get('test_case_name', 'Unnamed Setup Case')}'). "
                              f"Remaining setup tests and all main tests will be skipped.")
                        break

            except Exception as e:
                print(f"Error processing Setup sheet '{setup_sheet_name}': {e}")
                setup_success = False
            finally:
                # Print table for the setup sheet results
                self.console_reporter.print_sheet_results_table(setup_sheet_name, setup_results_list)

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
                        sheet_results_list.append(detailed_result)

                        if detailed_result["status"] in ["Failed", "Error"]:
                            sheet_has_failures = True

                    if not sheet_has_failures and not test_df.empty:
                        print(f"✅ All executed tests in sheet '{sheet_name}' PASSED")
                    elif test_df.empty:
                        print(f"ℹ️ No test cases found with 'test_case_name' in sheet '{sheet_name}'")

                except Exception as e:
                    print(f"Error processing Test sheet '{sheet_name}': {e}")
                    sheet_processing_error = e
                finally:
                    # Print table for the current sheet's results
                    self.console_reporter.print_sheet_results_table(sheet_name, sheet_results_list)
                    if sheet_processing_error:
                        print(f"‼️ Processing of sheet '{sheet_name}' encountered an error: {sheet_processing_error}")

        # --- Print Console Summary ---
        self.console_reporter.print_summary(self.results)

        return self.results

    def generate_pdf_report(self, output_path: str = "test_report.pdf"):
        """Generates a PDF report of the test results"""
        self.pdf_reporter.generate_report(self.results, output_path)