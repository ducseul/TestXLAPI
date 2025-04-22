from collections import OrderedDict, defaultdict
from typing import Dict, List, Any
import pandas as pd

from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch


class ConsoleReporter:
    """Handles console reporting of test results"""

    def print_sheet_results_table(self, sheet_name: str, results_list: List[Dict[str, Any]]) -> None:
        """Prints the results for a single sheet in a formatted table, including Response Time."""
        if not results_list:
            print(f"\nNo test cases executed in sheet '{sheet_name}'.")
            return

        print(f"\n--- Results for Sheet: {sheet_name} ---")

        # Define columns and their corresponding keys in the result dictionary
        columns = OrderedDict([
            ("Test Name", "test_name"),
            ("Response Time", "elapsed_time_ms"),
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

    def print_combined_sheet_results(self, sheet_name: str, results: Dict[str, Dict[str, Any]]) -> None:
        """Prints the combined results across multiple cycles for a single sheet."""
        # Filter results for the current sheet
        sheet_results = {k: v for k, v in results.items() if k.startswith(f"{sheet_name}::")}

        if not sheet_results:
            print(f"\nNo test cases with multiple cycles executed in sheet '{sheet_name}'.")
            return

        print(f"\n--- Combined Results for Sheet: {sheet_name} (Multiple Cycles) ---")

        # Define columns for statistics output
        columns = OrderedDict([
            ("Test Name", "test_name"),
            ("Status", "status"),
            ("Success Rate", "success_rate"),
            ("Min Time", "min_time_ms"),
            ("Max Time", "max_time_ms"),
            ("Avg Time", "avg_time_ms"),
            ("StdDev", "std_dev_ms"),
        ])

        # Define maximum width for the test name to keep the table manageable
        max_name_width = 30

        # Calculate column widths dynamically based on headers and content
        col_widths = {header: len(header) for header in columns.keys()}
        for result in sheet_results.values():
            for header, key in columns.items():
                value = result.get(key, '')

                # Format time values for display and width calculation
                if key in ["min_time_ms", "max_time_ms", "avg_time_ms", "median_time_ms", "std_dev_ms"]:
                    if isinstance(value, (int, float)):
                        value_str = f"{value:.2f} ms"
                    else:
                        value_str = str(value)
                else:
                    value_str = str(value)

                if header == "Test Name" and len(value_str) > max_name_width:
                    value_str = value_str[:max_name_width]  # Truncate for width calculation

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
        for full_test_name, result in sorted(sheet_results.items()):
            row_data = []
            for header, key in columns.items():
                value = result.get(key, '')

                # Format for different column types
                if key in ["min_time_ms", "max_time_ms", "avg_time_ms", "median_time_ms", "std_dev_ms"]:
                    if isinstance(value, (int, float)):
                        value_str = f"{value:.2f} ms"
                    else:
                        value_str = "N/A"
                else:
                    value_str = str(value)

                # Truncate test name if too long
                if header == "Test Name" and len(value_str) > max_name_width:
                    value_str = value_str[:max_name_width - 3] + "..."

                row_data.append(value_str.ljust(padded_widths[header] - col_padding))

            print("| " + " | ".join(row_data) + " |")

        print("-" * len(header_row))  # Match separator length to header row

    def print_summary(self, results: Dict[str, Dict[str, Any]]) -> None:
        """Prints the test execution summary based on results dictionary."""
        print("\n=== Overall Test Run Summary ===")
        total_attempted = len(results)
        passed_count = 0
        failed_count = 0
        error_count = 0
        skipped_count = 0
        unknown_count = 0

        if total_attempted == 0:
            print("No test cases were attempted.")
            return

        # Sort results by sheet and test name for consistent output
        sorted_test_names = sorted(results.keys())

        for full_test_name in sorted_test_names:
            result_data = results[full_test_name]
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


class PDFReporter:
    """Handles PDF report generation for test results"""

    def generate_report(self, results: Dict[str, Dict[str, Any]], output_path: str = "test_report.pdf",
                        cycles: int = 1) -> None:
        """Generates a PDF report of the test results with per-sheet insights,
           failed/errored tests, and slowest tests."""
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()

        # Add a main title
        cycle_text = f" ({cycles} Cycles)" if cycles > 1 else ""
        elements.append(Paragraph(f"API Test Report{cycle_text}", styles['Title']))
        elements.append(Spacer(1, 0.5 * inch))

        # --- Overall Summary ---
        elements.append(Paragraph("Overall Test Run Summary", styles['Heading1']))
        total_attempted_overall = len(results)
        passed_count_overall = 0
        failed_count_overall = 0
        error_count_overall = 0
        skipped_count_overall = 0

        for result_data in results.values():
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
        sorted_full_test_names = sorted(results.keys())
        for full_test_name in sorted_full_test_names:
            sheet_name, test_name = full_test_name.split("::", 1)
            results_by_sheet[sheet_name].append((test_name, results[full_test_name]))

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

            for _, result_data in sheet_results_list:
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

            # Add Performance Statistics Section if multiple cycles were run
            if cycles > 1:
                elements.append(Paragraph("Performance Statistics", styles['Heading3']))
                elements.append(Spacer(1, 0.1 * inch))

                # Define multi-cycle statistics table headers
                perf_stats_data = [['Test Case', 'Success Rate', 'Min Time', 'Max Time', 'Avg Time', 'Std Dev']]

                # Add each test case's performance data
                for test_name, result_data in sheet_results_list:
                    success_rate = result_data.get("success_rate", "N/A")
                    min_time = f"{result_data.get('min_time_ms', 'N/A'):.2f} ms" if isinstance(
                        result_data.get('min_time_ms'), (int, float)) else "N/A"
                    max_time = f"{result_data.get('max_time_ms', 'N/A'):.2f} ms" if isinstance(
                        result_data.get('max_time_ms'), (int, float)) else "N/A"
                    avg_time = f"{result_data.get('avg_time_ms', 'N/A'):.2f} ms" if isinstance(
                        result_data.get('avg_time_ms'), (int, float)) else "N/A"
                    std_dev = f"{result_data.get('std_dev_ms', 'N/A'):.2f} ms" if isinstance(
                        result_data.get('std_dev_ms'), (int, float)) else "N/A"

                    perf_stats_data.append([test_name, success_rate, min_time, max_time, avg_time, std_dev])

                # Create and add the table
                perf_stats_table = Table(perf_stats_data,
                                         colWidths=[2.5 * inch, 1 * inch, 0.8 * inch, 0.8 * inch, 0.8 * inch,
                                                    0.8 * inch])

                # Style for the performance stats table
                perf_table_style = [
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                ]

                perf_stats_table.setStyle(perf_table_style)
                elements.append(perf_stats_table)
                elements.append(Spacer(1, 0.4 * inch))

            # Add Detailed Results for this Sheet
            elements.append(Paragraph("Detailed Results (This Sheet)", styles['Heading3']))
            elements.append(Spacer(1, 0.1 * inch))

            for test_name, result_data in sheet_results_list:
                status = result_data.get("status", "Unknown")

                # For multiple cycles, include cycle-specific info
                if cycles > 1:
                    cycles_run = result_data.get("cycles_run", 0)
                    passed_count = result_data.get("passed_count", 0)
                    failed_count = result_data.get("failed_count", 0)
                    error_count = result_data.get("error_count", 0)

                    text_color = colors.black
                    if status == "Passed":
                        text_color = colors.green
                    elif status in ["Failed", "Error"]:
                        text_color = colors.red

                    elements.append(Paragraph(f"<b>Test Case:</b> {test_name}", styles['Normal']))
                    elements.append(
                        Paragraph(f"<b>Overall Status:</b> <font color='{text_color}'>{status}</font>",
                                  styles['Normal']))
                    elements.append(
                        Paragraph(
                            f"<b>Cycles Run:</b> {cycles_run} | <b>Passed:</b> {passed_count} | <b>Failed:</b> {failed_count} | <b>Errors:</b> {error_count}",
                            styles['Normal']))

                    # Add response time statistics
                    if isinstance(result_data.get("avg_time_ms"), (int, float)):
                        elements.append(
                            Paragraph(
                                f"<b>Response Times:</b> Min: {result_data.get('min_time_ms', 'N/A'):.2f} ms | "
                                f"Max: {result_data.get('max_time_ms', 'N/A'):.2f} ms | "
                                f"Avg: {result_data.get('avg_time_ms', 'N/A'):.2f} ms | "
                                f"StdDev: {result_data.get('std_dev_ms', 'N/A'):.2f} ms",
                                styles['Normal']
                            )
                        )

                    details = result_data.get("details", "")
                    if details:
                        details_str = str(details) if not pd.isna(details) else ""
                        elements.append(Paragraph(f"<b>Details:</b> {details_str}", styles['Normal']))
                else:
                    # For single cycle, show standard info
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
        failed_errored_tests_items = [
            (full_name, result) for full_name, result in results.items()
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

                # For multiple cycles, include failure stats
                if cycles > 1:
                    cycles_run = result_data.get("cycles_run", 0)
                    passed_count = result_data.get("passed_count", 0)
                    failed_count = result_data.get("failed_count", 0)
                    error_count = result_data.get("error_count", 0)
                    failure_rate = result_data.get("failure_rate", "N/A")

                    elements.append(Paragraph(f"<b>Test Case:</b> {sheet_name}::{test_name}", styles['Normal']))
                    elements.append(
                        Paragraph(f"<b>Overall Status:</b> <font color='{colors.red}'>{status}</font>",
                                  styles['Normal']))
                    elements.append(
                        Paragraph(
                            f"<b>Cycles Run:</b> {cycles_run} | <b>Passed:</b> {passed_count} | <b>Failed:</b> {failed_count} | <b>Errors:</b> {error_count} | <b>Failure Rate:</b> {failure_rate}",
                            styles['Normal']))
                else:
                    actual_code = result_data.get("actual_code", "N/A")

                    # Use the extracted sheet_name and test_name for the title
                    elements.append(Paragraph(f"<b>Test Case:</b> {sheet_name}::{test_name}", styles['Normal']))
                    elements.append(
                        Paragraph(f"<b>Status:</b> <font color='{colors.red}'>{status}</font>", styles['Normal']))
                    elements.append(Paragraph(f"<b>Response Code:</b> {actual_code}", styles['Normal']))

                details = result_data.get("details", "")
                if details:
                    details_str = str(details) if not pd.isna(details) else ""
                    elements.append(Paragraph(f"<b>Details:</b> {details_str}", styles['Normal']))
                elements.append(Spacer(1, 0.25 * inch))

        # --- Section for Slowest Tests ---
        if cycles > 1:
            # For multiple cycles, show tests with highest average times
            tests_with_avg_time = [
                (full_name, result) for full_name, result in results.items()
                if isinstance(result.get("avg_time_ms"), (int, float))
            ]

            # Sort by average response time in descending order
            slowest_tests_items = sorted(tests_with_avg_time,
                                         key=lambda item: item[1].get("avg_time_ms", 0),
                                         reverse=True)
        else:
            tests_with_time_items = [
                (full_name, result) for full_name, result in results.items()
                if isinstance(result.get("elapsed_time_ms"), (int, float))
            ]

            # Sort by elapsed time in descending order using the value from the dictionary
            slowest_tests_items = sorted(tests_with_time_items,
                                         key=lambda item: item[1].get("elapsed_time_ms", 0),
                                         reverse=True)

        # Define how many slowest tests to show (e.g., top 10)
        top_n_slowest = 10
        slowest_tests_to_show_items = slowest_tests_items[:top_n_slowest]

        if slowest_tests_to_show_items:
            elements.append(PageBreak())
            time_type = "Average" if cycles > 1 else ""
            elements.append(Paragraph(f"Top {top_n_slowest} {time_type} Slowest Test Cases", styles['Heading1']))
            elements.append(Spacer(1, 0.25 * inch))

            for full_test_name, result_data in slowest_tests_to_show_items:
                # Extract sheet_name and test_name from the full_test_name key
                sheet_name, test_name = full_test_name.split("::", 1)

                status = result_data.get("status", "Unknown")

                # Use the extracted sheet_name and test_name for the title
                elements.append(Paragraph(f"<b>Test Case:</b> {sheet_name}::{test_name}", styles['Normal']))

                if cycles > 1:
                    # For multiple cycles, show statistical info
                    min_time = result_data.get("min_time_ms", "N/A")
                    max_time = result_data.get("max_time_ms", "N/A")
                    avg_time = result_data.get("avg_time_ms", "N/A")
                    std_dev = result_data.get("std_dev_ms", "N/A")

                    if isinstance(avg_time, (int, float)):
                        elements.append(
                            Paragraph(
                                f"<b>Response Times:</b> Min: {min_time:.2f} ms | "
                                f"Max: {max_time:.2f} ms | "
                                f"Avg: {avg_time:.2f} ms | "
                                f"StdDev: {std_dev:.2f} ms",
                                styles['Normal']
                            )
                        )
                    else:
                        elements.append(Paragraph(f"<b>Response Times:</b> {avg_time}", styles['Normal']))
                else:
                    elapsed_time_ms = result_data.get("elapsed_time_ms", "N/A")

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