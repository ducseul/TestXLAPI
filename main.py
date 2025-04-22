import argparse
import template_generator
from framework import APITestFramework


def run_example(test_file, report_name='report', cycles=1):
    test_framework = APITestFramework(test_file, cycles=cycles)
    test_framework.run_tests()
    test_framework.generate_pdf_report(f"{report_name}.pdf")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="""
        This framework provides a powerful yet simple way to automate API testing using familiar Excel spreadsheets. 
        By understanding the file structure, column definitions, and supported features like environment variables, 
        evaluation conditions, and actions, you can create comprehensive test suites for your APIs. 
        The detailed console output helps you debug and understand the results of each test case.
    """)
    parser.add_argument("test_file", help="Path to the Excel (.xlsx) test file to use or generate")
    parser.add_argument("--report-name", default="report", help="Name of the PDF report file (without extension)")
    parser.add_argument("--generate-template", action="store_true", help="Only generate a test template Excel file and exit")
    parser.add_argument('--cycle', type=int, default=1,
                        help='Number of cycles to run each test sheet. Provides statistical analysis when > 1.')
    args = parser.parse_args()

    if args.generate_template:
        template_generator.create_template_xlsx(args.test_file)
        print(f"Template generated: {args.test_file}")
    else:
        run_example(args.test_file, args.report_name, args.cycle)