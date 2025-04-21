from exl_api_fw import APITestFramework

# # First create your template Excel file
# import template_generator
#
# template_generator.create_template_xlsx("my_api_tests.xlsx")


# Now you can edit the Excel file manually with your API test details
# After editing, run the tests:
def run_example():
    test_framework = APITestFramework("my_api_tests.xlsx")
    results = test_framework.run_tests()

    # # You can also examine the results programmatically
    # print("\n=== Test Results Summary ===")
    # total_executed = 0
    # passed_count = 0
    # failed_count = 0
    # skipped_count = 0
    #
    # for test_name, result_data in results.items():
    #     status = result_data.get("status", "Unknown")
    #     if status == "Executed":  # Cases that ran but maybe didn't have validations resulting in Pass/Fail yet
    #         print(f"- {test_name}: Status Unknown (executed but no final status)")
    #         total_executed += 1  # Count executed tests
    #     elif status == "Passed":
    #         print(f"✅ {test_name}: PASSED")
    #         passed_count += 1
    #         total_executed += 1
    #     elif status == "Failed":
    #         error_msg = result_data.get("error", "")
    #         print(f"❌ {test_name}: FAILED {error_msg}")
    #         failed_count += 1
    #         total_executed += 1
    #     elif status == "Skipped":
    #         reason = result_data.get("reason", "")
    #         print(f"⏭️ {test_name}: SKIPPED {reason}")
    #         skipped_count += 1
    #     else:
    #         print(f"? {test_name}: Status Unknown ({status})")
    #         total_executed += 1  # Assume executed if status is weird
    #
    # print("-" * 30)
    # print(f"Total Test Cases Defined/Attempted: {len(results)}")
    # print(f"Total Executed: {total_executed}")
    # print(f"Passed: {passed_count}")
    # print(f"Failed: {failed_count}")
    # print(f"Skipped: {skipped_count}")
    # print("-" * 30)

if __name__ == "__main__":
    run_example()