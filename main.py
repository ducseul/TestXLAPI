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

    # You can also examine the results programmatically
    for test_name, result in results.items():
        print(f"Test: {test_name}, Status Code: {result['code']}")


if __name__ == "__main__":
    run_example()