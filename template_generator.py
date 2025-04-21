import pandas as pd
import os


def create_template_xlsx(output_path="api_test_template.xlsx"):
    """Create a template Excel file for API testing"""
    # Create a new Excel writer
    with pd.ExcelWriter(output_path) as writer:
        # Environment sheet
        env_data = {
            'Key': ['base_url', 'username', 'password', 'client_id', 'client_secret'],
            'Value': ['https://api.example.com', 'testuser', 'testpass', 'client123', 'secret456']
        }
        env_df = pd.DataFrame(env_data)
        env_df.to_excel(writer, sheet_name='Environment', index=False)

        # Setup sheet
        setup_columns = [
            'test_case_name', 'api_path', 'query_param', 'method', 'inject_header',
            'body', 'expect_response_code', 'expect_response_body',
            'expect_response_header', 'action', 'verbose'
        ]
        setup_data = [{
            'test_case_name': 'Get Authentication Token',
            'api_path': '$base_url/auth/token',
            'query_param': '',
            'method': 'POST',
            'inject_header': "[{'Content-Type', 'application/json'}]",
            'body': '{"client_id": "$client_id", "client_secret": "$client_secret"}',
            'expect_response_code': 200,
            'expect_response_body': "contains(result.body, 'access_token')",
            'expect_response_header': '',
            'action': '$accessToken = result.body.access_token',
            'verbose': 'false'
        }]
        setup_df = pd.DataFrame(setup_data)
        setup_df = setup_df[setup_columns]  # Ensure columns are in the correct order
        setup_df.to_excel(writer, sheet_name='Setup', index=False)

        # Add verbose=true to one of the journey test cases to demonstrate usage
        user_journey1_data = [{
            'test_case_name': 'Get User Profile',
            'api_path': '$base_url/api/users/profile',
            'query_param': '',
            'method': 'GET',
            'inject_header': "[{'Authorization', 'Bearer $accessToken'}, {'Content-Type', 'application/json'}]",
            'body': '',
            'expect_response_code': 200,
            'expect_response_body': "contains(result.body, 'id') and contains(result.body, 'email')",
            'expect_response_header': '',
            'action': '$userId = result.body.id',
            'verbose': 'true'  # Enable verbose output for this test
        }]
        user_journey1_df = pd.DataFrame(user_journey1_data)
        user_journey1_df = user_journey1_df[setup_columns]  # Use the same columns as setup
        user_journey1_df.to_excel(writer, sheet_name='User Journey 1', index=False)

        # User Journey 2: Create and Delete Resource
        user_journey2_data = [
            {
                'test_case_name': 'Create Resource',
                'api_path': '$base_url/api/resources',
                'query_param': '',
                'method': 'POST',
                'inject_header': "[{'Authorization', 'Bearer $accessToken'}, {'Content-Type', 'application/json'}]",
                'body': '{"name": "Test Resource", "owner_id": "$userId"}',
                'expect_response_code': 201,
                'expect_response_body': "contains(result.body, 'id')",
                'expect_response_header': '',
                'action': '$resourceId = result.body.id'
            },
            {
                'test_case_name': 'Delete Resource',
                'api_path': '$base_url/api/resources/$resourceId',
                'query_param': '',
                'method': 'DELETE',
                'inject_header': "[{'Authorization', 'Bearer $accessToken'}]",
                'body': '',
                'expect_response_code': 204,
                'expect_response_body': '',
                'expect_response_header': '',
                'action': ''
            }
        ]
        user_journey2_df = pd.DataFrame(user_journey2_data)
        user_journey2_df = user_journey2_df[setup_columns]
        user_journey2_df.to_excel(writer, sheet_name='User Journey 2', index=False)

    print(f"Template created at: {os.path.abspath(output_path)}")

if __name__ == "__main__":
    create_template_xlsx()