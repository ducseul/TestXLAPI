-----

# API Test Framework User Manual

## 1\. Introduction

This document provides a guide for using the Python-based API Test Framework. This framework allows you to define and execute API test cases using a structured Excel file, manage environment-specific variables, perform validations on API responses, and chain requests by using data from one response in subsequent requests. Test results are printed directly to the console in a clear, tabular format, sheet by sheet.

The test cases, environment variables, and setup steps are all defined within a single `.xlsx` Excel file.

## 2\. Prerequisites

  * Python 3.6 or higher installed.
  * `pip` (Python package installer) available.

## 3\. Setup and Installation

1.  Save the provided Python code as a file named `api_test_framework.py`.
2.  Open your terminal or command prompt.
3.  Navigate to the directory where you saved the file.
4.  Install the required Python libraries using pip:
    ```bash
    pip install pandas requests openpyxl xlrd
    ```
    *(Note: `openpyxl` and `xlrd` are needed by `pandas` to read `.xlsx` and `.xls` files respectively, depending on your pandas version and file type, even though the framework doesn't directly use `openpyxl` for writing to the file anymore).*

## 4\. Excel File Structure

The test framework requires a specific structure for the Excel file (`.xlsx`). It must contain at least two sheets, but typically will have three or more.

### 4.1. Sheet 1: Environment Variables

  * The **first sheet** in the workbook is used to define environment variables.
  * It should contain key-value pairs.
  * The framework reads the **first two columns** of this sheet.
  * Column A is treated as the **Variable Key**.
  * Column B is treated as the **Variable Value**.
  * **No header row is strictly required by the code**, but including one (e.g., "Key" and "Value" in row 1) is highly recommended for readability. The framework starts reading data from the first row regardless of content.
  * Any row where Column A is empty will be ignored.

**Example (Sheet 1):**

| Key      | Value                     |
| :------- | :------------------------ |
| BASE\_URL | [https://api.example.com](https://www.google.com/search?q=https://api.example.com)   |
| API\_KEY  | abcdef123456              |
| USERNAME | testuser                  |
| PASSWORD | testpass                  |
| \# Comment | This row is ignored      |

These variables can then be referenced in your test cases using the syntax `$VARIABLE_KEY`.

### 4.2. Sheet 2: Setup Test Cases

  * The **second sheet** is reserved for **Setup** test cases.
  * This sheet uses the **same column structure** as the main test sheets (described [on Section #5 below](#5-test-case-definition-columns).)
  * Tests in this sheet are executed **before** any tests in subsequent sheets.
  * The primary purpose of the Setup sheet is to perform actions like:
      * Logging in to get an authentication token or session cookie.
      * Creating prerequisite data.
      * Setting environment variables based on the response of a setup request (e.g., `$TOKEN = result.body.access_token`).
  * **If any test case in the Setup sheet fails or encounters an error, the execution of all subsequent test sheets will be skipped.**

### 4.3. Sheet 3 onwards: Main Test Cases

  * All sheets starting from the **third sheet** are treated as **main test suites**.
  * Each of these sheets uses the **same column structure** as the Setup sheet.
  * Tests within a sheet are executed sequentially from top to bottom.
  * Execution proceeds to the next sheet even if a test case in the current sheet fails (unless the failure occurred in the Setup sheet).

## 5\. Test Case Definition (Columns)

Each row in the Setup sheet and all main test sheets defines a single test case. The columns expected are as follows (column order in the Excel file does **not** strictly matter, but using the order below is recommended for readability):

| Column Name          | Required? | Description                                                                                                                                                              | Expected Format / Example Content                                                                                                |
| :------------------- | :-------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------- |
| `test_case_name`     | Yes       | A unique name for the test case within its sheet. Used in output and result tracking.                                                                                    | `Get User Profile`, `Create New Item`                                                                                            |
| `api_path`           | Yes       | The API endpoint path (can be a full URL or relative if `BASE_URL` is in env vars). Environment variables (`$VAR_NAME`) are replaced here.                               | `https://api.example.com/users/$USER_ID` or `$BASE_URL/items`                                                                    |
| `method`             | Yes       | The HTTP method for the request.                                                                                                                                         | `GET`, `POST`, `PUT`, `DELETE`, `PATCH`                                                                                          |
| `query_param`        | No        | Query parameters for the request. Should be a string representation of a Python list of dictionaries `[{'key':'value'}, {'key2':'value2'}]`. Env vars are replaced. | `[{'page': '1'}, {'limit': '$PAGE_SIZE'}]`                                                                                     |
| `inject_header`      | No        | Headers to inject into the request. Should be a string representation of a Python dictionary `{'Header-Name':'Value'}` or a list of dictionaries `[{'Header-Name':'Value'}, ...]`. Env vars are replaced. JSON format is preferred. | `[{'Content-Type': 'application/json'}, {'Authorization': 'Bearer $TOKEN'}]` or `{'Accept': 'application/xml'}`              |
| `body`               | No        | The request body. Should be a string representation of a JSON object or array. Env vars are replaced.                                                                  | `{"name": "New Item", "value": 123}` or `[{"id": 1}, {"id": 2}]`                                                                  |
| `verbose`            | No        | If `true`, `yes`, or `1`, prints detailed request and response information for this test case to the console.                                                          | `true`, `yes`, `false`, `no`, `1`, `0`                                                                                           |
| `expect_response_code` | No      | The expected HTTP status code (integer). If provided, the test will fail if the actual code does not match.                                                          | `200`, `201`, `400`                                                                                                              |
| `expect_response_body` | No      | A Python-like boolean expression to validate the response body. See "Evaluation Conditions" section. Env vars are replaced.                                            | `result.body.status == 'success' and contains(result.body.data, 'something')`                                                  |
| `expect_response_header`| No      | A Python-like boolean expression to validate the response headers. See "Evaluation Conditions" section. Env vars are replaced.                                         | `result.headers.Content-Type == 'application/json' and greatThan(result.headers.Content-Length, 100)`                          |
| `action`             | No        | An action to perform after the request, typically setting environment variables from the response. See "Actions" section. Multiple actions can be separated by `;` or newline. Env vars are replaced. | `$USER_ID = result.body.user.id ; $SESSION = result.cookies.JSESSIONID`                                                         |

## 6\. Running the Tests

1.  Save your test definitions in an Excel file (e.g., `my_api_tests.xlsx`) following the structure described above.

2.  Open your terminal or command prompt.

3.  Navigate to the directory where you saved `api_test_framework.py`.

4.  Run the script, providing the path to your Excel file as a command-line argument:

    ```bash
    python api_test_framework.py path/to/your/my_api_tests.xlsx
    ```

The script will load the environment variables, execute the Setup sheet, then execute each subsequent test sheet, printing results per sheet and a final summary to the console.

## 7\. Supported Features

### 7.1. Environment Variables

  * Variables defined in the first sheet (`$VAR_KEY = value`) are stored and can be used in almost any string field in your test cases: `api_path`, `query_param`, `inject_header`, `body`, `expect_response_body`, `expect_response_header`, and `action`.
  * Syntax: Prefix the variable key with a dollar sign (`$`). Example: `$BASE_URL`.
  * Variable names are case-sensitive based on Python dictionary lookup.
  * Variables set using the `action` column during a test run overwrite variables with the same key loaded from the Environment sheet or previously set by another action. These changes are available immediately for subsequent test cases within the same run.

### 7.2. Data Parsing

The framework automatically attempts to parse specific string inputs from your Excel sheet:

  * `query_param`: Parses a string representation of a Python list of dictionaries (e.g., `[{'key':'value'}]`).
  * `inject_header`: Parses a string representation of a Python dictionary (e.g., `{'Header':'Value'}`) or a list of dictionaries (e.g., `[{'Header':'Value'}]`). JSON format is preferred, but a basic fallback parser exists.
  * `body`: Parses a string representation of a JSON object or array.

**Important:** Ensure your input strings in the Excel cells are correctly formatted Python/JSON strings (using appropriate quotes and delimiters) after environment variable replacement. Using single quotes `'` within the Excel cell is often handled, but valid JSON syntax (`"key": "value"`) is the most reliable.

### 7.3. Accessing API Response Data (`result` object)

Within `expect_response_body`, `expect_response_header`, and `action` columns, you have access to a `result` object which represents the response from the current API call.

The `result` object has the following top-level attributes:

  * `result.code`: The HTTP status code as an integer (e.g., `200`).
  * `result.headers`: A dictionary containing the response headers. Header names are generally accessed using dot notation (see below).
  * `result.cookies`: A dictionary containing the parsed response cookies. Cookie names are accessed using dot notation (see below).
  * `result.body`: The parsed response body. The structure depends on the `Content-Type` of the response:
      * If `Content-Type` is `application/json`, this will be a Python dictionary or list.
      * If `Content-Type` is text-based (like `text/plain`, `text/html`), this will be a dictionary like `{"text": "..."}` where `"text"` contains the raw response text.
      * For other content types (e.g., images), this will be a dictionary providing content type and a preview, like `{"content_type": "image/png", "content_preview": "..."}`.
      * If JSON parsing fails, this will be a dictionary containing error information and the raw text, like `{"decoding_error": "...", "raw_response_text": "..."}`.
  * `result.elapsed_time_ms`: The duration of the request in milliseconds as a float.

**Navigating Nested Data (within `result.body`, `result.headers`, `result.cookies`):**

You can traverse into the `body`, `headers`, or `cookies` using dot notation for dictionary keys and bracket notation `[index]` for list elements.

  * **Accessing Dictionary Keys (Dot Notation):**
    Use `.` followed by the key name. This works for nested dictionaries within the body, and for accessing values in the `headers` and `cookies` dictionaries.

      * `result.body.some_key`
      * `result.body.level1.level2_key`
      * `result.headers.Content-Type` (Accessing the value of the `Content-Type` header using the header name as the key)
      * `result.cookies.JSESSIONID` (Accessing the value of the `JSESSIONID` cookie using the cookie name as the key)

  * **Accessing List Elements (Bracket Notation):**
    If the value is a list (e.g., `result.body` is a list, or a dictionary value is a list), use `[index]` where `index` is a 0-based integer.

      * `result.body[0]` (If the body is a list)
      * `result.body.list_of_items[0]` (Accessing the first item in a list stored under `list_of_items`)
      * `result.body.list_of_items[1].nested_key` (Accessing a key within an object that is the second item in a list)

  * **Combined Navigation:**
    Chain dot and bracket notation to access deeply nested data.

      * `result.body.data.users[0].address.city`

  * **List Length:**
    A special `.length` property is available on list references to get the size of the list.

      * `result.body.list_of_items.length`

  * **Accessing Raw Text Body:**
    If the response body was not JSON but text, access the raw text using `result.body.text`.

**Examples:**

Assuming a response body like:

```json
{
  "status": "success",
  "data": {
    "id": 123,
    "username": "testuser",
    "roles": ["admin", "editor"],
    "address": {
      "street": "123 Main St",
      "city": "Anytown"
    }
  },
  "message": "User retrieved"
}
```

And headers like:

```
Content-Type: application/json
X-Request-ID: abc-123
Set-Cookie: SESSIONID=xyz; Path=/
```

And a cookie named `JSESSIONID` with value `xyz`.

  * `result.code` will be `200` (or other status code).
  * `result.body.status` will be the string `'success'`.
  * `result.body.data.id` will be the integer `123`.
  * `result.body.data.username` will be the string `'testuser'`.
  * `result.body.data.roles` will be the list `['admin', 'editor']`.
  * `result.body.data.roles[0]` will be the string `'admin'`.
  * `result.body.data.address.city` will be the string `'Anytown'`.
  * `result.body.data.roles.length` will be the integer `2`.
  * `result.message` will be the string `'User retrieved'`.
  * `result.headers.Content-Type` will be the string `'application/json'`.
  * `result.cookies.JSESSIONID` will be the string `'xyz'`.

### 7.4. Evaluation Conditions (`expect_response_body`, `expect_response_header`)

You can write Python-like boolean expressions to validate complex aspects of the API response body or headers.

  * **Where Used:** `expect_response_body` and `expect_response_header` columns.
  * **Syntax:** A single expression that evaluates to `True` or `False`.
  * **Accessing Results:** Use the `result` object and its paths as described in the "Accessing API Response Data" section above.
  * **Supported Functions:**
      * `contains(data, value)`: Checks if the string representation of `value` is a substring within the string representation of `data`.
      * `equal(data, value)`: Checks if `data` is equal to `value` (`data == value`).
      * `greatThan(data, value)`: Checks if `float(data) > float(value)`.
      * `lessThan(data, value)`: Checks if `float(data) < float(value)`.
  * **Supported Literals:**
      * `true`: Python boolean `True`.
      * `false`: Python boolean `False`.
      * `null`: Python `None`. Use this to check for null values (`result.body.optional_field is null`).
  * **Supported Operators:** Standard Python comparison (`==`, `!=`, `>`, `<`, `>=`, `<=`) and logical operators (`and`, `or`, `not`).

**Examples (`expect_response_body` or `expect_response_header`):**

  * `result.code == 200` (Also checkable directly with `expect_response_code`)
  * `equal(result.body.status, 'success')`
  * `contains(result.body.message, 'User created successfully')`
  * `result.body.user.id is not null and greatThan(result.body.user.age, 18)`
  * `result.body.items.length > 0 and equal(result.body.items[0].name, 'First Item')`
  * `result.headers.Content-Type == 'application/json'`
  * `equal(result.cookies.SESSIONID, $EXPECTED_SESSION_ID)`

### 7.5. Actions (`action`)

Actions allow you to perform simple assignments based on the API response, typically to update environment variables for use in subsequent test cases.

  * **Where Used:** `action` column.
  * **Syntax:** `$VARIABLE_NAME = result.path.to.value`.
  * Multiple actions can be specified by separating them with a semicolon (`;`) or placing them on separate lines in the Excel cell.
  * **Accessing Results:** Uses the same `result` object and path navigation rules as described in the "Accessing API Response Data" section above.
  * The value retrieved from the `result.path` is converted to a string (JSON string for complex objects, "true"/"false" for booleans, "null" for None) before being stored in the environment variable dictionary (`self.environment_vars`).
  * Setting an environment variable will overwrite any existing variable with the same name for the remainder of the test run.

**Examples (`action`):**

  * `$AUTH_TOKEN = result.body.access_token`
  * `$USER_ID = result.body.data.user.id`
  * `$SESSION_COOKIE = result.cookies.JSESSIONID`
  * `$CREATED_ITEM_ID = result.body.id; $STATUS_MESSAGE = result.body.message` (Multiple actions)
  * `$FIRST_ERROR_CODE = result.body.errors[0].code`
  * `$TOTAL_ITEMS = result.body.items.length`
  * `$RAW_TEXT_BODY = result.body.text`

## 8\. Console Output

The framework prints its progress and results to the console:

  * It indicates when environment variables are loaded.
  * It prints the test case name being executed.
  * If `verbose` is enabled for a test case, it prints detailed request and response information.
  * It indicates whether each test case PASSED or FAILED based on validations.
  * After executing all test cases on a sheet, it prints a formatted **table** summarizing the results for that specific sheet.
  * Finally, after all sheets are processed, it prints an **overall summary** count of passed, failed, error, and skipped tests across the entire run.

**Table Columns:**

  * **Test Name:** The name of the test case.
  * **Status:** Execution status (Passed, Failed, Skipped, Error).
  * **Code:** The actual HTTP status code received.
  * **Body Val:** Status of the `expect_response_body` validation (Passed, Failed, N/A).
  * **Header Val:** Status of the `expect_response_header` validation (Passed, Failed, N/A).
  * **Details:** Contains information about validation failures, request errors, or skipping reasons. This column is truncated in the console table display (max \~80 characters) for readability.

## 9\. Tips for Writing Test Cases

  * **Start Simple:** Create basic GET requests with status code checks first.
  * **Use the Setup Sheet:** Place login requests and variable assignments for authentication tokens/cookies in the Setup sheet. This keeps your main tests cleaner.
  * **Leverage Environment Variables:** Use `$VAR_NAME` for base URLs, credentials, and dynamic data obtained from previous requests.
  * **Organize Sheets:** Group related API endpoints or test scenarios into separate sheets.
  * **Meaningful Names:** Give your test cases clear and descriptive names.
  * **Use `verbose`:** Enable `verbose` for specific test cases when developing or debugging to see the full request/response details and the exact structure of `result.body`, `result.headers`, etc.
  * **Validate Effectively:** Use `expect_response_code`, `expect_response_body`, and `expect_response_header` to create robust checks. Don't just check the status code; verify important data in the response body and required headers/cookies.
  * **Correct Data Formats:** Pay close attention to the required string formats for `query_param`, `inject_header`, and `body` in your Excel cells (lists of dicts, JSON). Use online JSON validators if needed.
  * **Handle Dynamic Data:** Use the `action` column to extract IDs, tokens, or other data from responses and store them in environment variables for use in subsequent requests.
  * **LLM Prompt Helping** You can use prompt on next section after import this docs to LLM like `Chatgpt`, `Gemini`... and let them help writing test case.
## 10\. Troubleshooting

  * **`FileNotFoundError`:** Ensure the path to the Excel file provided on the command line is correct.
  * **Library Import Errors (`ModuleNotFoundError`):** Make sure you have installed the required libraries (`pandas`, `requests`, `openpyxl`, `xlrd`) using `pip`.
  * **Excel Reading Errors:** Check if the Excel file is closed. Ensure it's a `.xlsx` file.
  * **`JSONDecodeError`:** The string you provided for `query_param`, `inject_header`, or `body` is not valid JSON (after environment variable replacement). Check your syntax carefully (commas, colons, quotes).
  * **`Warning: Path part '...' not found...`:** Your `result.path` expression in an evaluation condition or action is incorrect. Check the structure of the actual API response (using `verbose`) and ensure your path matches the keys and array indices (`result.body.user.address[0].city`). Remember `result.cookies.CookieName` and `result.headers.Header-Name` for cookies/headers (using dot notation). Use `result.body.text` for raw text bodies.
  * **Validation Failed (in output table/summary):** Check the 'Details' column in the output table for specific reasons (Status Code mismatch, Body/Header condition evaluated to False). Use `verbose` for that test case to inspect the full response data and verify your path and condition logic.
  * **Setup Failed, skipping rest:** A test case in the second (Setup) sheet failed or encountered an error. Fix that test case before running the full suite.

## 11\. Conclusion

This framework provides a powerful yet simple way to automate API testing using familiar Excel spreadsheets. By understanding the file structure, column definitions, and supported features like environment variables, evaluation conditions, and actions, you can create comprehensive test suites for your APIs. The detailed console output helps you debug and understand the results of each test case.

-----

# Prompt Library for Generating Excel Test Cases with an LLM

This library provides prompts you can use with a large language model (LLM) like ChatGPT or Gemini to help you create the content for your API Test Framework Excel file.

**Important First Step:**

Start your conversation with the LLM by providing the entire user manual document. This gives the LLM the necessary context about the Excel structure, columns, variable syntax, evaluation functions, and action formats.

**Prompt 1: Initializing the LLM with the Manual**

```markdown
I will provide you with a user manual for an API test framework. Please read it carefully to understand the required Excel file structure, column definitions, data formats (especially for query_param, inject_header, and body), environment variable syntax ($VAR_NAME), **how to access response data using the `result` object and paths (like result.body.path, result.headers.Name, result.cookies.Name, result.body.list[index]),** **how to write evaluation conditions (using supported functions like `contains`, `equal`, `greatThan`, `lessThan`, supported literals like `true`, `false`, `null`, and Python operators on the result data),** and **how to write actions (using the $VAR_NAME = result.path syntax).**

Once you have read and understood the manual, please confirm you are ready to help me generate content for the Excel sheets based on this document, strictly following all formatting and syntax rules defined in the manual.

[Paste the entire content of the API Test Framework User Manual here]
```

**Prompt 2: Generating the Environment Sheet Content**

After the LLM confirms it understands the manual, use this prompt to generate the first sheet.

```markdown
Based on the user manual you just read, generate the content for the **first sheet** of the Excel file, which is for **Environment Variables**.

I need the following environment variables defined:
- Base URL for the API: [Your Base URL, e.g., https://api.myservice.com/v1]
- An API Key: [Your API Key, e.g., abcdef1234567890]
- A default username: [e.g., testuser]
- A default password: [e.g., password123]
- [List any other required environment variables]

Follow the two-column structure (Key, Value) as specified in the manual. Format the output as a Markdown table.
```

**Prompt 3: Generating the Setup Sheet Content (Explicit Evaluation/Action)**

Use this prompt to generate the content for the second sheet (Setup). Describe the login or initial setup steps you need, being explicit about validations and actions using the manual's syntax.

```markdown
Based on the user manual, generate the content for the **second sheet** of the Excel file, which is the **Setup sheet**.

I need a test case here to perform a login request.
- The test case name should be "Login and Get Token".
- It should use the environment variable for the base URL and target the path `/auth/login`.
- The method should be POST.
- The request body should be a JSON object containing the username and password, referencing the environment variables `$USERNAME` and `$PASSWORD`.
- It should expect a 200 OK status code.
- **For the `expect_response_body` column, write the evaluation expression based on the manual's rules (Section 7.4) to validate that the response body contains a field named `access_token` and that its value is not null.** (Example expression: `result.body.access_token is not null`).
- **For the `action` column, write the action expression based on the manual's rules (Section 7.5) to save the value of `access_token` from the response body into an environment variable named `$AUTH_TOKEN`.** (Example action: `$AUTH_TOKEN = result.body.access_token`). Multiple actions should be separated by `;` or newline in the cell content.
- Please also include the 'verbose' column set to 'true' for this test case to help with initial debugging.

Follow the complete column structure and data formats for test cases specified in the manual. Format the output as a Markdown table.
```

**Prompt 4: Generating Content for a Main Test Sheet (Explicit Evaluation/Action)**

Use this prompt to generate content for a main test sheet (Sheet 3 or later). Describe the API calls you want to test, being explicit about validations and actions using the manual's syntax.

```markdown
Based on the user manual, generate the content for a **main test sheet**. Let's name this sheet "[Your Sheet Name, e.g., User Profile Tests]".

I need the following test cases defined, **making sure to write the `expect_response_body`, `expect_response_header`, and `action` column content using the `result` object, paths, functions, and syntax rules from the manual (Sections 7.3, 7.4, 7.5):**

1.  **Test Case Name:** "Get My Profile"
    * **Description:** Retrieve the profile of the authenticated user.
    * **Path:** Use the base URL variable and the path `/users/me`.
    * **Method:** GET.
    * **Headers:** Include a header `Authorization` with the value `Bearer $AUTH_TOKEN` (referencing the variable set in Setup). Include `Content-Type` as `application/json`.
    * **Expected Code:** 200.
    * **Body Validation:** Write the expression to validate that the response body contains `id`, `username`, and `email` fields, and that `result.body.username` equals the original `$USERNAME` environment variable. (Example: `result.body.id is not null and result.body.username is not null and result.body.email is not null and equal(result.body.username, $USERNAME)`).
    * **Header Validation:** Write the expression to validate that the `Content-Type` header is `application/json`. (Example: `result.headers.Content-Type == 'application/json'`).
    * **Action:** No specific action needed for this test.
    * **Details:** [Add any specific details about expected response structure if needed]

2.  **Test Case Name:** "Update My Profile Email"
    * **Description:** Update the email address for the authenticated user.
    * **Path:** Use the base URL variable and the path `/users/me`.
    * **Method:** PUT.
    * **Headers:** Include the same `Authorization` and `Content-Type` headers as the previous test.
    * **Body:** A JSON object `{"email": "new.email@example.com"}`.
    * **Expected Code:** 200.
    * **Body Validation:** Write the expression to validate that `result.body.email` equals `"new.email@example.com"`. (Example: `equal(result.body.email, 'new.email@example.com')`).
    * **Action:** No specific action needed for this test.
    * **Details:** [Add any specific details about expected response structure if needed]

3.  [Describe your next test case following the same structure, explicitly asking the LLM to generate the validation/action content using the `result` object and rules from the manual]

Follow the complete column structure and data formats for test cases specified in the manual. Format the output as a Markdown table.
```

**Prompt 5: Adding a New Test Case to an Existing Sheet (Explicit Evaluation/Action)**

If you've already generated a sheet's content and want to add a new test case, use this prompt.

```markdown
Using the Markdown table content you previously generated for the "[Sheet Name]" sheet, please add **one more test case** to it.

The new test case details are, **making sure to write the `expect_response_body`, `expect_response_header`, and `action` column content using the `result` object, paths, functions, and syntax rules from the manual (Sections 7.3, 7.4, 7.5):**
-   **Test Case Name:** "[New Test Case Name]"
-   **Description:** [Describe the new test scenario]
-   **Path:** [API path, referencing env vars if applicable]
-   **Method:** [HTTP Method]
-   **Headers:** [Headers needed, referencing env vars if applicable]
-   **Query Params:** [Query parameters needed]
-   **Body:** [Request body, JSON string, referencing env vars if applicable]
-   **Expected Code:** [Expected status code]
-   **Body Validation:** Write the expression for this validation using the manual's rules.
-   **Header Validation:** Write the expression for this validation using the manual's rules.
-   **Action:** Write the action expression(s) for this test using the manual's rules.
-   **Details:** [Any notes]

Please provide the **complete, updated** content for the "[Sheet Name]" sheet as a Markdown table again, including the new test case.
```

**Prompt 6: Modifying an Existing Test Case (Explicit Evaluation/Action)**

If you need to change a specific test case you've already generated, use this prompt.

```markdown
Using the Markdown table content you previously generated for the "[Sheet Name]" sheet, please modify the test case named "[Test Case Name]".

Apply the following changes, **making sure to write the updated `expect_response_body`, `expect_response_header`, and `action` column content using the `result` object, paths, functions, and syntax rules from the manual (Sections 7.3, 7.4, 7.5):**
-   Change the Method to: [New Method, e.g., PUT]
-   Add the following Header: [e.g., {'X-Correlation-ID': 'abc'}]
-   Update the Body Validation to: [New validation expression using the manual's rules]
-   Update the Action to: [New action expression(s) using the manual's rules, e.g., `$NEW_VAR = result.body.something`]
-   Remove the Header Validation currently defined.
-   [List any other specific changes]

Please provide the **complete, updated** content for the "[Sheet Name]" sheet as a Markdown table again, with the modifications applied to "[Test Case Name]".
```

**Prompt 7: Converting a cURL Command to a Test Case Row**

Use this prompt when you have a `curl` command (e.g., exported from Postman or copied from browser developer tools) and you want to quickly generate the corresponding row content for a main test sheet in your Excel file. The LLM will extract the core request details and format them according to the manual.

Based on the user manual you previously read, I will provide you with a `curl` command. Your task is to convert this `curl` command into a **single row** for a main test sheet in the Excel file structure defined in the manual.

Parse the `curl` command to extract the following information:
1.  **Method:** Identify the HTTP method (e.g., GET, POST, PUT, DELETE) usually specified with `-X` or implied (GET is often implied if `-X` is absent and no body is sent).
2.  **URL:** Extract the request URL.
3.  **Headers:** Extract all headers specified with `-H` flags. Format these headers into the string format required for the `inject_header` column (a Python dictionary string like `{'Header-Name': 'Value'}` or a list of dictionaries `[{'Header-Name': 'Value'}, {'Another-Header': 'Value2'}]`). Handle multiple `-H` flags.
4.  **Body:** Extract the request body specified with `-d` or `--data` flags. Format this body content into the string format required for the `body` column (a JSON string if it looks like JSON, or a plain string otherwise).

For the output, generate a single row following the standard test case column structure from the manual (`test_case_name`, `api_path`, `method`, `query_param`, `inject_header`, `body`, `verbose`, `expect_response_code`, `expect_response_body`, `expect_response_header`, `action`, `details`).

* Set the `method` column to the extracted method.
* Set the `api_path` column to the extracted URL.
* Set the `inject_header` column to the formatted headers string.
* Set the `body` column to the formatted body string (or leave blank if no body).
* Create a placeholder value for the `test_case_name` column (e.g., "Generated from cURL YYYY-MM-DD HH:MM:SS").
* Leave the `query_param`, `verbose`, `expect_response_code`, `expect_response_body`, `expect_response_header`, `action`, and `details` columns blank or as 'N/A'. You, the user, will need to fill these in later based on the API's expected behavior.

Generate the output as a Markdown table containing just this single row, including the header row that lists all the column names from the manual.

Here is the `curl` command I want you to convert:

```bash
# Paste your curl command here.
# Example:
# curl -X POST '[https://api.example.com/users](https://api.example.com/users)' \
# -H 'Content-Type: application/json' \
# -H 'Authorization: Bearer YOUR_TOKEN' \
# -d '{"username":"newuser","password":"securepassword"}'
```
```

---

**Explanation of this Prompt:**

1.  **Context:** It reminds the LLM of the manual and the goal (converting `curl` to a specific Excel row format).
2.  **Parsing Instructions:** It explicitly lists the parts of the `curl` command the LLM needs to identify (`-X`, URL, `-H`, `-d`/`--data`).
3.  **Formatting Instructions:** It tells the LLM how to format the extracted headers and body into the *string representations* required for the `inject_header` and `body` columns in the Excel file (`[{'Name':'Value'}]` or `{'Name':'Value'}` for headers, JSON string for body).
4.  **Column Mapping:** It specifies which `curl` part goes into which Excel column (`method`, `api_path`, `inject_header`, `body`).
5.  **Placeholder Instructions:** It explicitly tells the LLM to leave other columns (`query_param`, validations, actions, verbose, details) empty or as placeholders, as this information isn't typically present in a simple `curl` command. The user will need to manually add expected outcomes (`expect_response_code`, `expect_response_body`, etc.) and potentially actions later.
6.  **Output Format:** It requires the output as a Markdown table with the correct column headers, containing just the single generated row.

This prompt provides a structured way for the LLM to process `curl` commands and generate a starting point for defining test cases in your Excel file. The user will still need to manually add validations and actions based on the API's expected response.
-----