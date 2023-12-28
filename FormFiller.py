from google.oauth2 import service_account
from googleapiclient.discovery import build
from FileFinder import get_first_line
import string

# Replace the following values with your actual service account credentials and scopes
SERVICE_ACCOUNT_FILE = 'FormKeys/client_secret.json'
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

issue_header_arr = []  # issue - column header
county_header_arr = []  # county - row header
descr_issue_arr = []  # data

item_header_arr = []  # issue - column header
resource_data_arr = []  # county - row header
data_arr = []  # data

def find_intersection_and_update_arrays(credentials, spreadsheet_id, sheets):
    # Create a Google Sheets API service using the provided credentials
    sheets_service = build('sheets', 'v4', credentials=credentials)

    for sheet_name in sheets:
        # Read existing data from the specified range (C2:E)
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f'{sheet_name}!C2:H'
        ).execute()

        values = result.get('values', [])

        # Update global arrays with the retrieved data
        for row in values:
            if len(row) >= 5:  # Ensure at least five columns are present (C, D, E, F, G)
                issue_header_arr.append(row[0])
                county_header_arr.append(row[1])
                descr_issue_arr.append(row[2])
                
                item_header_arr.append(row[3])
                resource_data_arr.append(row[4])
                data_arr.append(row[5])


def find_intersection(credentials, spreadsheet_id, sheet_name, row_header, column_header):
    # Create a Google Sheets API service using the provided credentials
    sheets_service = build('sheets', 'v4', credentials=credentials)

    # Read existing data from the sheet
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A1:ZZZ'
    ).execute()

    values = result.get('values', [])

    # Find the position based on the intersection of row_header and column_header
    target_row = None
    target_column = None

    for i, row in enumerate(values):
        if len(row) > 0 and row[0] == row_header:
            target_row = i + 1  # Adjust for 1-based indexing
            break

    for j, col in enumerate(values[0]):
        if col == column_header:
            target_column = j + 1  # Adjust for 1-based indexing
            break

    # If both row and column headers are found, return the intersection
    if target_row is not None and target_column is not None:
        return sheet_name, target_row, target_column

    # If not found in the sheet, return None
    return None

def edit_google_sheets(credentials, spreadsheet_id, sheets):
    # Example: Add values to the Google Sheets spreadsheet
    # Get user input for the operation
    
        # Editing existing data
    # Use data gathered from find_intersection_and_update_arrays
    for i in range(len(issue_header_arr)):
        column_header = county_header_arr[i]
        row_header = issue_header_arr[i]  
        new_data = descr_issue_arr[i]

        # Find the intersection in any of the sheets
        for sheet_name in sheets:
            intersection_info = find_intersection(credentials, spreadsheet_id, sheet_name, row_header, column_header)

            if intersection_info:
                sheet_name, target_row, target_column = intersection_info

                # Create a Google Sheets API service using the provided credentials
                sheets_service = build('sheets', 'v4', credentials=credentials)

                range_ = f'{sheet_name}!{string.ascii_uppercase[target_column - 1]}{target_row}'  # Corrected column index to letter

                # Retrieve existing data from the cell
                existing_data = sheets_service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_
                ).execute().get('values', [])

                # Check if the new data is not already present in the existing content
                if existing_data and existing_data[0]:
                    existing_content = existing_data[0][0]
                    if new_data not in existing_content:
                        # Concatenate new data to existing data with two new line characters before the new data
                        new_data = f"{existing_content}\n\n{new_data}"
                    else:
                        print(f"Data '{new_data}' is already present in the cell. Skipping update.")
                        continue  # Skip the update if the data is already present

                body = {'values': [[new_data]]}

                try:
                    result = sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=range_,
                        body=body,
                        valueInputOption='RAW'
                    ).execute()

                    print(f"Appended data to the cell. Response: {result}")
                except Exception as e:
                    print(f"Error updating Google Sheets: {e}")
                break  # Exit loop after updating the data
        else:
            print(f"Headers not found in any sheet.")

def addResources(credentials, spreadsheet_id, sheets):
    for i in range(len(item_header_arr)):
        column_header = item_header_arr[i]
        row_header = resource_data_arr[i]
        new_data = data_arr[i]

        # Find the intersection in any of the sheets
        for sheet_name in sheets:
            intersection_info = find_intersection(credentials, spreadsheet_id, sheet_name, row_header, column_header)

            if intersection_info:
                sheet_name, target_row, target_column = intersection_info

                # Create a Google Sheets API service using the provided credentials
                sheets_service = build('sheets', 'v4', credentials=credentials)

                range_ = f'{sheet_name}!{string.ascii_uppercase[target_column - 1]}{target_row}'  # Corrected column index to letter

                # Retrieve existing data from the cell
                existing_data = sheets_service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_
                ).execute().get('values', [])

                # Check if the new data is not already present in the existing content
                if existing_data and existing_data[0]:
                    existing_content = existing_data[0][0]
                    if new_data not in existing_content:
                        # Concatenate new data to existing data with two new line characters before the new data
                        new_data = f"{existing_content}\n\n{new_data}"
                    else:
                        print(f"Data '{new_data}' is already present in the cell. Skipping update.")
                        continue  # Skip the update if the data is already present

                body = {'values': [[new_data]]}

                try:
                    result = sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=range_,
                        body=body,
                        valueInputOption='RAW'
                    ).execute()

                    print(f"Appended data to the cell. Response: {result}")
                except Exception as e:
                    print(f"Error updating Google Sheets: {e}")
                break  # Exit loop after updating the data
        else:
            print(f"Headers not found in any sheet.")




def main():

    try:
        credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    except:
        print("Error in Credentials")
        return None
    
    try:
        form_id = get_first_line('FormKeys/FormKey.txt')  
        spreadsheet_id = get_first_line('FormKeys/SheetToFill.txt')  
        
        sheets = ['County/School District, Organization/Issues', 'Statewide, Organization/Issues', 'County/School District, Current Issues']  # Add your sheet names here

        find_intersection_and_update_arrays(credentials, form_id, ['Form Responses 1'])
        edit_google_sheets(credentials, spreadsheet_id, sheets)
        addResources(credentials, spreadsheet_id, ['Resources'])

    except: 
        print("Error in Credentials or Sheet ID")
        return None

if __name__ == '__main__':
    main()
