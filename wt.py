import os
import win32com.client
import openpyxl

def read_script_preferences(excel_file):
    """Reads script paths and their execution preferences from an Excel file."""
    script_preferences = []
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        script_path = row[0]
        execute = row[1].lower()  # Convert 'Yes'/'No' to lowercase
        if execute in ['yes', 'no']:
            script_preferences.append((script_path, execute == 'yes'))
        else:
            print(f"Invalid preference '{execute}' for script '{script_path}'. Skipping.")
    wb.close()
    return script_preferences

def run_uft_script(script_path):
    try:
        # Create the UFT application object
        uft = win32com.client.Dispatch("QuickTest.Application")
        
        # Open UFT if not already open
        if not uft.Launched:
            uft.Launch()
            
        # Make UFT visible (optional)
        uft.Visible = True

        # Open the test
        uft.Open(script_path)

        # Get the test object
        test = uft.Test

        # Run the test
        test.Run()
        
        # Wait for the test to complete (optional)
        while test.Status == "Running":
            continue
        
        # Get the result (example)
        result = test.LastRunResults.Status
        print(f'Test Result: {result}')
    
    except Exception as e:
        print(f'Error occurred: {str(e)}')
    
    finally:
        # Close the test
        test.Close()

        # Quit UFT
        uft.Quit()

def execute_selected_scripts(excel_file):
    """Execute UFT scripts based on preferences in the Excel file."""
    script_preferences = read_script_preferences(excel_file)
    for script_path, execute in script_preferences:
        if execute:
            print(f"Executing script: {script_path}")
            run_uft_script(script_path)
        else:
            print(f"Skipping script: {script_path}")

# Example usage
excel_file = "script_preferences.xlsx"  # Replace with your Excel file path
execute_selected_scripts(excel_file)
