import xlwings as xw
import openpyxl
from xlwings.utils import rgb_to_int
import win32com.client
from openpyxl.drawing.image import Image
from PIL import Image as PILImage, ImageDraw, ImageFont
import os


# Define the comparison files and corresponding sheets
def main(file_previous, file_latest): 
    print(f"{file_previous} + {file_latest}")
    comparison_files = [
        ('ComparedResults/Full_Comparison.xlsx', 'Dashboard'),
        ('ComparedResults/CCCTA_Comparison.xlsx', 'CCCTA'),
        ('ComparedResults/LAVTA_Comparison.xlsx', 'LAVTA')
    ]

    # Open the Dashboard workbook
    dashboard_file = 'ComparedResults/Dashboard.xlsm'
    app = xw.App(visible=False)

    try:
        # Open the existing workbook (Dashboard)
        wb_dashboard = app.books.open(dashboard_file)

        # Loop through each comparison file and corresponding sheet
        for comparison_file, sheet_name in comparison_files:

            # Load the comparison workbook and sheet
            wb_comparison = openpyxl.load_workbook(comparison_file)
            sheet_summary = wb_comparison['Summary']

            # paste_picture(comparison_files, dashboard_file)

            # Retrieve Payment values from the 'Summary' sheet in the comparison file
            full_prev_amount = f"{float(sheet_summary['B6'].value):,.2f}"
            full_lat_amount = f"{float(sheet_summary['C6'].value):,.2f}"

            # Retrieve values from the 'Summary' sheet in the comparison file
            full_amount_diff = round(abs(sheet_summary['D6'].value), 2)
            full_amount_diff_status = sheet_summary['E6'].value

            prev_full_trip_made = f"{(sheet_summary['B2'].value):,}"
            lat_full_trip_made = f"{(sheet_summary['C2'].value):,}"
            diff_full_trip_made = f"{float(sheet_summary['D2'].value):.0f}"

            prev_full_hrs_op = f"{(sheet_summary['B3'].value):,.0f}"
            lat_full_hrs_op = f"{(sheet_summary['C3'].value):,.0f}"
            diff_full_hrs_op = f"{float(sheet_summary['D3'].value):.0f}"

            prev_full_op = f"{(sheet_summary['B4'].value):,}"
            lat_full_op = f"{(sheet_summary['C4'].value):,}"
            diff_full_op = f"{float(sheet_summary['D4'].value):.0f}"

            prev_full_days = f"{(sheet_summary['B5'].value):,}"
            lat_full_days = f"{(sheet_summary['C5'].value):,}"
            diff_full_days = f"{float(sheet_summary['D5'].value):.0f}"

            # Get the corresponding sheet in the dashboard
            sheet_dashboard = wb_dashboard.sheets[sheet_name]

            # Access the file name values shape via the API and set the value
            txt_prev_file = sheet_dashboard.shapes['txtPrevFile'].api
            txt_prev_file.TextFrame2.TextRange.Text = f"{file_previous}"

            txt_lat_file = sheet_dashboard.shapes['txtLatFile'].api
            txt_lat_file.TextFrame2.TextRange.Text = f"{file_latest}"

            # Access the total diff value shape via the API and set the value
            txt_full_amount_diff = sheet_dashboard.shapes['TextBox 88'].api
            txt_full_amount_diff.TextFrame2.TextRange.Text = f"$ {full_amount_diff}"

            txt_full_prev_amount = sheet_dashboard.shapes['txtPrevPAyment'].api
            txt_full_prev_amount.TextFrame2.TextRange.Text = f"$ {full_prev_amount}"

            txt_full_lat_amount = sheet_dashboard.shapes['txtLatPayment'].api
            txt_full_lat_amount.TextFrame2.TextRange.Text = f"$ {full_lat_amount}"

            # Access the total diff status shape via the API and set the value
            txt_full_amount_diff_status = sheet_dashboard.shapes['TextBox 90'].api
            txt_full_amount_diff_status.TextFrame2.TextRange.Text = f"{full_amount_diff_status}"

            # Access the trips made shape via the API and set the value
            txt_full_trip_made = sheet_dashboard.shapes['txtDTripsMAde'].api
            txt_full_trip_made.TextFrame2.TextRange.Text = f"{prev_full_trip_made} trips to {lat_full_trip_made} trips"
            txt_full_trip_diff = sheet_dashboard.shapes['txtTripsDiff'].api
            txt_full_trip_diff.TextFrame2.TextRange.Text = f"{diff_full_trip_made} trips"

            # Access the hours operated shape via the API and set the value
            txt_full_hrs_op = sheet_dashboard.shapes['txtDHoursOp'].api
            txt_full_hrs_op.TextFrame2.TextRange.Text = f"{prev_full_hrs_op} hours to {lat_full_hrs_op} hours"
            txt_full_hrs_diff = sheet_dashboard.shapes['txtHoursDiff'].api
            txt_full_hrs_diff.TextFrame2.TextRange.Text = f"{diff_full_hrs_op} hours"

            # Access the operators shape via the API and set the value
            txt_full_op = sheet_dashboard.shapes['txtDOperators'].api
            txt_full_op.TextFrame2.TextRange.Text = f"{prev_full_op} to {lat_full_op} operators"
            txt_full_op_diff = sheet_dashboard.shapes['txtOpsDiff'].api
            txt_full_op_diff.TextFrame2.TextRange.Text = f"{diff_full_op} operators"

            # Access the days operated shape via the API and set the value
            txt_full_days = sheet_dashboard.shapes['txtDDays'].api
            txt_full_days.TextFrame2.TextRange.Text = f"{prev_full_days} hours to {lat_full_days} hours"
            txt_full_days_diff = sheet_dashboard.shapes['txtDaysDiff'].api
            txt_full_days_diff.TextFrame2.TextRange.Text = f"{diff_full_days} days"

            # Run the VBA macro to update the color based on the status
            try:
                # Parameters: TextBox name and status
                textBoxName = "TextBox 90"
                status = full_amount_diff_status  # Use the status from the comparison file

                # Call the VBA macro to update color
                wb_dashboard.macro("UpdateTextBoxColor")(sheet_name, textBoxName, status)
                print(f"Successfully updated color for {textBoxName} with status '{status}'.")
            except Exception as e:
                print(f"An error occurred: {e}")

            # Run the VBA macro to update the color based on the values
            try:
                # Parameters: TextBox names and corresponding values
                textBoxNames = ["txtTripsDiff", "txtHoursDiff", "txtOpsDiff", "txtDaysDiff"]
                values = [diff_full_trip_made, diff_full_hrs_op, diff_full_op, diff_full_days]

                # Loop through the text boxes and update colors based on the values
                for i, textBoxName in enumerate(textBoxNames):
                    wb_dashboard.macro("UpdateSummaryColor")(sheet_name, textBoxName, values[i])
                    print(f"Successfully updated color for {textBoxName} with value '{values[i]}'.")
            except Exception as e:
                print(f"An error occurred: {e}")

        # Save the changes to the dashboard workbook
        wb_dashboard.save()
        wb_dashboard.close()
        print(f"{dashboard_file} has been successfully updated and saved.")

        # paste_picture(comparison_files, dashboard_file)
        paste_picture()

    except Exception as e:
        print(f"An error occurred: {e}")
    # finally:
    #     wb_dashboard.save()
    #     wb_dashboard.close()
        # app.quit()
        
        # Reopen the Excel file
        # app = xw.App(visible=True)  # Open Excel with the app visible
        # wb_dashboard = app.books.open(dashboard_file)  # Reopen the file

def paste_picture():
    comparison_files = [
        ('ComparedResults/Full_Comparison.xlsx', 'Dashboard'),
        ('ComparedResults/CCCTA_Comparison.xlsx', 'CCCTA'),
        ('ComparedResults/LAVTA_Comparison.xlsx', 'LAVTA')
    ]
    
    # Target cells for each sheet in the comparison file
    target_cells = {
        'TripsComparison': (11, 21),
        'HoursComparison': (44, 4),
        'OperatorChanges': (44, 16)
    }

    relative_dashboard_path = "ComparedResults\\Dashboard.xlsm"
    # Get the absolute path of the current script's directory
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # Build the full path to the dashboard file by joining the script directory and the relative path
    dashboard_file = os.path.join(script_dir, relative_dashboard_path)

    try:
        # Initialize Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Set to True for debugging

        # Check if the file exists
        if not os.path.exists(dashboard_file):
            print(f"Error: Dashboard file does not exist at {dashboard_file}")
            return
        
        # Open the Dashboard workbook
        wb_dashboard = excel.Workbooks.Open(dashboard_file)
        if wb_dashboard is None:
            print(f"Failed to open the Dashboard workbook at {dashboard_file}")
            return

        # Delete existing pictures if they exist
        for target_sheet_name in ['Dashboard', 'CCCTA', 'LAVTA']:
            ws_dashboard = wb_dashboard.Sheets(target_sheet_name)
            for picture_name in ['TripsTable', 'HoursTable', 'OperatorTable']:
                try:
                    ws_dashboard.Shapes(picture_name).Delete()  # Attempt to delete the picture
                    print(f"Deleted existing picture: {picture_name} in {target_sheet_name}")
                except Exception:
                    print(f"No existing picture named {picture_name} found in {target_sheet_name}")

        # Process each comparison file
        for comparison_file, target_sheet_name in comparison_files:
            # Build the full path for the comparison file
            comparison_file_path = os.path.join(script_dir, comparison_file)
            
            # Check if the comparison file exists
            if not os.path.exists(comparison_file_path):
                print(f"Error: Comparison file does not exist at {comparison_file_path}")
                continue

            # Open the comparison workbook
            wb_comparison = excel.Workbooks.Open(comparison_file_path)
            if wb_comparison is None:
                print(f"Failed to open the comparison workbook at {comparison_file_path}")
                continue

            # Process each sheet in the comparison file (TripsComparison, HoursComparison, OperatorChanges)
            for sheet_name, target_cell in target_cells.items():
                sheet = wb_comparison.Sheets(sheet_name)
                if sheet is None:
                    print(f"Failed to access the '{sheet_name}' sheet in {comparison_file_path}")
                    continue

                # Get the used range
                used_range = sheet.UsedRange

                # Export the range to a temporary clipboard as a picture
                used_range.CopyPicture(Format=2)  # Format=2 -> Bitmap format (default)

                # Activate the target sheet in the Dashboard workbook
                ws_dashboard = wb_dashboard.Sheets(target_sheet_name)
                if ws_dashboard is None:
                    print(f"Failed to access the sheet '{target_sheet_name}' in the Dashboard workbook.")
                    wb_comparison.Close(SaveChanges=False)
                    continue

                # Paste as a picture into the target sheet
                ws_dashboard.Activate()
                row, col = target_cell
                target_cell_range = ws_dashboard.Cells(row, col)  # Adjust as needed
                ws_dashboard.Paste(target_cell_range)

                # Position and resize the pasted picture
                pasted_picture = ws_dashboard.Shapes(ws_dashboard.Shapes.Count)
                pasted_picture.Left = target_cell_range.Left
                pasted_picture.Top = target_cell_range.Top

                # Name the pasted picture according to the sheet
                if sheet_name == 'TripsComparison':
                    pasted_picture.Name = 'TripsTable'
                elif sheet_name == 'HoursComparison':
                    pasted_picture.Name = 'HoursTable'
                elif sheet_name == 'OperatorChanges':
                    pasted_picture.Name = 'OperatorTable'

            # Close the comparison workbook without saving
            wb_comparison.Close(SaveChanges=True)

        # Save and close the Dashboard workbook
        wb_dashboard.Save()
        wb_dashboard.Close()
        excel.Quit()

        print("Data pasted as pictures successfully.")

        # app = xw.App(visible=True)  # Open Excel with the app visible
        # wb_dashboard = app.books.open(dashboard_file)  # Reopen the file
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure Excel is properly quit and the object is released
        if excel:
            excel.Quit()  # Close the Excel application
            del excel 

# if __name__ == '__main__':
    # main()
    # comparison_files = [
    #     ('Compared Results/Full_Comparison.xlsx', 'Dashboard'),
    #     ('Compared Results/CCCTA_Comparison.xlsx', 'CCCTA'),
    #     ('Compared Results/LAVTA_Comparison.xlsx', 'LAVTA')
    # ]
    # dashboard_file = 'Compared Results/Dashboard.xlsm'
    # paste_picture(comparison_files, dashboard_file)
    # paste_picture(comparison_files, dashboard_file)