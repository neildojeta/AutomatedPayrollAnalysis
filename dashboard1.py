import xlwings as xw
import openpyxl
from xlwings.utils import rgb_to_int

# Open the Full_Comparison.xlsx file and retrieve value from cell D6 in the Summary sheet
full_comparison_file = 'Compared Results/Full_Comparison.xlsx'

# Load the workbook and sheet
wb_full_comparison = openpyxl.load_workbook(full_comparison_file)
sheet_summary = wb_full_comparison['Summary']

# Get the value from cell D6
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

# Open the Dashboard.xlsm file
dashboard_file = 'Compared Results/Dashboard.xlsm'

# Create the Excel application instance (visible=True to see the operation)
app = xw.App(visible=True)

try:
    # Open the existing workbook
    wb_dashboard = app.books.open(dashboard_file)

    # Get the "Dashboard" sheet
    sheet_dashboard = wb_dashboard.sheets['Dashboard']

    # Access the total diff value shape via the API and set the value
    txt_full_amount_diff = sheet_dashboard.shapes['TextBox 88'].api
    txt_full_amount_diff.TextFrame2.TextRange.Text = f"$ {full_amount_diff}"

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

    # Access the hours operated shape via the API and set the value
    txt_full_op = sheet_dashboard.shapes['txtDOperators'].api
    txt_full_op.TextFrame2.TextRange.Text = f"{prev_full_op} to {lat_full_op} operators"
    txt_full_op_diff = sheet_dashboard.shapes['txtOpsDiff'].api
    txt_full_op_diff.TextFrame2.TextRange.Text = f"{diff_full_op} operators"

    # Access the days operated shape via the API and set the value
    txt_full_days = sheet_dashboard.shapes['txtDDays'].api
    txt_full_days.TextFrame2.TextRange.Text = f"{prev_full_days} hours to {lat_full_days} hours"
    txt_full_days_diff = sheet_dashboard.shapes['txtDaysDiff'].api
    txt_full_days_diff.TextFrame2.TextRange.Text = f"{diff_full_days} days"

    # Run the VBA macro
    try:
        # Parameters: TextBox name and status
        textBoxName = "TextBox 90"
        status = full_amount_diff_status  # Change this as needed

        # Call the VBA macro
        wb_dashboard.macro("UpdateTextBoxColor")(textBoxName, status)
        print(f"Successfully updated color for {textBoxName} with status '{status}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

    # Run the VBA macro to update the color based on the value
    try:
        # Parameters: TextBox names and corresponding values
        textBoxNames = ["txtTripsDiff", "txtHoursDiff", "txtOpsDiff", "txtDaysDiff"]
        values = [diff_full_trip_made, diff_full_hrs_op, diff_full_op, diff_full_days]  # Adjust the logic for each text box
        
        # Loop through the text boxes and update colors based on the values
        for i, textBoxName in enumerate(textBoxNames):  # Use enumerate for proper iteration
            wb_dashboard.macro("UpdateSummaryColor")(textBoxName, values[i])
            print(f"Successfully updated color for {textBoxName} with value '{values[i]}'.")
            
    except Exception as e:
        print(f"An error occurred: {e}")

    # Save the changes to the workbook
    wb_dashboard.save()
    print(f"{dashboard_file} has been successfully updated and saved.")

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    wb_dashboard.save()
    wb_dashboard.close()
    app.quit()

    # Reopen the Excel file
    app = xw.App(visible=True)  # Open Excel with the app visible
    wb_dashboard = app.books.open(dashboard_file)  # Reopen the file
