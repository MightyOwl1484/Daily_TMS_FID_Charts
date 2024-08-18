CreateStackedColumnCharts() VBA Script Overview
Introduction
The CreateStackedColumnCharts() VBA script is designed to automate the creation of stacked column charts in Excel. This script loops through data on a worksheet named "Daily," identifies unique values for TMS, and creates separate worksheets and charts for each TMS. The resulting charts visually represent the count of BUNO over time, with distinct colors assigned to different stakeholders. The script ensures that charts are dynamically generated with appropriate titles and formatted axes.

Script Breakdown
1. Initialization
Screen Updating:

The script begins by disabling screen updating (Application.ScreenUpdating = False) to improve performance and prevent flickering during the chart creation process.
Dictionary Objects:

Three dictionaries (tmsDict, serviceDict, tmsDatesDict) are initialized using CreateObject("Scripting.Dictionary"). These dictionaries store unique values for TMS, Services, and dates associated with each TMS.
Service Colors:

A serviceColors dictionary is also initialized to map specific Services to their corresponding colors using RGB values.
2. Data Collection
Loop Through "Daily" Worksheet:
The script iterates over each row of the "Daily" worksheet to extract unique values for TMS, Service, and dates.
If a TMS or Service value does not already exist in the dictionary, it is added.
Dates are stored in a nested dictionary within tmsDatesDict for each TMS.
3. Chart Creation
Create New Worksheets:

For each unique TMS, a new worksheet is created. The worksheet is named after the TMS value, and the chart title is set to "TMS induction date projection".
X-Axis Values:

The unique dates for each TMS are extracted and sorted using a custom SortArray() function to ensure they are in chronological order.
Add Chart:

A new stacked column chart is added to each TMS worksheet.
The chart is populated with series corresponding to each Service, using the count of BUNO values for the Y-axis and the sorted dates for the X-axis.
4. Data Labeling and Axis Formatting
Data Labels:

Data labels are added to each series, but labels for zero values are suppressed to maintain chart clarity.
Axis Titles:

The Y-axis is labeled "Count of BUNO", and the X-axis labels are formatted to display the month and year ("mmm yyyy").
5. Saving the Workbook
Save As:
After all charts have been created, the workbook is saved with a specified filename ("TMS_FID_Projection_AIRRS.xlsx") in the same directory as the original workbook.
6. Re-enable Screen Updating
Final Step:
The script concludes by re-enabling screen updating (Application.ScreenUpdating = True).
Supporting Function: SortArray()
Purpose:

The SortArray() function is a custom sorting algorithm used to sort the array of dates before they are applied to the X-axis of the charts.
Implementation:

The function employs a basic bubble sort algorithm to reorder the dates in ascending order.
