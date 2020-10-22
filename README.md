# Merging-Spreadsheets


This VBA Code:

1 .Merges data from Unit Submissions into one Worksheet
2. Deletes any rows that contain the word "Sample" in Column A or have Blanks in Column A

Since this workbook contains the VBA code, keep it open until you have finished performing the merging 

Instructions:

1. Save all Unit Submissions spreadsheets in a single folder.
2. Open (or create) the workbook where you want the unit submissions to be copied into.

3. If you do not already have "Developer" tab in the top ribbon, you can enable it in File>Options>Customize Ribbon and checking "Developer" box on right side

4. Before you run the VBA, you will need to open up the code to edit the 1. location of the unit submissions spreadsheets, 2. cell where the unit submission data begins and ends and 3. name of the workbook where the submissions will be copied into.

 To view the code, hit Developer>"Visual Basic"

On the Project Explorer (top left side), drill down VBA Project (VBA Instructions for Merging Unit Submissions) > Modules > double click "Module 1". This will open up the VBA code.


Within the VBA screen, change the highlighted code below to match your data.

Yellow
Blue
Purple
Green





5. Once changed, now you can run the macro. Developer tab > Macros >  MergeUnitSubmissions
6. If you want to get rid of any blank or  "Sample" rows that may have been copied, you can run the second Macro titled "DeleteSampleAndBlanks"
