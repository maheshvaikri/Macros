Attribute VB_Name = "Formats"
Sub Levels_Passed_Formatting()
Attribute Levels_Passed_Formatting.VB_Description = "Inserts headings into the first row of the ""Active"" and ""FL Certificates"" sheets, deletes all extra tabs, removes the ID columns from the ""Active"" and ""FL Certificates"" tabs and changes the row height to 15"
Attribute Levels_Passed_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Levels_Passed_Formatting Macro
' Inserts headings into the first row of the "Active" and "FL Certificates" sheets, deletes all extra tabs, removes the ID columns from the "Active" and "FL Certificates" tabs and changes the row height to 15
'

'
    Sheets("Active").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Admin codes and info").Select
    Rows("9:9").Select
    Selection.Copy
    Sheets("Active").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("FL Certificates").Select
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    Sheets("Admin codes and info").Select
    Rows("9:9").Select
    Selection.Copy
    Sheets("FL Certificates").Select
    ActiveSheet.Paste
     Cells.Select
    Selection.RowHeight = 15
    Application.CutCopyMode = False
    '--- Disables Alerts
    Application.DisplayAlerts = False
    Sheets("Admin codes and info").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Misc accounts").Delete
    Sheets("Coach and Dist Finished").Delete
    Sheets("Sub cancelled").Delete
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("Active").Select
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Selection.RowHeight = 15
    Range("A1").Select
    '-- Save a copy for Blair.  Will need to Edit Save locations.  Need to change file type to plain xls.  No macro enabled
    ActiveWorkbook.SaveAs Filename:="Z:\Windows Shared Folder\01.Work - Brava\Strada7\Levels Passed by Members.xls", FileFormat:=xlAddIn8
    '--Copy the filterd worksheet into new tabs.  Will come back to the certificate tab as it will require logic
    '--Fulton Hogan's report is first
    ActiveWorkbook.Sheets("Active").Copy
    '--Fulton Hogan's book is now the active Workbook
    '--Code Below should select all contiguous cells "
    '--only add the headings after the sort so that it isn't selected by the script
    Range("A2").CurrentRegion.Select
    '--Need to deselect only the first row
    '--Saves as Levels Passed by Members - Fulton Hogan: ActiveWorkbook.SaveAs Filename:="Z:\Windows Shared Folder\01.Work - Brava\Strada7\Levels Passed by Members - Fulton Hogan.xls", FileFormat:=xlAddIn8
    '--- Enable the alerts
    Application.DisplayAlerts = True
End Sub
