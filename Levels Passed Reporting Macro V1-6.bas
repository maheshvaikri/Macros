Attribute VB_Name = "Formats"

Enum Relationship
    COACH
    COMPANY
    DISTRIBUTOR
End Enum

'using this to store the relationship passed into the CoachesReport() subroutine
Dim relationshipRelationship As Relationship



'Function Saves a copy of the workbook for sorting, leaving the origional un-molested
Sub SaveWorkingCopy()
 
    Dim relativePath As String
    Application.DisplayAlerts = False
    
    relativePath = ThisWorkbook.Path & "\" & "Levels Passed by Members " & Day(Now()) & "-" & Month(Now()) & "-" & Year(Now()) & " Filtered.xlsx"

    ActiveWorkbook.SaveAs Filename:=relativePath

    Application.DisplayAlerts = True

End Sub


'function to add a row to the top of a workbook. Correct sheet must be active
Sub AddFirstRow()
    Range("A1").Select
    ActiveCell.EntireRow.Insert
End Sub

'Copies header from "Admin codes and info" sheet
Sub CopyHeader()
    Sheets("Admin codes and info").Select
    Rows("9:9").Select
    Selection.Copy
End Sub

'adds header to Active and FL Certificate sheets. Header cannot be added before sorting or selection
Sub Add_Header(ByVal Name As String)
    Sheets(Name).Select
    Call AddFirstRow
    Call CopyHeader
    Sheets(Name).Select
    Rows("1:1").Select
    ActiveSheet.Paste

End Sub

'Deletes extra sheets
Sub DeleteExtraSheets()
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Admin codes and info").Delete
    Sheets("Misc accounts").Delete
    Sheets("Coach and Dist Completed").Delete
    Sheets("Sub cancelled").Delete
    Application.DisplayAlerts = True
    Application.Goto (ActiveWorkbook.Sheets("Active").Range("A1")) ' just to finish on with the first tab in view
End Sub

Sub DeleteCurrentSheet(ByVal Name As String)
Application.DisplayAlerts = False
    On Error Resume Next
    Sheets(Name).Delete
    Application.DisplayAlerts = True
    
End Sub


'Formats sheets by removing first column and setting row height and column widths.
'A worksheet must be selected before it is called.
Sub RemoveIDAndFormatRow()
    Rows.Select
    Selection.RowHeight = 15
    Columns("A:A").Delete
    'set width of certain columns
    Columns("A").ColumnWidth = 25
    Columns("B").ColumnWidth = 15
    Columns("C").ColumnWidth = 24.71
    Columns("E").ColumnWidth = 15.29
    Columns("I:R").ColumnWidth = 11
    
    
End Sub

'Creates filtered report that all the other reports will be built from.
Sub CreateFilteredReport()
    Sheets("Active").Select
    Range("A1").CurrentRegion.Select
    ActiveWorkbook.Worksheets("Active").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Active").Sort.SortFields.Add Key:=Selection.Columns(6), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Active").Sort.SortFields.Add Key:=Selection.Columns(5) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Active").Sort.SortFields.Add Key:=Selection.Columns(3) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Active").Sort.SortFields.Add Key:=Selection.Columns(1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ActiveWorkbook.Worksheets("Active").Sort
        .SetRange Range("A1").CurrentRegion ' should select all contiguous cells
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

' Subroutine does the heavy lifting, sorting appropriate fields and deleting unessecary rows from the report by calling the
' RemoveOtherCoaches subroutine.  The file is then cleaned up by calling the AddHeader, SendTabsToFile and DeleteCurrentSheet.

Sub CoachesReport(ByVal Name As String, ByVal Role As Relationship)
    relationshipRelationship = Role 'Relationship parameter value
    'is retained here until the CoachesReport method is called again
    
    ActiveWorkbook.Sheets("Active").Copy After:=Worksheets("Active")
    Sheets("Active (2)").Name = Name
    Sheets(Name).Select
    Range("A1").CurrentRegion.Select
     
    If Role = DISTRIBUTOR Then
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(6) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(5) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ElseIf Role = COMPANY Then
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ElseIf Role = COACH Then
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(5) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    
   With ActiveWorkbook.Worksheets(Name).Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    RemoveOtherCoaches (Name)
    Add_Header (Name) ' Adds a header to the page using the same name parameter
    SendTabsToFile (Name)
    DeleteCurrentSheet (Name)
 
End Sub

   
   ' Takes a coach/company/distributor name as a parameter to search for and filter by.
   ' Rows for all other entities are deleted from the sheet.
    Sub RemoveOtherCoaches(ByVal Name As String)

    Dim FirstRow As Long
    Dim LastRow As Long
    Dim Lrow As Long

    FirstRow = ActiveSheet.UsedRange.Cells(1).Row
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "F").End(xlUp).Row

For Lrow = LastRow To FirstRow Step -1
    
    If Not IsError(ActiveSheet.Cells(Lrow, "E").Value) Then ' fix this
        
        Select Case relationshipRelationship
        Case DISTRIBUTOR '--Branch for Distributors
            If Not ActiveSheet.Cells(Lrow, "F").Value Like Name And Not ActiveSheet.Cells(Lrow, "F").Row = "1" Then
                ActiveSheet.Cells(Lrow, "F").EntireRow.Delete
            End If
            
        Case COMPANY
            If Name Like "Fulton Hogan" Then '--Branch for Fulton Hogan (NZ & Aus).  Can't use wilds so making seperate branch
                If Not (ActiveSheet.Cells(Lrow, "C").Value Like "Fulton Hogan Au" Or ActiveSheet.Cells(Lrow, "C").Value Like "Fulton Hogan NZ") And Not ActiveSheet.Cells(Lrow, "C").Row = "1" Then
                    ActiveSheet.Cells(Lrow, "C").EntireRow.Delete
                End If
            Else
                If Not ActiveSheet.Cells(Lrow, "C").Value Like Name And Not ActiveSheet.Cells(Lrow, "C").Row = "1" Then
                    ActiveSheet.Cells(Lrow, "C").EntireRow.Delete
                End If
            End If
            
        Case COACH '--Branch for Coaches
            If Not ActiveSheet.Cells(Lrow, "E").Value Like Name And Not ActiveSheet.Cells(Lrow, "E").Row = "1" Then
                ActiveSheet.Cells(Lrow, "E").EntireRow.Delete
            End If
        End Select
    End If
Next Lrow
    
End Sub

Sub SendTabsToFile(ByVal Name As String)
    Dim relativePath As String
    Application.DisplayAlerts = False
    
    relativePath = ThisWorkbook.Path & "\" & "Levels Passed by Members " & Day(Now()) & "-" & Month(Now()) & "-" & Year(Now()) & " " & Name & ".xlsx"
'    ActiveWorkbook.SaveAs Filename:=relativePath

   ActiveSheet.Copy
    With ActiveWorkbook
        .SaveAs Filename:=relativePath
        .Close 0
    End With



    Application.DisplayAlerts = True
  
End Sub

' "MAIN" is the macro that you need to select from Excels Developer -> Macros menu to run the script. The MAIN Subroutine
'  calls all of the required Subroutines to create custom reports for each coach/distributor/company.
Sub MAIN()
    Application.ScreenUpdating = False
    Call SaveWorkingCopy
    Sheets("FL Certificates").Select
    Call RemoveIDAndFormatRow
    Sheets("Active").Select
    Call RemoveIDAndFormatRow
    Call CreateFilteredReport
    
    ' the CoachesReport("name", RELATIONSHIP) method is used to create a new report.  It calls all of the other methods
    ' neccessary to filter, format and send the report to file.
    ' The "name" parameter must match exactly the coach/distributor/company name used in the spreadsheet.
    ' Values for the relationship parameter are COACH, DISTRIBUTOR, COMPANY.
    Call CoachesReport("Brad Munns", COACH)
    Call CoachesReport("Colin Douglas", COACH)
    Call CoachesReport("Michelle Dalley", COACH)
    Call CoachesReport("CIGNA", COMPANY)
    Call CoachesReport("Paul", COACH)
    Call CoachesReport("Onirik", DISTRIBUTOR)
    Call CoachesReport("Fulton Hogan", COMPANY)
    Call CoachesReport("Harrison Grierson", COMPANY)
    
    ' This Subroutine removes extra tabs from the workbooks such as "admin codes and info" and "sub cancelled", leaving only
    ' the active sheet.
    Call DeleteExtraSheets
    
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
    'Next Steps:
    ' for this iteration ignore FL certificates, but plan how to deal with them.
        
End Sub



