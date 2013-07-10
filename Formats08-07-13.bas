Attribute VB_Name = "Formats"
'Function to select rows.  will be called later in program
Sub SelectMemberData()
    '--Code Below should select all contiguous cells "
    '--only add the headings after the sort so that it isn't selected by the script
    Range("A2").CurrentRegion.Select()
    '--Need to deselect only the first row
    '--Saves as Levels Passed by Members - Fulton Hogan: ActiveWorkbook.SaveAs Filename:="Z:\Windows Shared Folder\01.Work - Brava\Strada7\Levels Passed by Members - Fulton Hogan.xls", FileFormat:=xlAddIn8
    '--- Enable the alerts
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
Sub Add_Header()
    Sheets("Active").Select
    Call AddFirstRow
    Call CopyHeader
    Sheets("Active").Select
    Rows("1:1").Select
    ActiveSheet.Paste
    Sheets("FL Certificates").Select
    Call AddFirstRow
    Call CopyHeader
    Worksheets("FL Certificates").Select
    Rows("1:1").Select
    ActiveSheet.Paste
End Sub

'Deletes extra sheets
Sub DeleteExtraSheets()
    Application.DisplayAlerts = False
    Sheets("Admin codes and info").Delete
    Sheets("Misc accounts").Delete
    Sheets("Coach and Dist Finished").Delete
    Sheets("Sub cancelled").Delete
    Application.DisplayAlerts = True
End Sub

'Formats sheets by removing first column and setting row height to 15.  Sheet must be selected
Sub RemoveIDAndFormatRow()
    Rows.Select
    Selection.RowHeight = 15
    Columns("A:A").Select
    Selection.Delete
End Sub

'Creates filtered report, that all the other reports will be built from.
Sub CreateFilteredReport()
    Sheets("Active").Select
    Range("A1").CurrentRegion.Select
    'sort from recorded macros - need to find range and apply
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

Sub CoachesReport(ByVal CoachName As String)
    ActiveWorkbook.Sheets("Active").Copy After:=Worksheets("Active")
    Sheets("Active (2)").name = CoachName
    Sheets(CoachName).Select
    Range("A1").CurrentRegion.Select
    
    
    ' TODO
    '--Fix Fulton Hogan - try to get it using wild cards
     
    If CoachName Like "Onirik" Then '--Branch for suppliers
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(6) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(5) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ElseIf CoachName Like "Harrison Grierson" Or CoachName Like "*Fulton Hogan*" Or CoachName Like "CIGNA" Then '-- Branch for companies
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    Else '--Branch for coaches
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(5) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(CoachName).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    
    With ActiveWorkbook.Worksheets(CoachName).Sort
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    RemoveOtherCoaches (CoachName) 'Pass in the name passed in at the top
    End Sub
   
   Sub RemoveOtherCoaches(ByVal CoachName As String) 'pass coach name to search for and filter by

Dim FirstRow As Long
Dim LastRow As Long
Dim Lrow As Long

FirstRow = ActiveSheet.UsedRange.Cells(1).Row
LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "F").End(xlUp).Row

For Lrow = LastRow To FirstRow Step -1
    
    If Not IsError(ActiveSheet.Cells(Lrow, "E").Value) Then ' fix this
        Select Case CoachName
        Case "Onirik" '--Branch for Distributors
            If Not ActiveSheet.Cells(Lrow, "F").Value Like CoachName And Not ActiveSheet.Cells(Lrow, "F").Row = "1" Then
                ActiveSheet.Cells(Lrow, "F").EntireRow.Delete
            End If
        Case "*Fulton Hogan*", "Harrison Grierson", "CIGNA" '--Branch for Companies - can I use wilds for Fulton Hogan?
            If Not ActiveSheet.Cells(Lrow, "C").Value Like "*" & CoachName & "*" And Not ActiveSheet.Cells(Lrow, "C").Row = "1" Then
                ActiveSheet.Cells(Lrow, "C").EntireRow.Delete
            End If
        Case Else '--Branch for Coaches
            If Not ActiveSheet.Cells(Lrow, "E").Value Like CoachName And Not ActiveSheet.Cells(Lrow, "E").Row = "1" Then
                ActiveSheet.Cells(Lrow, "E").EntireRow.Delete
            End If
        End Select
    End If
Next Lrow
  
    
     

    
End Sub

Sub SendTabsToFile()
     '-- Save a copy for Blair.  Will need to Edit Save locations.  Need to change file type to plain xls.  No macro enabled
   ' ActiveWorkbook.SaveAs Filename:="Z:\Windows Shared Folder\01.Work - Brava\Strada7\Levels Passed by Members.xls", FileFormat:=xlAddIn8
    '--Copy the filterd worksheet into new tabs.  Will come back to the certificate tab as it will require logic
    '--Fulton Hogan's report is first
    'ActiveWorkbook.Sheets("Active").Copy
End Sub

' Levels_Passed_Formatting Macro:  Main method calls all of the other methods to create custom reports for each coach.
Sub Main()
    Sheets("FL Certificates").Select
    Call RemoveIDAndFormatRow
    Sheets("Active").Select
    Call RemoveIDAndFormatRow
    Call CreateFilteredReport
    Call CoachesReport("Brad Munns")
    Call CoachesReport("Colin Douglas")
    Call CoachesReport("Michelle Dalley")
    Call CoachesReport("CIGNA")
    Call CoachesReport("Paul")
    Call CoachesReport("Onirik")
    Call CoachesReport("Fulton Hogan")
    Call CoachesReport("Harrison Grierson")
    
    'Next Steps:
    ' 3) write processes to create each report. Includes copying each sheet
    ' 4) copy header into the top line of each report - Call Add_Header
    ' 5) export each report to file
    ' 6) delete all extra tabs from filtered report - Call DeleteExtraSheets
    ' 7) for this iteration ignore FL certificates, but plan how to deal with them.
    

    
End Sub



