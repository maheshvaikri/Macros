Attribute VB_Name = "Formats"
'Function Saves a copy for sorting, leaving the origional un-molested
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
'    Sheets("FL Certificates").Select
'    Call AddFirstRow
'    Call CopyHeader
'    Worksheets("FL Certificates").Select
'    Rows("1:1").Select
'    ActiveSheet.Paste
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

Sub CoachesReport(ByVal Name As String)
  
    ActiveWorkbook.Sheets("Active").Copy After:=Worksheets("Active")
    Sheets("Active (2)").Name = Name
    Sheets(Name).Select
    Range("A1").CurrentRegion.Select
     
     
    If Name Like "Onirik" Then '--Branch for suppliers
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(6) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(5) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ElseIf Name Like "Harrison Grierson" Or Name Like "Fulton Hogan" Or Name Like "CIGNA" Then '-- Branch for companies
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(Name).Sort.SortFields.Add Key:=Selection.Columns(1) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    Else '--Branch for coaches
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
    
    RemoveOtherCoaches (Name) 'Pass in the name passed in at the top
    Add_Header (Name) ' Adds a header to the page using the same name parameter
    SendTabsToFile (Name)
    DeleteCurrentSheet (Name)
 
    End Sub
   
   Sub RemoveOtherCoaches(ByVal Name As String) 'pass coach name to search for and filter by

Dim FirstRow As Long
Dim LastRow As Long
Dim Lrow As Long

FirstRow = ActiveSheet.UsedRange.Cells(1).Row
LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "F").End(xlUp).Row

For Lrow = LastRow To FirstRow Step -1
    
    If Not IsError(ActiveSheet.Cells(Lrow, "E").Value) Then ' fix this
        Select Case Name
        Case "Onirik" '--Branch for Distributors
            If Not ActiveSheet.Cells(Lrow, "F").Value Like Name And Not ActiveSheet.Cells(Lrow, "F").Row = "1" Then
                ActiveSheet.Cells(Lrow, "F").EntireRow.Delete
            End If
        Case "Fulton Hogan" '--Branch for Fulton Hogan (NZ & Aus).  Can't use wilds so making seperate branch
            If Not (ActiveSheet.Cells(Lrow, "C").Value Like "Fulton Hogan Au" Or ActiveSheet.Cells(Lrow, "C").Value Like "Fulton Hogan NZ") And Not ActiveSheet.Cells(Lrow, "C").Row = "1" Then
                ActiveSheet.Cells(Lrow, "C").EntireRow.Delete
            End If
        Case "Harrison Grierson", "CIGNA" '--Branch for other Companies
            If Not ActiveSheet.Cells(Lrow, "C").Value Like Name And Not ActiveSheet.Cells(Lrow, "C").Row = "1" Then
                ActiveSheet.Cells(Lrow, "C").EntireRow.Delete
            End If
        Case Else '--Branch for Coaches
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

' Levels_Passed_Formatting Macro:  Main method calls all of the other methods to create custom reports for each coach.
Sub MAIN()
    Application.ScreenUpdating = False
    Call SaveWorkingCopy
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
    Call DeleteExtraSheets
    
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True
    
    'Next Steps:
    ' for this iteration ignore FL certificates, but plan how to deal with them.
    

    
End Sub



