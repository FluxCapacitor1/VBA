Attribute VB_Name = "Module4"

Sub tripUploadGen()

    Dim wb1, wb2 As Workbook
    Dim wS, xWs As Worksheet
    
    Dim Cntr1, Cntr2, Cntr3 As Integer
    Dim rowCntr1, rowCntr2, rowCntr3 As Long
    Dim DateNow As Date
    Dim DateStr, DateMod, DateSave As String
    Dim CellVal1, CellVal2, CellVal3 As String
    Dim filePath As String
    Dim tempPath, finalPath As String
    Dim wbName1, wbName2 As String
    
    DateNow = Now
    DateSave = Replace(CStr(DateNow), "/", "-")
    DateSave = Left(Replace(CStr(DateSave), ":", "-"), 16)
    DateMod = Left(DateSave, 10)
        
    Set wb1 = ThisWorkbook
    wbName1 = wb1.Name
    wb1.Activate
    Sheets("TripUploadv1").Select
    Set Range1 = Range("A2:A5")
    For Each Cell In Range1
        If IsEmpty(Cell) Then
            Sheets("Home Page").Select
            MsgBox ("The 'TripUploadv1' Sheet does not contain any data.")
            Exit Sub
        Exit For
        End If
    Next
    Workbooks.Add
    Set wb2 = ActiveWorkbook
    
    Dim fldrD As FileDialog
    Dim sStr As String
    Set fldrD = Application.FileDialog(msoFileDialogFolderPicker)
        With fldrD
            .Title = "Select a Folder"
            .AllowMultiSelect = False
            If .Show <> -1 Then GoTo NextCode
                sStr = .SelectedItems(1)
            End With
NextCode:
        filePath = sStr
        Set fldrD = Nothing

    finalPath = filePath & "\UB TripUpload File " & DateSave & ".xlsx"
        
    wbName2 = wb2.Name
    Workbooks(wbName2).Activate
    Set wS = Worksheets.Add
    wS.Name = "TripUploadv1"
    
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> "TripUploadv1" Then
            xWs.Delete
            End If
        Next
    Application.DisplayAlerts = True
    
    Workbooks(wbName1).Activate
    Sheets("TripUploadv1").Select
    rowCntr1 = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A1:L" & rowCntr1).Copy
    
    Workbooks(wbName2).Activate
    Sheets("TripUploadv1").Select
    Range("A1").PasteSpecial
    rowCntr2 = Cells(Rows.Count, 1).End(xlUp).Row
    Columns("A:O").EntireColumn.AutoFit
       
    ActiveWorkbook.SaveAs finalPath
    
    Sheets("Home Page").Select
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




