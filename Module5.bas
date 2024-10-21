Attribute VB_Name = "Module5"

Sub reset()

    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "Home Page" And xWs.Name <> "Report" And xWs.Name <> "Orders" _
            And xWs.Name <> "MasterData" And xWs.Name <> "Drivers" And xWs.Name <> "Vehicles" _
            And xWs.Name <> "Contracts" And xWs.Name <> "Sites" Then
            xWs.Delete
        End If
    Next
    
    Sheets("Home Page").Select

End Sub

Sub clearTripsheet()

    Dim rowCntr1 As Long
    Sheets("TripUploadv1").Select
    rowCntr1 = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A2:N" & rowCntr1).ClearContents
    Sheets("Home Page").Select
    
End Sub
