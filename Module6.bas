Attribute VB_Name = "Module6"

Sub vehicleTypeAPIcall()
    
    Dim wS As Worksheet
    Dim shtUrl As String
    Dim shtApiKey As String
    Dim pause As Integer
    Dim rng As Range
    Dim fleetNumber, regNumber, typeName As String
    Dim str1, str2, str3, strflt, strreg, strname, strdescrip As String
    Dim actRow, rowCntr1, rowCntr2, rowCntr3 As Long
    
    Application.DisplayAlerts = False
        For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> "Home Page" And xWs.Name <> "Report" And xWs.Name <> "Orders" _
            And xWs.Name <> "MasterData" And xWs.Name <> "Drivers" And xWs.Name <> "Vehicles" _
            And xWs.Name <> "Contracts" And xWs.Name <> "Sites" And xWs.Name <> "TripUploadv1" Then
            xWs.Delete
            End If
        Next
    Application.DisplayAlerts = True
    
    Set wS = Worksheets.Add(After:=Worksheets("MasterData"))
    wS.Name = "APIoutput"
    Sheets("APIoutput").Select
    
    Application.ScreenUpdating = False
    shtUrl = "https://onelogix.onroadvantage.com/api/vehicle?perPage=900"
    spliter = "active"
    pause = 0
    shtApiKey = "WyI4NGRjY2NhNmE4MWI0MTA3ODQ0ZDRiZTEyZGNmYzEyZCJd.Yt__Xw._4LHF3CZzl0eEgr8b-_Q6hsvZ3g"
    postcell = ""
    
    On Error Resume Next
    Dim GetResponse As String
    Url = shtUrl
    GetResponse = ""
            If Url Like "*http*" Then
                    pst = ""
                    zapr = "GET"
                Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
                 xmlhttp.Open zapr, Url, False: DoEvents
                 xmlhttp.SetRequestHeader "Authentication-Token", shtApiKey
                 xmlhttp.Send pst
                 GetResponse = xmlhttp.ResponseText
                 'MsgBox GetResponse
                 Set xmlhttp = Nothing
            End If
    
            If pause > 0 Then
                t = Timer
                Do
                    DoEvents
                    If t > Timer Then Exit Do
                Loop Until Timer - t > pause
            End If
          
            s = Split(GetResponse, spliter)
            If UBound(s) > 0 Then
            Sheets("APIoutput").Select
            Cells.Clear
            actRow = 1
                   For i = 1 To UBound(s)
actRow = actRow + 1
NextLine:
                        ps = Split(s(i), ",")
                        For j = 0 To UBound(ps)
                            str2 = ps(j)
                            If InStr(str2, "fleetNumber") Then
                                str1 = str2
                                strflt = Replace(Mid(str1, 13, 25), ":", "")
                                If strflt = "null" And i < UBound(s) Then
                                    i = i + 1
                                    GoTo NextLine
                                Else
                                    Cells(actRow, 1).Value = strflt
                                    End If
                            End If
                            
                            If j < 8 And InStr(str2, "description") Then
                                    str1 = str2
                                    strdescrip = Mid(str1, 15, 40)
                                    Cells(actRow, 2).Value = strdescrip
                            End If
                            
                            If InStr(str2, "registrationNumber") Then
                                    str1 = str2
                                    strreg = Mid(str1, 21, 31)
                                    Cells(actRow, 3).Value = strreg
                            End If
                            
                            If j > 18 And InStr(str2, "name") Then
                                str1 = str2
                                strname = Replace(Mid(str1, 8, 25), "}", "")
                                    Cells(actRow, 4).Value = strname
                            End If
                            
                        Next
                    Next
                Else:
                    Sheets("Sheet1").Select
                    Range("D1").Formula = "not available":
                End If
                
                Sheets("APIoutput").Select
                Range("A1").Formula = "FleetNumber"
                Range("B1").Formula = "Description"
                Range("C1").Formula = "Registration Number"
                Range("D1").Formala = "Name/Type"
                rowCntr1 = Cells(Rows.Count, 2).End(xlUp).Row
                For Each rng In Range("A1:D" & rowCntr1)
                    rng.Value = Replace(rng, """", "")
                    rng.Value = Replace(rng, ":", "")
                Next rng
                Columns("A:D").EntireColumn.AutoFit
        Application.ScreenUpdating = True
        
        '''''''''''''''After API call is completed match Types to Vehicles in Report "if report contains info"''''''''''''
End Sub

Sub matching()

        Sheets("APIoutput").Select
        rowCntr3 = Cells(Rows.Count, 1).End(xlUp).Row
        
        Sheets("Report").Select
        rowCntr2 = Cells(Rows.Count, 3).End(xlUp).Row
        
        Dim fleetNumReport, fleetNumApi, vehicleTypeReport, vehicleTypeApi As String
        
        For l = 2 To rowCntr2
            Sheets("Report").Select
            fleetNumReport = Range("D" & l).Value
            
            For m = 2 To rowCntr3
                Sheets("APIoutput").Select
                fleetNumApi = Range("A" & m).Value
                If fleetNumApi = fleetNumReport Then
                    vehicleTypeApi = Range("C" & m).Value
                    Exit For
                End If
            Next
            
            Sheets("Report").Select
            Range("J" & l).Formula = vehicleTypeApi
            
        Next
        
        Application.ScreenUpdating = True
                
End Sub
          

