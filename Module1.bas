Attribute VB_Name = "Module1"

Sub createTripSheet()

    Dim wb1, wb2 As Workbook
    Dim rWs, oWs, cWs, dws, vWs, sWs, tuWs As Worksheet
    Dim btripStart, btripEnd As Boolean
    Dim actRow, rowCntrRprt, rowCntrOrders, rowCntrContr, rowCntrDriver, _
    rowCntrVeh, rowCntrSites, rowCntrTU, rowCntr1, rowCntr2, rowCntr3 As Long
    Dim tripNum, cntrCode, cntrName, fleetNum, vehicRegNum, driver, drvrTag, _
    tripStrtLoc, tripStrtExRef, tripLoading, tripOffloading, tripEndLoc, tripEndExRef, _
    ordrNum, tasktempNodeType, instructions As String
    Dim tempCntrt, tempOrdr, tempOrdrRow, tempDrvr, tempVehic, tempSites As String
    Dim dateArr, dateDepp, tripDate, tempDate As Date
    Dim seq, ordrIterator, uniqueNum As Integer
    Dim strDay, strMon, strYear As String
 
    Application.ScreenUpdating = False
        
    Set wb1 = ThisWorkbook
    StrPath = wb1.Path
    strname = wb1.FullName
    
    Set rWs = Sheets("Report")
    Set oWs = Sheets("Orders")
    Set cWs = Sheets("Contracts")
    Set dws = Sheets("Drivers")
    Set vWs = Sheets("Vehicles")
    Set sWs = Sheets("Sites")

    Application.DisplayAlerts = False
        For Each xWs In Application.ActiveWorkbook.Worksheets
            If xWs.Name <> "Home Page" And xWs.Name <> "Report" And xWs.Name <> "Orders" _
            And xWs.Name <> "MasterData" And xWs.Name <> "Drivers" And xWs.Name <> "Vehicles" _
            And xWs.Name <> "Contracts" And xWs.Name <> "Sites" Then
            xWs.Delete
            End If
        Next
    Application.DisplayAlerts = True
    
    Call Module2.prepSheet
    
    Set tuWs = Sheets("TripUploadv1")
    rWs.Select
    rowCntrRprt = Cells(Rows.Count, 2).End(xlUp).Row
    oWs.Select
    rowCntrOrders = Cells(Rows.Count, 2).End(xlUp).Row
    cWs.Select
    rowCntrContr = Cells(Rows.Count, 2).End(xlUp).Row
    dws.Select
    rowCntrDriver = Cells(Rows.Count, 2).End(xlUp).Row
    vWs.Select
    rowCntrVeh = Cells(Rows.Count, 2).End(xlUp).Row
    sWs.Select
    rowCntrSites = Cells(Rows.Count, 2).End(xlUp).Row
       
    rowCntrTU = 2
    actRow = 2
    uniqueNum = 1
    seq = 1
    
    
    For i = 2 To rowCntrRprt
        
        seq = 1
        ordrIterator = 0
        Sheets("Report").Select
        
        cntrCode = Range("B" & i).Value
        ordrNum = Range("C" & i).Value
        fleetNum = Range("D" & i).Value
        tripStrtLoc = Range("E" & i).Value
        tripEndLoc = Range("F" & i).Value
        driver = Range("G" & i).Value
        dateArr = Range("H" & i).Value
        dateDepp = Range("I" & i).Value
        
        If tripStrtLoc <> "" Then
            ordrIterator = ordrIterator + 1
            btripStart = True
            Else
                btripStart = False
        End If
        If tripEndLoc <> "" Then
            ordrIterator = ordrIterator + 1
            btripEnd = True
            Else
                btripEnd = False
        End If
        
        If dateArr <> "" Then
            strDay = Mid(dateArr, 9, 2)
            strMon = Mid(dateArr, 6, 2)
            strYear = Mid(dateArr, 1, 4)
            tripDate = strDay & "." & strMon & "." & strYear
        End If
        
        If dateDepp <> "" Then
            strDay = Mid(dateDepp, 9, 2)
            strMon = Mid(dateDepp, 6, 2)
            strYear = Mid(dateDepp, 1, 4)
            tripDate = strDay & "." & strMon & "." & strYear
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''Get Contract Name''''''''''''''''''''
        For j = 2 To rowCntrContr
            cWs.Select
            tempCntrt = Range("A" & j).Value
                If tempCntrt = cntrCode Then
                    cntrName = Range("B" & j).Value
                    Exit For
                End If
        Next
        
        '''''''''''''''''''''''''''''''''''''''''''''(Get Vehicle Reg)''''''''''''''''''''
        For m = 2 To rowCntrVeh
            vWs.Select
            tempVehic = Range("B" & m).Value
                If tempVehic = fleetNum Then
                    vehicRegNum = Range("A" & m).Value
                    Exit For
                End If
        Next
                
        '''''''''''''''''''''''''''''''''''''''''''''''Get Order Load/Unloading'''''''''''''
        For k = 2 To rowCntrOrders
            oWs.Select
            tempOrdr = Range("A" & k).Value
                If tempOrdr = ordrNum Then
                    tripLoading = Range("C" & k).Value
                    tripOffloading = Range("E" & k).Value
                    ordrIterator = ordrIterator + 2
                    Exit For
                End If
        Next
        
        ''''''''''''''''''''''''''''''''''''''''''''''''Get Driver'''''''''''''''''''''''''''
        For l = 2 To rowCntrDriver
            dws.Select
            tempDrvr = Range("H" & l).Value
                If tempDrvr = driver Then
                    drvrTag = Range("G" & l).Value
                    Exit For
                End If
        Next
        
        '''''''''Put into TripUpload Sheet - Don't ask how I dunno I just built it''''''''''''
        
        tuWs.Select
        rowCntr1 = Cells(Rows.Count, 1).End(xlUp).Row
        
        For x = 1 To ordrIterator
            tuWs.Select
            
            Range("A" & x + rowCntr1).Formula = cntrCode & "-" & tripDate ''& "-00" & uniqueNum
            Range("B" & x + rowCntr1).Formula = cntrName
            Range("C" & x + rowCntr1).Formula = vehicRegNum
            Range("D" & x + rowCntr1).Formula = "'" & drvrTag
            
            If ordrIterator = 4 And x = 1 Then
                Range("E" & x + rowCntr1).Formula = tripStrtLoc
                Range("E" & x + rowCntr1 + 1).Formula = tripLoading
                Range("J" & x + rowCntr1 + 1).Formula = "Loading"
                Range("E" & x + rowCntr1 + 2).Formula = tripOffloading
                Range("J" & x + rowCntr1 + 2).Formula = "Offloading"
                Range("E" & x + rowCntr1 + 3).Formula = tripEndLoc
            End If
            
            If ordrIterator = 3 And x = 1 Then
                If btripStart = True Then
                        Range("E" & x + rowCntr1).Formula = tripStrtLoc
                        Range("E" & x + rowCntr1 + 1).Formula = tripLoading
                        Range("J" & x + rowCntr1 + 1).Formula = "Loading"
                        Range("E" & x + rowCntr1 + 2).Formula = tripOffloading
                        Range("J" & x + rowCntr1 + 2).Formula = "Offloading"
                    Else
                            If btripEnd = True Then
                                Range("E" & x + rowCntr1).Formula = tripLoading
                                Range("J" & x + rowCntr1).Formula = "Loading"
                                Range("E" & x + rowCntr1 + 1).Formula = tripOffloading
                                Range("J" & x + rowCntr1 + 1).Formula = "Offloading"
                                Range("E" & x + rowCntr1 + 2).Formula = tripEndLoc
                            End If
                End If
            End If
            
            If ordrIterator = 2 And x = 1 Then
                Range("E" & x + rowCntr1).Formula = tripLoading
                Range("J" & x + rowCntr1).Formula = "Loading"
                Range("E" & x + rowCntr1 + 1).Formula = tripOffloading
                Range("J" & x + rowCntr1 + 1).Formula = "Offloading"
            End If
            
            Range("F" & x + rowCntr1).Formula = seq
            
            If dateArr = "" And Range("F" & x + rowCntr1) = 1 Then
                Range("H" & x + rowCntr1).Formula = dateDepp
                Else
                    If dateDepp = "" And Range("J" & x + rowCntr1) = "Loading" Then
                        Range("G" & x + rowCntr1).Formula = dateArr
                    End If
            End If
            
            If Range("J" & x + rowCntr1) = "Loading" Or Range("J" & x + rowCntr1) = "Offloading" Then
                Range("I" & x + rowCntr1).Formula = ordrNum
            End If
            
            Range("L" & x + rowCntr1).Formula = "60"
            seq = seq + 1
            
        Next
        uniqueNum = uniqueNum + 1
    Next
    
    Application.ScreenUpdating = True
    Call Module3.uniqueTripNum
    Sheets("TripUploadv1").Select
    Columns("A:L").EntireColumn.AutoFit
    
    Sheets("Home Page").Select
    
End Sub
