Attribute VB_Name = "Module3"

Sub uniqueTripNum()

    Dim tripNumOrg, tripNumMod, tripNumNew As String
    Dim rowCntr1, rowCntr2, Cntr2, uniqueNum As Long
    Dim fgCntr, ld1Cntr, ld2Cntr, oilCntr, sppCntr, _
    locCntr, acdCntr, dbCntr, ansCntr, mlkCntr, crtCntr, _
    crbcntr, ib40Cntr, ib41Cntr, ib42Cntr As Integer
    Dim cntrIdentif, siteName, siteNameSrch, taskTempNode As String
    
    fgCntr = 1        ''''''1
    ld1Cntr = 1       ''''''2
    ld2Cntr = 1       ''''''3
    oilCntr = 1       ''''''4
    sppCntr = 1       ''''''5
    locCntr = 1       ''''''6
    acdCntr = 1       ''''''7
    ansCntr = 1       ''''''8
    mlkCntr = 1       ''''''9
    crtCntr = 1       ''''''10
    crbcntr = 1       ''''''11
    ib40Cntr = 1      ''''''12
    ib41Cntr = 1      ''''''13
    ib42Cntr = 1      ''''''14
    
    Sheets("TripUploadv1").Select
    uniqueNum = 1
    rowCntr1 = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To rowCntr1
Start:
        tripNumOrg = Range("A" & i).Value
        cntrIdentif = Mid(tripNumOrg, 1, 5)
        
        If cntrIdentif = "UBLD1" Then    ''''''''''1
            uniqueNum = ld1Cntr
        End If
        If cntrIdentif = "UBLD2" Then    ''''''''''2
            uniqueNum = ld2Cntr
        End If
        If cntrIdentif = "UBOIL" Then    ''''''''''3
            uniqueNum = oilCntr
        End If
        If cntrIdentif = "UBSPP" Then    ''''''''''4
            uniqueNum = sppCntr
        End If
        If cntrIdentif = "UBLOC" Then    ''''''''''5
            uniqueNum = locCntr
        End If
        If cntrIdentif = "UBACD" Then    ''''''''''6
            uniqueNum = acdCntr
        End If
        If cntrIdentif = "UBANS" Then    ''''''''''7
            uniqueNum = ansCntr
        End If
        If cntrIdentif = "UBFG-" Then    ''''''''''8
            uniqueNum = fgCntr
        End If
        If cntrIdentif = "UBMLK" Then    ''''''''''9
            uniqueNum = mlkCntr
        End If
        If cntrIdentif = "UBDB-" Then    ''''''''''10
            uniqueNum = dbCntr
        End If
        If cntrIdentif = "NOCRT" Then    ''''''''''11
            uniqueNum = crtCntr
        End If
        If cntrIdentif = "UNIB40" Then    ''''''''''12
            uniqueNum = ib40Cntr
        End If
        If cntrIdentif = "UNIB41" Then    ''''''''''13
            uniqueNum = ib41Cntr
        End If
        If cntrIdentif = "UNIB42" Then    ''''''''''14
            uniqueNum = ib42Cntr
        End If
                
NextLine:
        If Range("F" & i + 1).Value <> 1 Then
            tripNumOrg = Range("A" & i).Value
                If uniqueNum < 10 Then
                    tripNumMod = tripNumOrg & "-00" & uniqueNum
                    Range("A" & i).Formula = tripNumMod
                    Else
                        If uniqueNum >= 10 Then
                        tripNumMod = tripNumOrg & "-0" & uniqueNum
                        Range("A" & i).Formula = tripNumMod
                        End If
                End If
                i = i + 1
                If i > rowCntr1 Then
                    Exit For
                End If
GoTo NextLine
            Else
                Range("A" & i).Formula = tripNumMod
                            
                If cntrIdentif = "UBLD1" Then    ''''''''''1
                    ld1Cntr = ld1Cntr + 1
                End If
                If cntrIdentif = "UBLD2" Then    ''''''''''2
                    ld2Cntr = ld2Cntr + 1
                End If
                If cntrIdentif = "UBOIL" Then    ''''''''''3
                    oilCntr = oilCntr + 1
                End If
                If cntrIdentif = "UBSPP" Then    ''''''''''4
                    sppCntr = sppCntr + 1
                End If
                If cntrIdentif = "UBLOC" Then    ''''''''''5
                    locCntr = locCntr + 1
                End If
                If cntrIdentif = "UBACD" Then    ''''''''''6
                    acdCntr = acdCntr + 1
                End If
                If cntrIdentif = "UBANS" Then    ''''''''''7
                    ansCntr = ansCntr + 1
                End If
                If cntrIdentif = "UBFG-" Then    ''''''''''8
                    fgCntr = fgCntr + 1
                End If
                If cntrIdentif = "UBMLK" Then    ''''''''''9
                    mlkCntr = mlkCntr + 1
                End If
                If cntrIdentif = "UBDB-" Then    ''''''''''10
                    dbCntr = dbCntr + 1
                End If
                If cntrIdentif = "NOCRT" Then    ''''''''''11
                    crtCntr = crtCntr + 1
                End If
                If cntrIdentif = "UNIB40" Then    ''''''''''12
                    ib40Cntr = ib40Cntr + 1
                End If
                If cntrIdentif = "UNIB41" Then    ''''''''''13
                    ib41Cntr = ib41Cntr + 1
                End If
                If cntrIdentif = "UNIB42" Then    ''''''''''14
                    ib42Cntr = ib42Cntr + 1
                End If
        End If
    Next
    Sheets("Sites").Select
    rowCntr2 = Cells(Rows.Count, 1).End(xlUp).Row
    
    For j = 2 To rowCntr1
        Sheets("TripUploadv1").Select
        siteName = Range("E" & j).Value
        If Range("J" & j).Value = "" Then
            For k = 2 To rowCntr2
                Sheets("Sites").Select
                siteNameSrch = Range("A" & k).Value
                If siteNameSrch = siteName Then
                    taskTempNode = Range("C" & k).Value
                    Exit For
                End If
             Next
             Sheets("TripUploadv1").Select
            Range("J" & j).Formula = taskTempNode
        End If
    Next
    
    
    
End Sub
