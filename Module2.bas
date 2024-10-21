Attribute VB_Name = "Module2"

Sub prepSheet()

    Set wS = Worksheets.Add(After:=Worksheets("Home Page"))
    wS.Name = "TripUploadv1"
    
    Sheets("TripUploadv1").Select
    
    Range("A1").Formula = "tripNumber"
    Range("B1").Formula = "contractName"
    Range("C1").Formula = "vehicleRegistrationNumber"
    Range("D1").Formula = "driverTagNumber"
    Range("E1").Formula = "siteExternalReference"
    Range("F1").Formula = "sequence"
    Range("G1").Formula = "arrivalDateTime"
    Range("H1").Formula = "depatureDateTime"
    Range("I1").Formula = "orderNumber"
    Range("J1").Formula = "taskTemplateNodeType"
    Range("K1").Formula = "instructions"
    Range("L1").Formula = "totalServiceTime"
    
    Columns("A:L").EntireColumn.AutoFit

End Sub
