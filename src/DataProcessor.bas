Attribute VB_Name = "DataProcessor"
' Professional Data Cleaning and Formatting Module
Sub CleanAndFormatData_Professional()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("RawData")
    Dim dataRange As Range
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set dataRange = ws.Range("A1").CurrentRegion
    dataRange.Replace What:=" ", Replacement:="", LookAt:=xlPart
    
    With dataRange
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(240, 240, 240)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    dataRange.FormatConditions.Delete
    dataRange.FormatConditions.AddUniqueValues
    dataRange.FormatConditions(1).DupeUnique = xlDuplicate
    dataRange.FormatConditions(1).Font.Color = vbRed

    MsgBox "Process Complete.", vbInformation
    
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub
