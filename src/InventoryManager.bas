Attribute VB_Name = "InventoryManager"
' Dynamic Inventory Update Module
Sub UpdateInventory_Dynamic()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("RawData")
    Dim idCol As Long, stockCol As Long, reorderCol As Long
    Dim prodID As String, foundCell As Range

    On Error Resume Next
    idCol = ws.Rows(1).Find("Product Id").Column
    stockCol = ws.Rows(1).Find("Stock Quantit").Column
    reorderCol = ws.Rows(1).Find("Reorder level").Column
    On Error GoTo 0

    If idCol = 0 Or stockCol = 0 Then Exit Sub

    prodID = Trim(InputBox("Enter Product ID:"))
    If prodID = "" Then Exit Sub

    Set foundCell = ws.Columns(idCol).Find(prodID, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        ' Example: reducing stock by 1
        ws.Cells(foundCell.Row, stockCol).Value = ws.Cells(foundCell.Row, stockCol).Value - 1
        MsgBox "Inventory Updated.", vbInformation
    End If
End Sub
