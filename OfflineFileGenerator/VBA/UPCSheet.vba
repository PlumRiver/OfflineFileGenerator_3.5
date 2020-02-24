Private Sub Worksheet_Activate()
    
    Dim FirstRange As Range
    Dim SKUCell As Range
    LastRow = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
    If LastRow = 1 Then
        Exit Sub
    End If
    SKULastRow = Sheets("WebSKU").Range("L" & Sheets("WebSKU").Rows.Count).End(xlUp).Row
    For i = 2 To LastRow Step 1
        sku = ActiveSheet.Range("H" + Trim(CStr(i))).Value
        sku = Trim(sku)
        skuQty = 0
        If sku <> "" Then
            Set FirstRange = Sheets("WebSKU").Range("A1:AZ" & SKULastRow)
        #If Mac Then
            Set SKUCell = FirstRange.Find(What:=sku, LookIn:=xlValues, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False)
        #Else
            Set SKUCell = FirstRange.Find(What:=sku, LookIn:=xlValues, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
        #End If
            If Not SKUCell Is Nothing Then
                SKUAddress = SKUCell.Address
                For Each Sh In Worksheets
                    If (Len(Sh.Name) = 9 And IsNumeric(Left(Sh.Name, 2)) = True And IsNumeric(Right(Sh.Name, 4)) = True) Then
                        SKUQtyValue = Sh.Range(SKUAddress).Value
                        If SKUQtyValue <> "" And IsNumeric(SKUQtyValue) Then
                           skuQty = skuQty + SKUQtyValue
                        End If
                    End If
                Next
            End If
        End If
        If skuQty > 0 Then
            ActiveSheet.Range("J" + CStr(i)).Value = CStr(skuQty)
        End If
    Next
    
End Sub

