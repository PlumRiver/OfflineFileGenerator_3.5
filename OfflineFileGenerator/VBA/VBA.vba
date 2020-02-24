Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Application.CellDragAndDrop = True
End Sub

Private Sub Workbook_Open()
  Application.CellDragAndDrop = False
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.Range("C2").Value = "" And IsValidDate(Sh.Name) = True Then
        Sh.Range("C8:C9").NumberFormat = "ddmmmyyyy"
        Sh.Range("E8:E8").HorizontalAlignment = xlGeneral
        Sh.Range("E9:E9").HorizontalAlignment = xlGeneral
        Sh.Range("A1").Font.Color = vbWhite
        Sh.Range("A1").NumberFormat = "ddmmmyyyy"
        If Sh.Name <> "Order Template" Then
            Sh.Unprotect
            Sh.Select
            Range("C5").Select

            Sh.Range("C5").NumberFormat = "@"
            Sh.Range("C6").NumberFormat = "@"
            Sh.Range("C7").NumberFormat = "@"
            Sh.Range("C5").Value = ""
            Sh.Range("C6").Value = ""
            Sh.Range("C7").Value = ""
            Sh.Range("C8").Value = Sh.Range("A1").Value
            Sh.Range("C9").Value = DateAdd("d", CInt(Sheets("GlobalVariables").Range("B102").Value), CDate(Sh.Range("A1").Value))
            Sh.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim soldto As String
    soldto = ""
    For Each Sh In Worksheets
        If (Len(Sh.Name) = 9 And IsNumeric(Left(Sh.Name, 2)) = True And IsNumeric(Right(Sh.Name, 4)) = True) Then
            If Sh.Range("C5").Value <> "" Then
                If soldto = "" Then
                    soldto = Sh.Range("C5").Value
                Else
                    If soldto <> Sh.Range("C5").Value Then
                        MsgBox "The offline form doesn't support different sold-to customers on the tabs of a single spreadsheet", vbExclamation
                        Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
End Sub
