Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    If Target.Row > 11 And Target.Column > 7 And Target.Count = 1 Then
        If Target.Value <> "" Then
           Dim err As Boolean
           err = True
           If TypeName(Target.Value) = "Integer" Or TypeName(Target.Value) = "Double" Then
               Dim qty As Double
               Dim multiple As Integer
               qty = CDbl(Target.Value)
               multiple = CDbl(Sheets("OrderMultiple").Cells(Target.Row, Target.Column).Value)
               If qty Mod multiple = 0 And qty >= 0 And qty = CInt(Target.Value) Then
                   err = False
               End If
           End If
           If err = True Then
               MsgBox "You must enter a multiple of " + CStr(CDbl(Sheets("OrderMultiple").Cells(Target.Row, Target.Column).Value)) + " in this cell.", vbCritical, "Multiple Value Cell"
               Application.Undo
           End If
        End If
    End If
End Sub
