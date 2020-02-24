
Function IsValidDate(ByVal DateStr As String) As Boolean

    On Error GoTo Invalid

    If Len(DateStr) < 11 Then
        'VariableDate = Mid(DateStr, 3, 3) + " " + Mid(DateStr, 1, 2) + " " + Mid(DateStr, 6)
        IsValidDate = DateSerial(CInt(Mid(DateStr, 6)), ConvertENMonthToNumber(Mid(DateStr, 3, 3)), CInt(Mid(DateStr, 1, 2)))
    Else
        'VariableDate = Mid(DateStr, 5, 3) + " " + Mid(DateStr, 3, 2) + " " + Mid(DateStr, 8)
        IsValidDate = DateSerial(CInt(Mid(DateStr, 8)), ConvertENMonthToNumber(Mid(DateStr, 5, 3)), CInt(Mid(DateStr, 3, 2)))
    End If

    IsValidDate = True
    Exit Function

Invalid:
    IsValidDate = False

End Function

Function ConvertToDate(ByVal DateStr As String) As Date
    If Len(DateStr) < 11 Then
        'VariableDate = Mid(DateStr, 3, 3) + " " + Mid(DateStr, 1, 2) + " " + Mid(DateStr, 6)
        ConvertToDate = DateSerial(CInt(Mid(DateStr, 6)), ConvertENMonthToNumber(Mid(DateStr, 3, 3)), CInt(Mid(DateStr, 1, 2)))
    Else
        'VariableDate = Mid(DateStr, 5, 3) + " " + Mid(DateStr, 3, 2) + " " + Mid(DateStr, 8)
        ConvertToDate = DateSerial(CInt(Mid(DateStr, 8)), ConvertENMonthToNumber(Mid(DateStr, 5, 3)), CInt(Mid(DateStr, 3, 2)))
    End If
    'ConvertToDate = CVDate(VariableDate)
End Function

Function ConvertENMonthToNumber(ByVal mth As String) As Integer
    Select Case mth
        Case "Jan"
            ConvertENMonthToNumber = 1
        Case "Feb"
            ConvertENMonthToNumber = 2
        Case "Mar"
            ConvertENMonthToNumber = 3
        Case "Apr"
            ConvertENMonthToNumber = 4
        Case "May"
            ConvertENMonthToNumber = 5
        Case "Jun"
            ConvertENMonthToNumber = 6
        Case "Jul"
            ConvertENMonthToNumber = 7
        Case "Aug"
            ConvertENMonthToNumber = 8
        Case "Sep"
            ConvertENMonthToNumber = 9
        Case "Oct"
            ConvertENMonthToNumber = 10
        Case "Nov"
            ConvertENMonthToNumber = 11
        Case "Dec"
            ConvertENMonthToNumber = 12
    End Select
End Function

