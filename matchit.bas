Function matchit(ByVal refNumber) As String
    Dim rng1 As Range, wb As Workbook
    Set wb = ActiveWorkbook
    Set rng1 = wb.ActiveSheet.Range("J3")
    If refNumber >= rng1.Offset(1, 0).Value And refNumber <= rng1.Offset(1, 1).Value Then
        matchit = rng1.Offset(1, 2).Value
    ElseIf refNumber >= rng1.Offset(2, 0).Value And refNumber <= rng1.Offset(2, 2).Value Then
        matchit = rng1.Offset(2, 2).Value
    ElseIf refNumber >= rng1.Offset(3, 0).Value And refNumber <= rng1.Offset(3, 3).Value Then
        matchit = rng1.Offset(3, 2).Value
    ElseIf refNumber >= rng1.Offset(4, 0).Value Then
        matchit = rng1.Offset(4, 2).Value
    End If

End Function
