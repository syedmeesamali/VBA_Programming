Attribute VB_Name = "Summary_Concise"
Option Explicit

Sub Summarize_Concise(wb As Workbook, percentVal As Integer)

Dim ws As Worksheet
Dim yesno As Boolean
Dim newRng As Range
Dim Count As Integer, counter As Integer, nCount As Integer, mCount As Integer

Dim costVal As Variant, priceVal As Variant
Dim valMat As Variant, valTotal As Variant, valMH As Variant
Dim valArea As Variant, valRate As Variant, valTotalCost As Variant
Dim valTotalPrice1 As Variant, valTotalPrice2 As Variant
Dim valCons As Variant, valTrans As Variant, valTools As Variant, rCell As Range
Dim valDate As Variant, valQtn As Variant, valClient As Variant, valSysName As Variant
Dim valUnit As Variant, valRev As Variant
Dim startRow As Integer, endRow As Integer, Diff As Integer
Dim locDesc As Range, locTotal As Range

Dim LSearchRow As Integer
Count = 0
Application.ScreenUpdating = False
Set newRng = wb.Sheets("Totals").Range("a65536").End(xlUp).Offset(1, 0)

On Error Resume Next
For Each ws In wb.Worksheets
        With ws
            ws.Unprotect
            'CHECK FOR COATING TYPE ESTIMATION SHEETS.
            If InStr(1, .Range("A2").Value, "ESTIMATION") > 0 Then
                valArea = .Range("B5").Value
                valClient = .Range("B4").Value
                valRev = .Range("E3").Value
                Count = Count + 1
                For Each rCell In .Range("B11:B35")
                  If InStr(1, rCell.Value, "Total") > 0 Then
                    valMat = rCell.Offset(0, 6).Value
                    valCons = rCell.Offset(0, 7).Value
                    valMH = rCell.Offset(0, 8).Value
                    valTools = rCell.Offset(0, 10).Value
                    valTrans = rCell.Offset(0, 11).Value
                    valTotal = rCell.Offset(0, 12).Value
                  Else
                  End If
                Next rCell
                
                For Each rCell In .Range("D3:J5")
                  If InStr(1, rCell.Value, "QTN") > 0 Then
                    valQtn = rCell.Offset(0, 1).Value
                    valDate = rCell.Offset(1, 1).Value
                    Else
                  End If
                Next rCell
         
         counter = wb.Sheets("Totals").Range("A1").Value
         'CALL THE DISPLAY ROUTINE.
         Call DisplayResult_Detailed(Count, ws, valTotal, valMat, newRng, valMH, _
         valArea, percentVal, priceVal, valCons, valTrans, valTools, valQtn, _
         valDate, valClient, counter, valRev)
         
         'FOR INJECTION SHEETS
         ElseIf InStr(1, .Range("B1").Value, "Project") > 0 Or InStr(1, .Range("B1").Value, "DicoTech") > 0 Then
            Count = Count + 1
                For Each rCell In .Range("B2:B55")
                  If InStr(1, rCell.Value, "Injectors") > 0 Then
                    valArea = rCell.Offset(-1, 1).Value
                  ElseIf InStr(1, rCell.Value, "Total Man") > 0 Then
                    valMH = rCell.Offset(0, 1).Value
                    valRate = rCell.Offset(0, 3).Value
                  ElseIf InStr(1, rCell.Value, "Material Cost") > 0 Then
                    valMat = rCell.Offset(0, 1).Value
                  ElseIf InStr(1, rCell.Value, "Tools") > 0 Then
                    valTools = rCell.Offset(0, 1).Value
                  ElseIf InStr(1, rCell.Value, "Total Price:") > 0 Then
                    valCons = rCell.Offset(0, 1).Value
                  ElseIf InStr(1, rCell.Value, "Consumables") > 0 Then
                    valCons = rCell.Offset(0, 1).Value
                  ElseIf InStr(1, rCell.Value, "Transportation") > 0 Then
                    valTrans = rCell.Offset(0, 1).Value
                  Else
                  End If
                Next rCell
             
             'Injection works display routine
             Call DisplayResult_Injection(Count, ws, valTotal, valMat, newRng, valMH, _
             valTotalCost, valArea, percentVal, priceVal, valCons, valTrans, valTools, _
             valRate)
             
            'FOR CIVIL WORKS SHEETS
            ElseIf InStr(1, .Range("A2").Value, "CIVIL WORKS") > 0 Then
            
                Set locDesc = .Range("B7:B30").Find(What:="Description", LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows)
                Set locTotal = .Range("B7:B35").Find(What:="Total", LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows)
                startRow = locDesc.Row
                endRow = locTotal.Row
                Diff = endRow - startRow
                For nCount = 1 To Diff
                  If Not IsEmpty(locDesc.Offset(nCount)) Then
                    Count = Count + 1
                    valSysName = locDesc.Offset(nCount)
                    valUnit = locDesc.Offset(nCount, 1)
                    valMat = locDesc.Offset(nCount, 6)
                    valArea = locDesc.Offset(nCount, 7)
                    valCons = locDesc.Offset(nCount, 8).Value
                    valMH = locDesc.Offset(nCount, 9).Value
                    valTools = locDesc.Offset(nCount, 11).Value
                    valTrans = 0
                    valTotal = locDesc.Offset(nCount, 12).Value
                    
                  Else
                  End If
                
                  'Civil works display routine
                  Call DisplayResult_Civil(Count, ws, valTotal, valMat, newRng, valMH, _
                    valArea, percentVal, priceVal, valCons, valTrans, valTools, valQtn, _
                    valDate, valClient, counter, valSysName)
                Next nCount
             
             'FOR CIVIL WORKS SHEETS
            ElseIf InStr(1, .Range("A2").Value, "PRELIMINARIES") > 0 Then
            
                Set locDesc = .Range("B7:B30").Find(What:="Description", LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows)
                Set locTotal = .Range("B7:B35").Find(What:="Total", LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows)
                startRow = locDesc.Row
                endRow = locTotal.Row
                Diff = endRow - startRow
                For mCount = 1 To Diff
                  If Not IsEmpty(locDesc.Offset(mCount)) Then
                    Count = Count + 1
                    valSysName = locDesc.Offset(mCount)
                    valUnit = locDesc.Offset(mCount, 1)
                    valMat = 0
                    valArea = locDesc.Offset(mCount, 3)
                    valCons = 0
                    valMH = 0
                    valTools = 0
                    valTrans = 0
                    valTotal = locDesc.Offset(mCount, 7).Value
                    
                  Else
                  End If
                
                  'Civil works display routine
                  Call DisplayResult_Civil(Count, ws, valTotal, valMat, newRng, valMH, _
                    valArea, percentVal, priceVal, valCons, valTrans, valTools, valQtn, _
                    valDate, valClient, counter, valSysName)
                Next mCount
             
             Else      'End of sheet selection criteria
             End If
    End With
Next
         newRng.Offset(Count + 3, 1).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 6).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 8).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 3).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 3).Interior.Color = vbYellow
         newRng.Offset(Count + 3, 3).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 3).NumberFormat = "#,##0"
         
         newRng.Offset(Count + 3, 4).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 4).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 4).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 4).NumberFormat = "#,##0"
         
         newRng.Offset(Count + 3, 5).FormulaR1C1 = "=RC[-2]-SUM(RC[4]:RC[7])"
         newRng.Offset(Count + 3, 5).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 5).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 5).NumberFormat = "#,##0"
         
         newRng.Offset(Count + 3, 7).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 7).Interior.Color = vbYellow
         newRng.Offset(Count + 3, 7).EntireColumn.AutoFit
         
         newRng.Offset(Count + 3, 9).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 9).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 9).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 10).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 10).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 10).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 11).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 11).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 11).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 12).FormulaR1C1 = "=SUM(R[-1]C:R[-" & Count & "]C)"
         newRng.Offset(Count + 3, 12).Interior.Color = vbGreen
         newRng.Offset(Count + 3, 12).EntireColumn.AutoFit
         
         newRng.Offset(Count + 3, 8).EntireColumn.AutoFit
         newRng.Offset(Count + 3, 8).FormulaR1C1 = "=(RC[-1]-RC[-5])/RC[-1]"
         newRng.Offset(Count + 3, 8).NumberFormat = "#.##%"
         newRng.Offset(Count + 3, 8).Interior.Color = vbYellow
         newRng.Offset(Count + 3, 8).Font.Color = vbRed
         newRng.Offset(Count + 3, 8).CurrentRegion.Borders.LineStyle = xlContinuous
         newRng.Offset(Count + 3, 8).CurrentRegion.BorderAround = True
         newRng.Offset(Count + 3, 8).EntireRow.Font.Bold = True

         newRng.Offset(2, 6).Value = "PRICE"
        ' "=concatenate(" & """Price @ """ & "," & "TEXT(" & "R[" & count + 1 & "]C[2]" & "," & """##%""" & ")" & ")"
        newRng.Offset(2, 6).Font.Bold = True
        newRng.Offset(2, 6).EntireColumn.AutoFit
         
Application.ScreenUpdating = True
Sheets("Totals").Activate

End Sub
