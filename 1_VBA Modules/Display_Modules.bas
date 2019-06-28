Option Explicit

'----------------DISPLAY ROUTINE FOR NORMAL SUMMARY-------------
Sub DisplayResult_Detailed(i As Integer, ws As Worksheet, valTotal As Variant, _
    valMat As Variant, newRng As Range, valMH As Variant, _
    valArea As Variant, pVal As Integer, priceVal As Variant, _
    consVal As Variant, transVal As Variant, toolsVal As Variant, valQtn As Variant, _
    valDate As Variant, valClient As Variant, counter As Integer, valRev As Variant)
        '------------DISPLAY SETTING FOR VALUES----------------
        
        valMH = Application.WorksheetFunction.RoundUp(valMH, 2)
        consVal = Application.WorksheetFunction.RoundUp(consVal, 2)
        toolsVal = Application.WorksheetFunction.RoundUp(toolsVal, 2)
        transVal = Application.WorksheetFunction.RoundUp(transVal, 2)
        
        If counter < 2 Then
            Range(newRng.Offset(1, 0), newRng.Offset(1, 9)).Font.Color = vbBlue
            newRng.Offset(1, 0).Value = "QTN # " & valQtn
            newRng.Offset(1, 1).Value = "REV:" & valRev
            newRng.Offset(1, 2).Value = "DATED: "
            newRng.Offset(1, 3).Value = valDate
            newRng.Offset(1, 3).Value = Format(valDate, "dd-mm-yy")
            newRng.Offset(1, 6).Value = "CLIENT: "
            newRng.Offset(1, 7).Value = valClient
        Else
        End If
        
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Interior.Color = vbBlack
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Font.Color = vbWhite
        newRng.Offset(2, 0).Value = "System Name"
        newRng.Offset(2, 0).Font.Bold = True
        newRng.Offset(2, 1).Value = "Mat Cost"
        newRng.Offset(2, 1).Font.Bold = True
        newRng.Offset(2, 2).Value = "Unit Cost"
        newRng.Offset(2, 2).Font.Bold = True
        
        newRng.Offset(2, 3).Value = "Total Cost"
        newRng.Offset(2, 3).Font.Bold = True
        newRng.Offset(2, 4).Value = "Manhours"
        newRng.Offset(2, 4).Font.Bold = True
        newRng.Offset(2, 5).Value = "Total QTY"
        newRng.Offset(2, 5).Font.Bold = True
        
        newRng.Offset(2, 6).Value = "Price at " & pVal & "%"
        newRng.Offset(2, 6).Font.Bold = True
        newRng.Offset(2, 7).Value = "Total Price"
        newRng.Offset(2, 7).Font.Bold = True
        newRng.Offset(2, 8).Value = "%age"
        newRng.Offset(2, 8).Font.Bold = True
        
        newRng.Offset(2, 9).Value = "Mat."
        newRng.Offset(2, 9).Font.Bold = True
        newRng.Offset(2, 10).Value = "Trans."
        newRng.Offset(2, 10).Font.Bold = True
        newRng.Offset(2, 11).Value = "T & E"
        newRng.Offset(2, 11).Font.Bold = True
        newRng.Offset(2, 12).Value = "Cons."
        newRng.Offset(2, 12).Font.Bold = True
        
        '----------ACTUAL VALUE CALCULATIONS------------
        newRng.Offset(i + 2, 0).Value = ws.Name
        newRng.Offset(i + 2, 0).EntireColumn.AutoFit
        newRng.Offset(i + 2, 1).Value = Format(valMat, "###.00")
        newRng.Offset(i + 2, 1).EntireColumn.AutoFit
        newRng.Offset(i + 2, 2).Value = Format(valTotal, "###.00")
        newRng.Offset(i + 2, 2).EntireColumn.AutoFit
        
        newRng.Offset(i + 2, 3).FormulaR1C1 = "=ROUNDUP(RC[2]*RC[-1],2)"
        newRng.Offset(i + 2, 3).NumberFormat = "#,##"
        newRng.Offset(i + 2, 3).EntireColumn.AutoFit
        newRng.Offset(i + 2, 4).FormulaR1C1 = "=ROUNDUP(RC[1]*" & valMH & ",2)"
        newRng.Offset(i + 2, 4).NumberFormat = "#,##"
        newRng.Offset(i + 2, 4).EntireColumn.AutoFit
        newRng.Offset(i + 2, 5).Value = Format(valArea, "#,##0")
        newRng.Offset(i + 2, 5).EntireColumn.AutoFit
        
        newRng.Offset(i + 2, 6).FormulaR1C1 = "=ROUNDUP(" & "RC[-4]/(1-" & (pVal / 100) & "),0)"
        newRng.Offset(i + 2, 6).NumberFormat = "#,##"
        newRng.Offset(i + 2, 6).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).FormulaR1C1 = "=(RC[-1]*RC[-2])" 'Relative Ref.
        newRng.Offset(i + 2, 7).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).NumberFormat = "#,##0"
        priceVal = priceVal + newRng.Offset(i + 2, 6).Value
        newRng.Offset(i + 2, 8).FormulaR1C1 = "=(RC[-2]-RC[-6])/RC[-2]" 'Relative Ref. for %age
        newRng.Offset(i + 2, 8).EntireColumn.AutoFit
        newRng.Offset(i + 2, 8).Style = "Percent"
        
        
        newRng.Offset(i + 2, 9).FormulaR1C1 = "=ROUNDUP(RC[-4]*RC[-8],2)"
        newRng.Offset(i + 2, 9).EntireColumn.AutoFit
        newRng.Offset(i + 2, 9).NumberFormat = "#,##"
        newRng.Offset(i + 2, 10).FormulaR1C1 = "=ROUNDUP(RC[-5]*" & transVal & ",2)"
        newRng.Offset(i + 2, 10).EntireColumn.AutoFit
        newRng.Offset(i + 2, 10).NumberFormat = "#,##"
        newRng.Offset(i + 2, 11).FormulaR1C1 = "=ROUNDUP(RC[-6]*" & toolsVal & ",2)"
        newRng.Offset(i + 2, 11).EntireColumn.AutoFit
        newRng.Offset(i + 2, 11).NumberFormat = "#,##"
        newRng.Offset(i + 2, 12).FormulaR1C1 = "=ROUNDUP(RC[-7]*" & consVal & ",2)"
        newRng.Offset(i + 2, 12).EntireColumn.AutoFit
        newRng.Offset(i + 2, 12).NumberFormat = "#,##"
End Sub



Sub DisplayResult_Injection(i As Integer, ws As Worksheet, valTotal As Variant, _
    valMat As Variant, newRng As Range, valMH As Variant, _
    valTotalCost As Variant, valArea As Variant, pVal As Integer, priceVal As Variant, _
    consVal As Variant, transVal As Variant, toolsVal As Variant, rateVal As Variant)
        '------------DISPLAY SETTING FOR VALUES----------------
        
        valMH = Application.WorksheetFunction.RoundUp(valMH, 2)
        consVal = Application.WorksheetFunction.RoundUp(consVal, 2)
        toolsVal = Application.WorksheetFunction.RoundUp(toolsVal, 2)
        transVal = Application.WorksheetFunction.RoundUp(transVal, 2)
        
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Interior.Color = vbBlack
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Font.Color = vbWhite
        newRng.Offset(2, 0).Value = "System Name"
        newRng.Offset(2, 0).Font.Bold = True
        newRng.Offset(2, 1).Value = "Mat Cost"
        newRng.Offset(2, 1).Font.Bold = True
        newRng.Offset(2, 2).Value = "Unit Cost"
        newRng.Offset(2, 2).Font.Bold = True
        
        newRng.Offset(2, 3).Value = "Total Cost"
        newRng.Offset(2, 3).Font.Bold = True
        newRng.Offset(2, 4).Value = "Manhours"
        newRng.Offset(2, 4).Font.Bold = True
        newRng.Offset(2, 5).Value = "Total QTY"
        newRng.Offset(2, 5).Font.Bold = True
        
        newRng.Offset(2, 6).Value = "Price @ " & pVal & "%"
        newRng.Offset(2, 6).Font.Bold = True
        newRng.Offset(2, 7).Value = "Total Price"
        newRng.Offset(2, 7).Font.Bold = True
        newRng.Offset(2, 8).Value = "%age"
        newRng.Offset(2, 8).Font.Bold = True
        
        newRng.Offset(2, 9).Value = "Mat."
        newRng.Offset(2, 9).Font.Bold = True
        newRng.Offset(2, 10).Value = "Trans."
        newRng.Offset(2, 10).Font.Bold = True
        newRng.Offset(2, 11).Value = "T & E"
        newRng.Offset(2, 11).Font.Bold = True
        newRng.Offset(2, 12).Value = "Cons."
        newRng.Offset(2, 12).Font.Bold = True
        
        '----------ACTUAL VALUE CALCULATIONS------------
        newRng.Offset(i + 2, 0).Value = ws.Name
        newRng.Offset(i + 2, 0).EntireColumn.AutoFit
        newRng.Offset(i + 2, 1).FormulaR1C1 = "=ROUNDUP(RC[8]/RC[4],2)"
        newRng.Offset(i + 2, 1).EntireColumn.AutoFit
        newRng.Offset(i + 2, 1).NumberFormat = "#,##.00"
        newRng.Offset(i + 2, 2).FormulaR1C1 = "=ROUNDUP(" & "(SUM(RC[7]:RC[10])+(RC[2]*" & rateVal & "))/(RC[3]),2)"
        newRng.Offset(i + 2, 2).EntireColumn.AutoFit
        newRng.Offset(i + 2, 2).NumberFormat = "#,##.00"
        
        newRng.Offset(i + 2, 3).FormulaR1C1 = "=ROUNDUP(RC[2]*RC[-1],2)"
        newRng.Offset(i + 2, 3).NumberFormat = "#,##"
        newRng.Offset(i + 2, 3).EntireColumn.AutoFit
        newRng.Offset(i + 2, 4).Value = valMH
        newRng.Offset(i + 2, 4).NumberFormat = "#,##"
        newRng.Offset(i + 2, 4).EntireColumn.AutoFit
        newRng.Offset(i + 2, 5).Value = Format(valArea, "#,##0")
        newRng.Offset(i + 2, 5).EntireColumn.AutoFit
        
        newRng.Offset(i + 2, 6).FormulaR1C1 = "=ROUNDUP(" & "RC[-4]/(1-" & (pVal / 100) & "),0)"
        newRng.Offset(i + 2, 6).NumberFormat = "#,##"
        newRng.Offset(i + 2, 6).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).FormulaR1C1 = "=(RC[-1]*RC[-2])" 'Relative Ref.
        newRng.Offset(i + 2, 7).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).NumberFormat = "#,##0"
        priceVal = priceVal + newRng.Offset(i + 2, 6).Value
        newRng.Offset(i + 2, 8).FormulaR1C1 = "=(RC[-2]-RC[-6])/RC[-2]" 'Relative Ref. for %age
        newRng.Offset(i + 2, 8).EntireColumn.AutoFit
        newRng.Offset(i + 2, 8).Style = "Percent"
        
        
        newRng.Offset(i + 2, 9).Value = valMat
        newRng.Offset(i + 2, 9).EntireColumn.AutoFit
        newRng.Offset(i + 2, 9).NumberFormat = "#,##"
        newRng.Offset(i + 2, 10).Value = transVal
        newRng.Offset(i + 2, 10).EntireColumn.AutoFit
        newRng.Offset(i + 2, 10).NumberFormat = "#,##"
        newRng.Offset(i + 2, 11).Value = toolsVal
        newRng.Offset(i + 2, 11).EntireColumn.AutoFit
        newRng.Offset(i + 2, 11).NumberFormat = "#,##"
        newRng.Offset(i + 2, 12).Value = consVal
        newRng.Offset(i + 2, 12).EntireColumn.AutoFit
        newRng.Offset(i + 2, 12).NumberFormat = "#,##"

End Sub

'----------------DISPLAY ROUTINE FOR CIVIL WORKS SUMMARY-------------
Sub DisplayResult_Civil(i As Integer, ws As Worksheet, valTotal As Variant, _
    valMat As Variant, newRng As Range, valMH As Variant, _
    AreaVal As Variant, pVal As Integer, priceVal As Variant, _
    consVal As Variant, transVal As Variant, toolsVal As Variant, valQtn As Variant, _
    valDate As Variant, valClient As Variant, counter As Integer, valSys As Variant)
        '------------DISPLAY SETTING FOR VALUES----------------
        
        valMH = Application.WorksheetFunction.RoundUp(valMH, 2)
        consVal = Application.WorksheetFunction.RoundUp(consVal, 2)
        toolsVal = Application.WorksheetFunction.RoundUp(toolsVal, 2)
        transVal = Application.WorksheetFunction.RoundUp(transVal, 2)
        
        If counter < 2 Then
            Range(newRng.Offset(1, 0), newRng.Offset(1, 9)).Font.Color = vbBlue
            newRng.Offset(1, 0).Value = "QTN # " & valQtn
            newRng.Offset(1, 2).Value = "DATED: "
            newRng.Offset(1, 3).Value = valDate
            newRng.Offset(1, 3).Value = Format(valDate, "dd-mm-yy")
            newRng.Offset(1, 6).Value = "CLIENT: "
            newRng.Offset(1, 7).Value = valClient
        Else
        End If
        
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Interior.Color = vbBlack
        Range(newRng.Offset(2, 0), newRng.Offset(2, 12)).Font.Color = vbWhite
        newRng.Offset(2, 0).Value = "System Name"
        newRng.Offset(2, 0).Font.Bold = True
        newRng.Offset(2, 1).Value = "Mat Cost"
        newRng.Offset(2, 1).Font.Bold = True
        newRng.Offset(2, 2).Value = "Unit Cost"
        newRng.Offset(2, 2).Font.Bold = True
        
        newRng.Offset(2, 3).Value = "Total Cost"
        newRng.Offset(2, 3).Font.Bold = True
        newRng.Offset(2, 4).Value = "Manhours"
        newRng.Offset(2, 4).Font.Bold = True
        newRng.Offset(2, 5).Value = "Total QTY"
        newRng.Offset(2, 5).Font.Bold = True
        
        newRng.Offset(2, 6).Value = "Price at " & pVal & "%"
        newRng.Offset(2, 6).Font.Bold = True
        newRng.Offset(2, 7).Value = "Total Price"
        newRng.Offset(2, 7).Font.Bold = True
        newRng.Offset(2, 8).Value = "%age"
        newRng.Offset(2, 8).Font.Bold = True
        
        newRng.Offset(2, 9).Value = "Mat."
        newRng.Offset(2, 9).Font.Bold = True
        newRng.Offset(2, 10).Value = "Trans."
        newRng.Offset(2, 10).Font.Bold = True
        newRng.Offset(2, 11).Value = "T & E"
        newRng.Offset(2, 11).Font.Bold = True
        newRng.Offset(2, 12).Value = "Cons."
        newRng.Offset(2, 12).Font.Bold = True
        
        '----------ACTUAL VALUE CALCULATIONS------------
        newRng.Offset(i + 2, 0).Value = valSys
        newRng.Offset(i + 2, 0).EntireColumn.AutoFit
        newRng.Offset(i + 2, 1).Value = Format(valMat, "###.00")
        newRng.Offset(i + 2, 1).EntireColumn.AutoFit
        newRng.Offset(i + 2, 2).Value = Format(valTotal, "###.00")
        newRng.Offset(i + 2, 2).EntireColumn.AutoFit
        
        newRng.Offset(i + 2, 3).FormulaR1C1 = "=ROUNDUP(RC[2]*RC[-1],2)"
        newRng.Offset(i + 2, 3).NumberFormat = "#,##"
        newRng.Offset(i + 2, 3).EntireColumn.AutoFit
        newRng.Offset(i + 2, 4).FormulaR1C1 = "=ROUNDUP(RC[1]*" & valMH & ",2)"
        newRng.Offset(i + 2, 4).NumberFormat = "#,##"
        newRng.Offset(i + 2, 4).EntireColumn.AutoFit
        newRng.Offset(i + 2, 5).Value = Format(AreaVal, "#,##0")
        newRng.Offset(i + 2, 5).EntireColumn.AutoFit
        
        newRng.Offset(i + 2, 6).FormulaR1C1 = "=ROUNDUP(" & "RC[-4]/(1-" & (pVal / 100) & "),0)"
        newRng.Offset(i + 2, 6).NumberFormat = "#,##"
        newRng.Offset(i + 2, 6).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).FormulaR1C1 = "=(RC[-1]*RC[-2])" 'Relative Ref.
        newRng.Offset(i + 2, 7).EntireColumn.AutoFit
        newRng.Offset(i + 2, 7).NumberFormat = "#,##0"
        priceVal = priceVal + newRng.Offset(i + 2, 6).Value
        newRng.Offset(i + 2, 8).FormulaR1C1 = "=(RC[-2]-RC[-6])/RC[-2]" 'Relative Ref. for %age
        newRng.Offset(i + 2, 8).EntireColumn.AutoFit
        newRng.Offset(i + 2, 8).Style = "Percent"
        
        
        newRng.Offset(i + 2, 9).FormulaR1C1 = "=ROUNDUP(RC[-4]*RC[-8],2)"
        newRng.Offset(i + 2, 9).EntireColumn.AutoFit
        newRng.Offset(i + 2, 9).NumberFormat = "#,##"
        newRng.Offset(i + 2, 10).FormulaR1C1 = "=ROUNDUP(RC[-5]*" & transVal & ",2)"
        newRng.Offset(i + 2, 10).EntireColumn.AutoFit
        newRng.Offset(i + 2, 10).NumberFormat = "#,##"
        newRng.Offset(i + 2, 11).FormulaR1C1 = "=ROUNDUP(RC[-6]*" & toolsVal & ",2)"
        newRng.Offset(i + 2, 11).EntireColumn.AutoFit
        newRng.Offset(i + 2, 11).NumberFormat = "#,##"
        newRng.Offset(i + 2, 12).FormulaR1C1 = "=ROUNDUP(RC[-7]*" & consVal & ",2)"
        newRng.Offset(i + 2, 12).EntireColumn.AutoFit
        newRng.Offset(i + 2, 12).NumberFormat = "#,##"
End Sub

