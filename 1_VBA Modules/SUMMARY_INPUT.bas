Option Explicit

Sub INPUT_SUMMARY(wb As Workbook)

Dim ws As Worksheet, newRng As Range, Count As Integer
Dim valMat As Variant, valDia As Variant, valSys As Variant
Dim valArea As Variant, rCell As Range
Count = 0
Application.ScreenUpdating = False
Set newRng = wb.Sheets("Totals").Range("a65536").End(xlUp)
On Error Resume Next
For Each ws In wb.Worksheets
        With ws
            ws.Unprotect
            If InStr(1, .Range("A1").Value, "DicoTech") > 0 Then
                valArea = .Range("B5").Value
                Count = Count + 1
                For Each rCell In .Range("B11:B25")
                  If InStr(1, rCell.Value, "Total") > 0 Then
                    valMat = rCell.Offset(0, 6).Value
                    valDia = .Range("E3").Value
                    valSys = .Range("B3").Value
                  Else
                  End If
                Next rCell
         'CALL THE DISPLAY ROUTINE.
         Call DisplayResult_Input(Count, ws, valMat, newRng, valArea, _
         valDia, valSys)
         Else
         End If
    End With
Next
         newRng.Offset(Count + 1, 1).CurrentRegion.Borders.LineStyle = xlContinuous
         newRng.Offset(Count + 1, 1).CurrentRegion.BorderAround = True
         
Application.ScreenUpdating = True
Sheets("Totals").Activate
End Sub

Sub DisplayResult_Input(i As Integer, ws As Worksheet, valMat As Variant, _
    newRng As Range, valArea As Variant, valDia As Variant, valSys As Variant)
        
        Range(newRng.Offset(1, 0), newRng.Offset(1, 12)).Interior.Color = vbBlack
        Range(newRng.Offset(1, 0), newRng.Offset(1, 12)).Font.Color = vbWhite
        newRng.Offset(1, 0).Value = "System Name"
        newRng.Offset(1, 0).Font.Bold = True
        newRng.Offset(1, 1).Value = "Mat Cost"
        newRng.Offset(1, 1).Font.Bold = True
        newRng.Offset(1, 2).Value = "Area"
        newRng.Offset(1, 2).Font.Bold = True
        
        newRng.Offset(1, 3).Value = "Total Mat"
        newRng.Offset(1, 3).Font.Bold = True
        newRng.Offset(1, 4).Value = "Dia. (in)"
        newRng.Offset(1, 4).Font.Bold = True
        newRng.Offset(1, 5).Value = "Dia. (mm)"
        newRng.Offset(1, 5).Font.Bold = True
        
        newRng.Offset(1, 6).Value = "Surface Prep."
        newRng.Offset(1, 6).Font.Bold = True
        newRng.Offset(1, 7).Value = "1st Coat"
        newRng.Offset(1, 7).Font.Bold = True
        newRng.Offset(1, 8).Value = "2nd Coat"
        newRng.Offset(1, 8).Font.Bold = True
        
        newRng.Offset(1, 9).Value = "3rd Coat"
        newRng.Offset(1, 9).Font.Bold = True
        newRng.Offset(1, 10).Value = "Cons"
        newRng.Offset(1, 10).Font.Bold = True
        newRng.Offset(1, 11).Value = "T & E"
        newRng.Offset(1, 11).Font.Bold = True
        newRng.Offset(1, 12).Value = "Special"
        newRng.Offset(1, 12).Font.Bold = True
        
        '----------ACTUAL VALUE CALCULATIONS------------
        newRng.Offset(i + 1, 0).Value = valSys
        newRng.Offset(i + 1, 0).EntireColumn.AutoFit
        newRng.Offset(i + 1, 1).Value = Format(valMat, "###.00")
        newRng.Offset(i + 1, 1).EntireColumn.AutoFit
        newRng.Offset(i + 1, 2).Value = Format(valArea, "#,##")
        newRng.Offset(i + 1, 2).EntireColumn.AutoFit
        
        newRng.Offset(i + 1, 3).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
        newRng.Offset(i + 1, 3).NumberFormat = "#,##"
        newRng.Offset(i + 1, 3).EntireColumn.AutoFit
        newRng.Offset(i + 1, 4).Value = valDia
        newRng.Offset(i + 1, 4).EntireColumn.AutoFit
        newRng.Offset(i + 1, 5).FormulaR1C1 = "=ROUNDUP(RC[-1]*25.4,2)"
        newRng.Offset(i + 1, 5).NumberFormat = "#,##.00"
        newRng.Offset(i + 1, 5).EntireColumn.AutoFit
        
End Sub

