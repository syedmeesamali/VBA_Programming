Option Explicit

Sub Summarize_Systems(wb As Workbook)

Dim ws As Worksheet
Dim locDesc1 As Range, locTotal1 As Range, Count As Integer, newRng As Range
Dim i As Integer, rngSheet As Range, rCell As Variant, m As Integer

Dim startRow As Integer, endRow As Integer

Count = 0
Set newRng = wb.Sheets("System_Summary").Range("a65536").End(xlUp).Offset(1, 0)

Application.ScreenUpdating = False
On Error Resume Next

i = 0
For Each ws In wb.Worksheets
        With ws
        ws.Unprotect
        If InStr(1, .Range("A1").Value, "DicoTech") > 0 Then
         Set locDesc1 = .Range("B9:B30").Find(What:="Description", LookIn:=xlValues, LookAt:= _
             xlPart, SearchOrder:=xlByRows)
         Set locTotal1 = .Range("B9:B30").Find(What:="Total", LookIn:=xlValues, LookAt:= _
             xlPart, SearchOrder:=xlByRows)
                startRow = locDesc1.Row
                endRow = locTotal1.Row
        
          Set rngSheet = .Range("B11" & ":" & "B" & endRow - 1)
          For Each rCell In rngSheet
             If Not IsEmpty(rCell.Value) Then
                i = i + 1
                newRng.Offset(i).Value = .Range("B3").Value
                newRng.Offset(i, 1).Value = rCell.Offset(0, -1).Value & ". " & rCell.Value
                If Not InStr(1, rCell.Value, "SURFACE") > 0 Then
                    newRng.Offset(i, 2).Value = FormatNumber(Application.WorksheetFunction.Round(rCell.Offset(0, 3).Value, 2)) & " " & rCell.Offset(0, 1).Value & "/" & .Range("C5").Value
                    newRng.Offset(i, 3).Value = FormatNumber(Application.WorksheetFunction.Round(rCell.Offset(0, 2).Value, 2)) & " QAR"
                    newRng.Offset(i, 4).Value = FormatNumber(Application.WorksheetFunction.Round(rCell.Offset(0, 3).Value, 2)) * .Range("B5").Value
                    newRng.Offset(i, 5).Value = rCell.Offset(0, 1).Value
                    
                    newRng.Offset(i).EntireColumn.AutoFit
                    newRng.Offset(i, 1).EntireColumn.AutoFit
                    newRng.Offset(i, 2).EntireColumn.AutoFit
                    newRng.Offset(i, 3).EntireColumn.AutoFit
                    newRng.Offset(i, 4).EntireColumn.AutoFit
                    newRng.Offset(i, 5).EntireColumn.AutoFit
            
                Else
                End If
            Else
            End If
            Next rCell
         
         Else
         End If
         
         End With
Next

newRng.CurrentRegion.Borders.LineStyle = xlContinuous
newRng.CurrentRegion.BorderAround = True

Sheets("System_Summary").Activate

End Sub
