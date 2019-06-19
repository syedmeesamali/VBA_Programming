Option Explicit
Sub Make_SOA2()

'Declaration of various variables
Dim accntRng As Range, testRng As Range, ws As Worksheet, wsTest As Worksheet, wb As Workbook
Dim cell_count As Integer, count As Integer, varUnique As Variant
Dim resultRng As Range

'We don't want screen updating or alerts during execution
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'wb is this workbook. May be used later for various functions.
Set wb = Application.ThisWorkbook
Set ws = ThisWorkbook.Worksheets("SAP")
Set wsTest = ThisWorkbook.Worksheets("test")

'Total No. of records in accounts range
Dim accntLastRow As Integer
accntLastRow = ws.Range("B65536").End(xlUp).Row
Set accntRng = ws.Range("B2:B" & accntLastRow)
Set testRng = wsTest.Range("A1")

'Below snippet will take unique values from the accounts column and treat it as account holders list
'accntRng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=testRng, Unique:=True
Dim lastRow As Integer
lastRow = wsTest.Range("A65536").End(xlUp).Row
Set resultRng = wsTest.Range("A1:A" & lastRow)
cell_count = resultRng.Cells.count
MsgBox "count is: " & cell_count

wsTest.Activate

'Below CALL will create new sheets based on unique customer names
Call AddSheets

Dim rCell As Range, tCell As Range, flag As Boolean
'MAIN algorithm for data gathering goes below
'////////////////////////////////////////
flag = True

For Each rCell In resultRng
Sheets("SOA_" & rCell.Value).Range("D7").Value = rCell.Value
Sheets("SOA_" & rCell.Value).Range("D8").Value = "Cusomer_" & rCell.Value
Sheets("SOA_" & rCell.Value).Range("D9").Value = "CITY_" & rCell.Value
        
        If (flag = True) Then
            count = 0
            flag = False
        Else
            count = 0
            flag = True
        End If
    
   For Each tCell In accntRng
    If (tCell.Value = rCell.Value) Then
        count = count + 1
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 2) = rCell.Value                      'Populate SAME account number
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 3) = "Cust Name: " & rCell.Value      'Populate Customer's name
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 4) = "Doctor: " & rCell.Value         'Populate Doctor's name
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 5) = "Inv: " & tCell.Offset(0, 5).Value   'Populate Invoice Ref
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 8) = tCell.Offset(0, 7).Value             'Populate Invoice Amount
        Sheets("SOA_" & rCell.Value).Range("A10").Offset(count, 9) = tCell.Offset(0, 3).Value             'Populate Date Value
    End If
   Next
Next

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

'--------------------------------------------
'Below portion is only for addition of sheets
'--------------------------------------------

Sub AddSheets()
    Dim xRg As Range
    Dim wSh As Worksheet, wBk As Workbook
    Set wBk = ActiveWorkbook
    Set wSh = wBk.Worksheets("SOABlank")
    
    Dim rngNames As Range, lastRow As Integer, i As Integer
    
    Application.ScreenUpdating = False
    lastRow = Range("A65536").End(xlUp).Row
    Set rngNames = Range("A1:A" & lastRow)
    Dim range_Cells As Integer
    range_Cells = rngNames.Cells.count
    
    
    For Each xRg In rngNames
        With wBk
            wSh.Activate
            wSh.Copy After:=ActiveWorkbook.Sheets("SOABlank")
            On Error Resume Next
            ActiveSheet.Name = "SOA_" & xRg.Value
            If Err.Number = 1004 Then
              Debug.Print xRg.Value & " already used as a sheet name"
            End If
            On Error GoTo 0
        End With
    Next xRg
    Application.ScreenUpdating = True
End Sub
