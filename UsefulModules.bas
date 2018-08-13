Public Declare Function SendCDcmd Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Dim lRet As Long

' Below function spells a number in plain-english format. e.g. input of 57,000 will produce
'"Fifty Seven Thousand only" 
      
      Function SpellNumber(ByVal MyNumber)
          Dim Dollars, Cents, Temp
          Dim DecimalPlace, Count
          ReDim Place(9) As String
          
          Place(2) = " Thousand "
          Place(3) = " Million "
          Place(4) = " Billion "
          Place(5) = " Trillion "

          ' String representation of amount.
          MyNumber = Trim(Str(MyNumber))
          
          ' Position of decimal place 0 if none.
          DecimalPlace = InStr(MyNumber, ".")
          
          ' Convert cents and set MyNumber to dollar amount.
          If DecimalPlace > 0 Then
              Cents = (Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
              MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
          End If
          
          Count = 1
          Do While MyNumber <> ""
              Temp = GetHundreds(Right(MyNumber, 3))
              If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
              If Len(MyNumber) > 3 Then
                  MyNumber = Left(MyNumber, Len(MyNumber) - 3)
              Else
                  MyNumber = ""
              End If
              Count = Count + 1
          Loop

          Select Case Dollars
              Case ""
                  Dollars = "Empty "
              Case "One"
                  Dollars = "One "
              Case Else
                  Dollars = Dollars & ""
          End Select

          Select Case Cents
              Case ""
                  Cents = " Only"
              Case "One"
                  Cents = " And " & " 1 / 100 Only"
              Case Else
                  Cents = " And " & Cents & " / 100 Only"
          End Select
          SpellNumber = Dollars & Cents
      End Function

      '*******************************************
      ' Converts a number from 100-999 into text *
      '*******************************************

      Function GetHundreds(ByVal MyNumber)

          Dim Result As String
          If val(MyNumber) = 0 Then Exit Function
          MyNumber = Right("000" & MyNumber, 3)
          ' Convert the hundreds place.
          If Mid(MyNumber, 1, 1) <> "0" Then
              Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
          End If

         ' Convert the tens and ones place.
          If Mid(MyNumber, 2, 1) <> "0" Then
              Result = Result & GetTens(Mid(MyNumber, 2))
          Else
              Result = Result & GetDigit(Mid(MyNumber, 3))
          End If
          GetHundreds = Result
      End Function
     
     '*********************************************
      ' Converts a number from 10 to 99 into text. *
      '*********************************************
     Function GetTens(TensText)
          Dim Result As String
          Result = ""           ' Null out the temporary function value.
          If val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...

              Select Case val(TensText)
              
                  Case 10: Result = "Ten"
                  Case 11: Result = "Eleven"
                  Case 12: Result = "Twelve"
                  Case 13: Result = "Thirteen"
                  Case 14: Result = "Fourteen"
                  Case 15: Result = "Fifteen"
                  Case 16: Result = "Sixteen"
                  Case 17: Result = "Seventeen"
                  Case 18: Result = "Eighteen"
                  Case 19: Result = "Nineteen"
                  Case Else

              End Select

          Else                                 ' If value between 20-99...
              Select Case val(Left(TensText, 1))
                  Case 2: Result = "Twenty "
                  Case 3: Result = "Thirty "
                  Case 4: Result = "Forty "
                  Case 5: Result = "Fifty "
                  Case 6: Result = "Sixty "
                  Case 7: Result = "Seventy "
                  Case 8: Result = "Eighty "
                  Case 9: Result = "Ninety "
                  Case Else
              End Select
              Result = Result & GetDigit _
                  (Right(TensText, 1))  ' Retrieve ones place.
          End If
          GetTens = Result
      End Function

      '*******************************************
      ' Converts a number from 1 to 9 into text. *
      '*******************************************

      Function GetDigit(Digit)
          Select Case val(Digit)
              Case 1: GetDigit = "One"
              Case 2: GetDigit = "Two"
              Case 3: GetDigit = "Three"
              Case 4: GetDigit = "Four"
              Case 5: GetDigit = "Five"
              Case 6: GetDigit = "Six"
              Case 7: GetDigit = "Seven"
              Case 8: GetDigit = "Eight"
              Case 9: GetDigit = "Nine"
              Case Else: GetDigit = ""
          End Select
      End Function


Function GetLastUsedColumn(rg As Range) As Long
Dim lMaxColumns As Long
lMaxColumns = ThisWorkbook.Worksheets(1).Columns.Count 'Total Column count

If IsEmpty(rg.Parent.Cells(rg.Row, lMaxColumns)) Then
    GetLastUsedColumn = rg.Parent.Cells(rg.Row, lMaxColumns).End(xlToLeft).Column
    Else
        GetLastUsedColumn = rg.Parent.Cells(rg.Row, lMaxColumns).Column
    End If
End Function


'Rho-value (concrete design) function
Function Rho_Value(Fc As Integer, Fy As Integer, b As Integer, d As Integer, Mu As Integer) As Single

Dim Rn As Single
Rn = (Mu * 12000) / (b * d ^ 2)
Rho_Value = ((0.85 * Fc) / Fy) * (1 - Sqr(1 - ((2 * Rn) / (0.85 * Fc))))

End Function

'Function to calculate commission based on sales amount.
Function Commission(Sales)
    Const tier1 = 0.08
    Const tier2 = 0.12
    Const tier3 = 0.14
    Const tier4 = 0.15

    Select Case Sales
    Case 0 To 10000: Commission = Sales * tier1
    Case 10001 To 20000: Commission = Sales * tier2
    Case 20001 To 30000: Commission = Sales * tier3
    Case Is > 30000: Commission = Sales * tier4
    End Select

End Function

'Round-down the contents of cell 
Sub MakeCellDown()
Dim val As Variant
Dim selRng As Range
Dim rCell As Range
Set selRng = Selection
For Each rCell In selRng.Cells
    If rCell.Value = "" Then
    rCell.Value = rCell.Value
    Else
    val = rCell.Value
    rCell.FormulaR1C1 = "=ROUNDDOWN(" & val & ",0)"
    End If
Next rCell
End Sub

'Round-up the contents of cell 
Sub MakeCellUP()
Dim val As Variant
Dim selRng As Range
Dim rCell As Range

Set selRng = Selection

For Each rCell In selRng.Cells
    If rCell.Value = "" Then
    rCell.Value = rCell.Value
    Else
    val = rCell.Value
    rCell.FormulaR1C1 = "=ROUNDUP(" & val & ",0)"
    End If
Next rCell

End Sub

'Program to guess password of protected sheet using BRUTE-FORCE
Sub GreatProgram()
  'Author unknown but submitted by brettdj of www.experts-exchange.com
  
  Dim i As Integer, j As Integer, k As Integer
  Dim l As Integer, m As Integer, n As Integer
  Dim i1 As Integer, i2 As Integer, i3 As Integer
  Dim i4 As Integer, i5 As Integer, i6 As Integer
  On Error Resume Next
  For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
  For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
  For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
  For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
     
        
 ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
      Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
      Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
  If ActiveSheet.ProtectContents = False Then
      MsgBox "One usable password is " & Chr(i) & Chr(j) & _
          Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
          Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
   ActiveWorkbook.Sheets(1).Select
   Range("a1").FormulaR1C1 = Chr(i) & Chr(j) & _
          Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
          Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
       Exit Sub
  End If
  Next: Next: Next: Next: Next: Next
  Next: Next: Next: Next: Next: Next
End Sub


Sub GetFileNames()
Dim FSO             As Object
Dim fPath           As String
Dim myFolder, myFile
Dim r As Integer
   
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count > 0 Then fPath = .SelectedItems(1) & "\"
    End With
    Set myFolder = FSO.GetFolder(fPath).Files

r = 1
'Pick the file names based on certain extensions.
On Error Resume Next
For Each myFile In myFolder
If LCase(myFile) Like "*.*" Then
    ActiveCell.Offset(r) = myFile
    r = r + 1
End If
Next myFile

End Sub

'Rename tabs of work-book based on some criteria
Sub Sheets_Tabs_Naming()

Dim ws As Worksheet, wb As Workbook
Set wb = ActiveWorkbook
Count = 0
Application.ScreenUpdating = False
On Error Resume Next
For Each ws In wb.Worksheets
        With ws
            ws.Unprotect
            'CHECK FOR COATING TYPE ESTIMATION SHEETS.
            If InStr(1, .Range("A1").Value, "Target") > 0 Then
                ws.Name = .Range("B3").Value
                         
             Else
             End If
    End With
Next
End Sub

'Program to import some forms
Sub ExportImportForm()
    
    Dim SourceWb As Workbook
    Dim DestinationWB As Workbook
    
    Set SourceWb = ThisWorkbook
    Set DestinationWB = Workbooks("Estimations_Latest 17-03-14 new.xls")
    
    SourceWb.VBProject.VBComponents("frmSummary").Export "frmSummary.frm"
    DestinationWB.VBProject.VBComponents.Import "frmSummary.frm"
    'Kill "frmChangeFDOB.frm"
    'Kill "frmChangeFDOB.frx"
    
End Sub
