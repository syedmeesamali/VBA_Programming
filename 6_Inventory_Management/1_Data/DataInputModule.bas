Attribute VB_Name = "DataInputModule"
'Program by "Syed Meesam Ali Naqvi"
'Version 1.1.0
'Latest production date: 08/01/2020
'Client: Mr. Sibte Ali

Option Explicit

'Module to take input data from various invoice files stored in a folder (currently called as original)

Sub Data_NewInput()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim objFs As Object, objFolder As Object
Dim file As Object

Set objFs = CreateObject("Scripting.FileSystemObject") 'Scripting system helps us pick files using windows folder picker

On Error Resume Next 'Important line to do error handling
Set objFolder = objFs.GetFolder(FolderPick())
Dim newRng As Range, rCell As Range, countRng As Range
Dim strtRow As Integer, totalRows As Integer, totalCols As Integer, refRow As Integer


Dim wb As Workbook, Ws As Worksheet, count As Integer, dateCell As Range, custName As String
Set wb = Application.ActiveWorkbook
wb.Sheets("StockOut").Unprotect Password:="ali"

'----------ONE LINE WON'T BE BIG ISSUE----------
Set newRng = wb.Sheets("StockOut").Range("A150000").End(xlUp).Offset(0, 0)
Dim fileRange As Range, fileCount As Integer
Set fileRange = wb.Sheets("Help").Range("M65000").End(xlUp).Offset(0, 0)

count = 0               'Counter to be used for rows in the input files

fileCount = 0           'Variable to keep track of number of excel files in the input folder & check duplication


On Error Resume Next
For Each file In objFolder.Files
    Dim src As Workbook
    
    
    If Not IsEmpty(fileRange.Offset(fileCount, 0)) Then
    '----CHECK IF THE FILE IS ALREADY IMPORTED
        If InStr(1, wb.Sheets("Help").Range("M1:" & fileRange.Offset(fileCount, 0)), file, vbBinaryCompare) Then
            MsgBox "Some file or files already imported! NOT ALLOWED TO IMPORT AGAIN !"
            Exit Sub
        Else
            fileRange.Offset(fileCount, 0) = file
            fileCount = fileCount + 1
            Set src = Workbooks.Open(file.Path, True, True)
            refRow = src.Worksheets("rptCustBill").Range("A65536").End(xlUp).Row    'Reference row to pick data
            Set dateCell = src.Worksheets("rptCustBill").Range("O4")                'This is to pick date from invoice sheets
            custName = src.Worksheets("rptCustBill").Range("H8")                    'Customer name
            Set countRng = src.Worksheets("rptCustBill").Range("A15:A" & refRow)    'Define range in the target sheet
            On Error Resume Next                                                    'Continue if some error encountered in some file
            For Each rCell In countRng
                newRng.Offset(count, 0) = Application.Max(wb.Sheets("StockOut").Range("A:A")) + 1   'Serial Number
                newRng.Offset(count, 1) = Format(Right(dateCell.Value, Len(dateCell.Value) - 6), "dd/mm/yyyy")       'Date
                newRng.Offset(count, 2) = "ABC123"                                              'Customer Code Goes here
                newRng.Offset(count, 3) = custName                                              'Customer name
                newRng.Offset(count, 4) = rCell.Value                                           'Code
                newRng.Offset(count, 5) = rCell.Offset(0, 1).Value                              'Product name
                'newRng.Offset(count, 4) = rCell.Offset(0, 7).Value                             'Pcs - if required
                newRng.Offset(count, 8) = rCell.Offset(0, 4).Value                              'Boxes
                newRng.Offset(count, 9) = rCell.Offset(0, 9).Value * rCell.Offset(0, 7).Value   'Total Price
                count = count + 1
            Next
            src.Close False
            Set src = Nothing
        End If  'End of file check related IF
     Else
            fileRange.Offset(fileCount, 0) = file
            fileCount = fileCount + 1
            Set src = Workbooks.Open(file.Path, True, True)
            refRow = src.Worksheets("rptCustBill").Range("A65536").End(xlUp).Row    'Reference row to pick data
            Set dateCell = src.Worksheets("rptCustBill").Range("O4")                'This is to pick date from invoice sheets
            custName = src.Worksheets("rptCustBill").Range("H8")                    'Customer name
            Set countRng = src.Worksheets("rptCustBill").Range("A15:A" & refRow)    'Define range in the target sheet
            On Error Resume Next                                                    'Continue if some error encountered in some file
            For Each rCell In countRng
                newRng.Offset(count, 0) = Application.Max(wb.Sheets("StockOut").Range("A:A")) + 1   'Serial Number
                newRng.Offset(count, 1) = Format(Right(dateCell.Value, Len(dateCell.Value) - 6), "dd/mm/yyyy")        'Date
                newRng.Offset(count, 2) = "ABC123"                                              'Customer Code Goes here
                newRng.Offset(count, 3) = custName                                              'Customer name
                newRng.Offset(count, 4) = rCell.Value                                           'Code
                newRng.Offset(count, 5) = rCell.Offset(0, 1).Value                              'Product name
                'newRng.Offset(count, 4) = rCell.Offset(0, 7).Value                             'Pcs - if required
                newRng.Offset(count, 8) = rCell.Offset(0, 4).Value                              'Boxes
                newRng.Offset(count, 9) = rCell.Offset(0, 9).Value * rCell.Offset(0, 7).Value   'Total Price
                count = count + 1
            Next
            src.Close False
            Set src = Nothing
     
     End If
Next

wb.Sheets("StockOut").Protect Password:="ali"
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


Sub Purchase_Input()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim objFs As Object, objFolder As Object
Dim file As Object

Set objFs = CreateObject("Scripting.FileSystemObject") 'Scripting system helps us pick files using windows folder picker

On Error Resume Next 'Important line to do error handling
Set objFolder = objFs.GetFolder(FolderPick())
Dim newRng As Range, rCell As Range, countRng As Range
Dim strtRow As Integer, totalRows As Integer, totalCols As Integer, refRow As Integer

Dim wb As Workbook, Ws As Worksheet, count As Integer, dateCell As Range, supName As String
Set wb = Application.ActiveWorkbook
wb.Sheets("StockIn").Unprotect Password:="ali"

'---------DEFINE RANGES IN THE TARGET SHEET----------

Set newRng = wb.Sheets("StockIn").Range("A150000").End(xlUp).Offset(0, 0)
Dim fileRange As Range, fileCount As Integer
Set fileRange = wb.Sheets("Help").Range("N65000").End(xlUp).Offset(0, 0)
count = 0
fileCount = 0           'Variable to keep tr

On Error Resume Next
For Each file In objFolder.Files
    Dim src As Workbook
    
    '----CHECK IF THE FILE IS ALREADY IMPORTED
    If Not IsEmpty(fileRange.Offset(fileCount, 0)) Then
        If InStr(1, wb.Sheets("Help").Range("N1:" & fileRange.Offset(fileCount, 0)), file, vbBinaryCompare) Then
            MsgBox "Some file or files already imported! NOT ALLOWED TO IMPORT AGAIN !"
            Exit Sub
        Else
            fileRange.Offset(fileCount, 0) = file
            fileCount = fileCount + 1
            Set src = Workbooks.Open(file.Path, True, True)
            refRow = src.Worksheets("MasterSheet").Range("A65536").End(xlUp).Row    'Reference row to pick data
            Set countRng = src.Worksheets("MasterSheet").Range("A2:A" & refRow)    'Define range in the target sheet
            On Error Resume Next                                                    'Continue if some error encountered in some file
            For Each rCell In countRng
                newRng.Offset(count, 4) = rCell.Value                                           'Product ID
                newRng.Offset(count, 0) = Application.Max(wb.Sheets("StockIn").Range("A:A")) + 1   'Serial Number
                newRng.Offset(count, 5) = rCell.Offset(0, 1).Value                              'Product Name
                newRng.Offset(count, 1) = Format(rCell.Offset(0, 2).Value, "dd/mm/yyyy")                             'Date
                newRng.Offset(count, 8) = rCell.Offset(0, 3).Value                              'Expiry
                newRng.Offset(count, 2) = rCell.Offset(0, 4).Value                              'Supplier ID
                newRng.Offset(count, 3) = rCell.Offset(0, 5).Value                              'Supplier Name
                newRng.Offset(count, 9) = rCell.Offset(0, 6).Value                              'Qty
                newRng.Offset(count, 10) = rCell.Offset(0, 7).Value                             'Cost
                count = count + 1
            Next
            src.Close False
            Set src = Nothing
        End If
    Else
            fileRange.Offset(fileCount, 0) = file
            fileCount = fileCount + 1
            Set src = Workbooks.Open(file.Path, True, True)
            refRow = src.Worksheets("MasterSheet").Range("A65536").End(xlUp).Row    'Reference row to pick data
            Set countRng = src.Worksheets("MasterSheet").Range("A2:A" & refRow)    'Define range in the target sheet
            On Error Resume Next                                                    'Continue if some error encountered in some file
            For Each rCell In countRng
                newRng.Offset(count, 4) = rCell.Value                                           'Product ID
                newRng.Offset(count, 0) = Application.Max(wb.Sheets("StockIn").Range("A:A")) + 1   'Serial Number
                newRng.Offset(count, 5) = rCell.Offset(0, 1).Value                              'Product Name
                newRng.Offset(count, 1) = Format(rCell.Offset(0, 2).Value, "dd/mm/yyyy")                             'Date
                newRng.Offset(count, 8) = rCell.Offset(0, 3).Value                              'Expiry
                newRng.Offset(count, 2) = rCell.Offset(0, 4).Value                              'Supplier ID
                newRng.Offset(count, 3) = rCell.Offset(0, 5).Value                              'Supplier Name
                newRng.Offset(count, 9) = rCell.Offset(0, 6).Value                              'Qty
                newRng.Offset(count, 10) = rCell.Offset(0, 7).Value                             'Cost
                count = count + 1
            Next
            src.Close False
            Set src = Nothing
    End If
Next
wb.Sheets("StockIn").Protect Password:="ali"          'Protect the sheet again
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub


'------------------------------------------------------------------
'Below function is for picking a folder for input using dialog box
'------------------------------------------------------------------
Function FolderPick() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker) 'Pick folder using a dialog box
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    FolderPick = sItem
    Set fldr = Nothing
End Function
