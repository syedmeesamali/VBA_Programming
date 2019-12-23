Attribute VB_Name = "Module1"
Public Sub SaveAttachmentsToFolder_Delete()
Dim objApp As Outlook.Application
Dim aMail As Outlook.MailItem 'Object
Dim oAttachments As Outlook.Attachments
Dim oSelection As Outlook.Selection
Dim i As Long
Dim iCount As Long
Dim sFile As String
Dim sFolderPath As String
Dim sDeletedFiles As String
  
    sFolderPath = CreateObject("WScript.Shell").SpecialFolders(16)
    On Error Resume Next
  
    Set objApp = CreateObject("Outlook.Application")
    Set oSelection = objApp.ActiveExplorer.Selection
    sFolderPath = sFolderPath & "\OLAttachments"
    
    For Each aMail In oSelection
  
    Set oAttachments = aMail.Attachments
    iCount = oAttachments.Count
      
        
    If iCount > 0 Then
        For i = iCount To 1 Step -1
            sFile = Format(aMail.SentOn, "YYYY-MM-DD_HHNNSS") & "_" & oAttachments.Item(i).FileName
            sFile = sFolderPath & "\" & sFile
            oAttachments.Item(i).SaveAsFile sFile
            oAttachments.Item(i).Delete
        Next i
            aMail.Save
            sDeletedFiles = ""
     
       End If
    Next
      
ExitSub:
  
Set oAttachments = Nothing
Set aMail = Nothing
Set oSelection = Nothing
Set objApp = Nothing
End Sub

Public Sub SaveAttachmentsToFolder_NoDelete()
Dim objApp As Outlook.Application
Dim aMail As Outlook.MailItem 'Object
Dim oAttachments As Outlook.Attachments
Dim oSelection As Outlook.Selection
Dim i As Long
Dim iCount As Long
Dim sFile As String
Dim sFolderPath As String

  
    sFolderPath = CreateObject("WScript.Shell").SpecialFolders(16)
    On Error Resume Next
  
    Set objApp = CreateObject("Outlook.Application")
    Set oSelection = objApp.ActiveExplorer.Selection
    sFolderPath = sFolderPath & "\OLAttachments"
    
    
    'Save attachments from selected mails
    For Each aMail In oSelection
    Set oAttachments = aMail.Attachments
    iCount = oAttachments.Count
    
    If iCount > 0 Then
        'Processing each attachment for every selected email
        For i = iCount To 1 Step -1
            sFile = Format(aMail.SentOn, "YYYY-MM-DD_HHNNSS") & "_" & oAttachments.Item(i).FileName
            sFile = sFolderPath & "\" & sFile
            oAttachments.Item(i).SaveAsFile sFile
        Next i
       End If
    Next
      
ExitSub:
  
Set oAttachments = Nothing
Set aMail = Nothing
Set oSelection = Nothing
Set objApp = Nothing
End Sub


Public Sub SaveMessageAsMsg()
  Dim oMail As Outlook.MailItem
  Dim objItem As Object
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim enviro As String
  Dim strFolderpath As String
  
    enviro = CStr(Environ("USERPROFILE"))
'Defaults to Documents folder
' get the function at //slipstick.me/u1a2d
strFolderpath = BrowseForFolder(enviro & "\documents\")
   
   For Each objItem In ActiveExplorer.Selection
   If objItem.MessageClass = "IPM.Note" Then
    Set oMail = objItem
    
  sName = oMail.Subject
  ReplaceCharsForFileName sName, "-"
  
  dtDate = oMail.ReceivedTime
  sName = Format(dtDate, "yyyy-mm-dd", vbUseSystemDayOfWeek, _
    vbUseSystem) & Format(dtDate, "_hhnnss", _
    vbUseSystemDayOfWeek, vbUseSystem) & "_" & sName & ".msg"
      
  sPath = strFolderpath & "\"
  MsgBox sPath & sName
  oMail.SaveAs sPath & sName, olMSG
   
  End If
  Next
   
End Sub

Function BrowseForFolder(Optional OpenAt As Variant) As Variant
  Dim ShellApp As Object
  Set ShellApp = CreateObject("Shell.Application"). _
 BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
 
 On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
 On Error GoTo 0
 
 Set ShellApp = Nothing
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":"
            If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\"
            If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else
            GoTo Invalid
    End Select
 Exit Function
 
Invalid:
 BrowseForFolder = False
End Function

Public Sub File_Search_Attach()
Dim objApp As Outlook.Application
Dim aMail As Outlook.MailItem               'Object
Dim myAttachments As Outlook.Attachments
Dim oSelection As Outlook.Selection
Dim i As Long, j As Integer
Dim iCount As Long
Dim sFile As String
Dim mailDate As String
Dim sFolderPath As String
Dim Search_Filter(1 To 9) As String

Search_Filter(1) = "*.jpg"
Search_Filter(2) = "*.pdf"
Search_Filter(3) = "*.rar"
Search_Filter(4) = "*.xls"
Search_Filter(5) = "*.jpeg"
Search_Filter(6) = "*.zip"
Search_Filter(7) = "*.png"
Search_Filter(8) = "*.dwg"
Search_Filter(9) = "*.fdb"

    'The following method is used to search for files in any folder.
    Dim Coll_Docs As New Collection
    Dim Search_path, Search_Fullname As String
    Dim DocName As String
    
    Set objApp = CreateObject("Outlook.Application")
    Set oSelection = objApp.ActiveExplorer.Selection
    
    
    Search_path = "C:\Users\SYED\Documents\OLAttachments"   'Folder path to search for files
    Set Coll_Docs = Nothing
    
    For j = 1 To 9
        DocName = Dir(Search_path & "\" & Search_Filter(j))
            Do Until DocName = ""            ' Build the collection
            Coll_Docs.Add Item:=DocName
        DocName = Dir
        Loop
    Next

    'Total numbers in collection. Can be verified directly by counting files of a types in that folder.
    MsgBox "There were " & Coll_Docs.Count & " file(s) found."
    
    For Each aMail In oSelection
            mailDate = Format(aMail.SentOn, "YYYY-MM-DD_HHNNSS")
            Set myAttachments = aMail.Attachments
        For i = Coll_Docs.Count To 1 Step -1              '
            Search_Fullname = Search_path & "\" & Coll_Docs(i)
                If Left(Coll_Docs(i), 17) = mailDate Then
                    myAttachments.Add Search_Fullname
                    Kill (Search_Fullname)
                Else
                End If
        Next
        aMail.Save
    Next
    MsgBox "Attachments Added !"
ExitSub:
  
Set myAttachments = Nothing
Set aMail = Nothing
Set oSelection = Nothing
Set objApp = Nothing

End Sub


Sub SubjectPrefix()
    Dim olItem As MailItem
    Dim sFilenum As String

    sFilenum = "Structural Report: "

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No Items selected!", vbCritical, "Error"
    End If

    '// Process each selected Mail Item
    For Each olItem In Application.ActiveExplorer.Selection
        olItem.Subject = "[" & sFilenum & "] " & olItem.Subject
        olItem.Save
    Next olItem
End Sub

Sub ProcessFolder(CurrentFolder As Outlook.MAPIFolder)
Const cstrProcedure = "ProcessFolder"
Dim i As Long
Dim olNewFolder As Outlook.MAPIFolder
Dim olTempItem As Object
Dim intColumn As Integer
Dim wb As Workbook
Dim wsEmails As Worksheet
'
2270 On Error GoTo HandleError
2280 Set wb = ThisWorkbook
2290 Set wsEmails = wb.Sheets("Emails")
2300 For i = 1 To CurrentFolder.Items.Count
2310     Set olTempItem = CurrentFolder.Items(i)
2320     If olTempItem <> "" Then
            ' record details
2360             intColumn = 1
2370             On Error Resume Next
2380             wsEmails.Cells(publngAuditRecord, intColumn) = CurrentFolder.Name
2390             intColumn = intColumn + 1
2400             wsEmails.Cells(publngAuditRecord, intColumn) = olTempItem.Subject
2410             intColumn = intColumn + 1
2420             wsEmails.Cells(publngAuditRecord, intColumn) = olTempItem.SenderEmailAddress
2430             intColumn = intColumn + 1
2440             wsEmails.Cells(publngAuditRecord, intColumn) = olTempItem.ReceivedTime
2450             intColumn = intColumn + 1
2460             wsEmails.Cells(publngAuditRecord, intColumn) = olTempItem.Attachments.Count
2470             intColumn = intColumn + 1
2480             wsEmails.Cells(publngAuditRecord, intColumn) = (olTempItem.Size / 1024)
2482             intColumn = intColumn + 1
2484             wsEmails.Cells(publngAuditRecord, intColumn) = olTempItem.Body ' <----- Added
2488             intColumn = intColumn + 1  ' <----- Added
2490             publngAuditRecord = publngAuditRecord + 1
2500             On Error GoTo HandleError
        End If
    Next i
' WARNING: Recursion ...
2550 For Each olNewFolder In CurrentFolder.Folders
2560     ProcessFolder olNewFolder
Next olNewFolder
'
2590 Set olTempItem = Nothing
2600 Set olNewFolder = Nothing
2610 Set wsEmails = Nothing
2620 Set wb = Nothing
HandleExit:
    Exit Sub
HandleError:
2660      MsgBox "Error " & Err & " Line " & Erl() & "in " & cstrModule & "." & cstrProcedure
    Resume HandleExit
End Sub

Sub PDF_Save_Message()
     
    Dim Selection As Selection
    Dim obj As Object
    Dim Item As MailItem
     
 
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    
    Set Selection = Application.ActiveExplorer.Selection

For Each obj In Selection
    Set wrdApp = CreateObject("Word.Application")
    Set Item = obj
    
    Dim FSO As Object, TmpFolder As Object
    Dim sName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set tmpfilename = FSO.GetSpecialFolder(2)
    
    sName = Item.Subject
    ReplaceCharsForFileName sName, "-"
    tmpfilename = tmpfilename & "\" & sName & ".mht"
    
    Item.SaveAs tmpfilename, olMHTML
    
    
Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpfilename, Visible:=True)
  
    Dim WshShell As Object
    Dim SpecialPath As String
    Dim strToSaveAs As String
    Set WshShell = CreateObject("WScript.Shell")
    MyDocs = WshShell.SpecialFolders(16)
       
strToSaveAs = MyDocs & "\" & "EMAIL.pdf"
 
' check for duplicate filenames
' if matched, add the current time to the file name
If FSO.FileExists(strToSaveAs) Then
   sName = sName & Format(Now, "hhmmss")
   strToSaveAs = MyDocs & "\" & sName & ".pdf"
End If
  
wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    strToSaveAs, ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
    Range:=wdExportAllDocument, From:=0, To:=0, Item:= _
    wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
    BitmapMissingFonts:=True, UseISO19005_1:=False
             
    wrdDoc.Close
    wrdApp.Quit

Next obj
    
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    Set WshShell = Nothing
    Set obj = Nothing
    Set Selection = Nothing
    Set Item = Nothing
 
End Sub
 
Private Sub ReplaceCharsForFileName(sName As String, sChr As String)
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub



Public Sub ClearSubject()

Dim objApp As Outlook.Application
Dim aMail As Outlook.MailItem
Dim oSelection As Outlook.Selection

Set objApp = CreateObject("Outlook.Application")
Set oSelection = objApp.ActiveExplorer.Selection
  
'Remove from each email subject from selection
For Each aMail In oSelection
  aMail.Subject = Replace(aMail.Subject, "[EXTERNAL]", "")
  aMail.Save
Next aMail

End Sub