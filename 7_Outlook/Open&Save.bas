Attribute VB_Name = "Module2"
Sub OpenAndSave()
    Dim strNewFolderName As String
    Dim save_to_folder As String
    Dim olkMsg As Outlook.MailItem, intCount As Integer
    Dim sFile As String, oAttachments As Outlook.Attachments
    Dim iCount As Long, strMessage As String
    Dim strTempFolder As String, strFileName As String, strFilePath As String
    Dim strFile As String, strPDF As String, j As Integer
    Dim stringSubject() As String, msgSubject As String, newSubj As String
    
    Const SpecialCharacters As String = "!,Chr(34),@,#,$,%,^,&,*,(,),{,[,],},?,<,>,;,:,/,\,"""""",'"
    
    'MsgBox Chr(34)
    For Each olkMsg In Outlook.ActiveExplorer.Selection
        
        '------------------------------------------------------------------------------------
        'Below portion is for creating FOLDER for each mail message with appropriate name
        '------------------------------------------------------------------------------------
        'msgSubject = olkMsg.Subject
        'Remove all the special characters
            For Each Char In Split(SpecialCharacters, ",")
                msgSubject = Replace(msgSubject, Char, "_")
            Next
        
        stringSubject() = Split(msgSubject, " ")
        
        'MsgBox "msgsubject: " + msgSubject
        'We want to use only first four words of subject in our new folders
        If UBound(stringSubject) > 4 Then
            msgSubject = stringSubject(0) & " " & stringSubject(1) & " " & _
                         stringSubject(2) & " " & stringSubject(3)
        Else
            msgSubject = msgSubject
        End If
        
        strNewFolderName = Format(olkMsg.SentOn, "YYYY-MM-DD_HHNN_") & msgSubject
        'MsgBox "String folder name: " + strNewFolderName
        'Make new folders
        MkDir ("C:\Users\SYED\Desktop\Docs\" & strNewFolderName)
        save_to_folder = "C:\Users\SYED\Desktop\Docs\" & strNewFolderName
        
        '-------------------------------------------
        'Below portion is for saving the attachments
        '-------------------------------------------
        
        Set oAttachments = olkMsg.Attachments
        iCount = oAttachments.Count
        If iCount > 0 Then
        'Processing each attachment for every selected email
        On Error Resume Next
        For i = iCount To 1 Step -1
            sFile = oAttachments.Item(i).FileName
            sFile = save_to_folder & "\" & sFile
            oAttachments.Item(i).SaveAsFile sFile
        Next i
        End If
        
        '---------------------------------------------------------------------------------------
        'Below portion is for making PDF of each email (body) and saving in respective folders
        '---------------------------------------------------------------------------------------
        
    Dim MyOlNamespace As Outlook.NameSpace
    Set MyOlNamespace = Application.GetNamespace("MAPI")
    Set MyOlSelection = Application.ActiveExplorer.Selection

    Set MySelectedItem = olkMsg
    
    'Get the user's TempFolder to store the item in
    Dim FSO As Object, TmpFolder As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    tmpfilename = save_to_folder & "\" & "Temp_File_1" & ".mht"
    
    'Save the mht-file
    MySelectedItem.SaveAs tmpfilename, olMHTML
    
    'Create a Word object
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    
    'Open the mht-file in Word without Word visible
    Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpfilename, Visible:=False)
    strCurrentFile = save_to_folder & "\" & "EMAIL.pdf"
     
        wrdApp.Documents.Open FileName:=tmpfilename
        wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            strCurrentFile, ExportFormat:= _
            wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
            Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=True, UseISO19005_1:=False
    
    wrdDoc.Close
    wrdApp.Quit
    Kill tmpfilename
    
    'Cleanup
    Set MyOlNamespace = Nothing
    Set MyOlSelection = Nothing
    Set MySelectedItem = Nothing
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    Set oRegEx = Nothing
        
    Next
    Set olkMsg = Nothing
    MsgBox "Emails attachments saved in their own respective folders respectively!"
End Sub

Function ReplaceCharsForSubject(sName As String, sChr As String)
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
End Function
