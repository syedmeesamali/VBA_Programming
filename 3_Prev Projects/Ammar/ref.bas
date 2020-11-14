Option Explicit
Public Sub Driving_Predecessors()
Dim n As Integer
    n = Sheets(Sheet1.Name).Cells(Rows.Count, "A").End(xlUp).Row
    Dim X As Integer
    'Main for loop
    For X = 2 To n
    
    Dim Criteria As String
    Criteria = Sheets(Sheet1.Name).Range("A" & X).Value
    Dim Data_range As Range
    Set Data_range = Sheets(Sheet1.Name).Range("A2:AD" & n)
    Dim activityStatus As String, PreDetails As String, activityType As String
    
    activityStatus = Application.WorksheetFunction.VLookup(Criteria, Data_range, 28, False)
    PreDetails = Application.WorksheetFunction.VLookup(Criteria, Data_range, 29, False)
    activityType = Application.WorksheetFunction.VLookup(Criteria, Data_range, 9, False)
    
        If (activityStatus <> "Completed") And (PreDetails <> "") And (activityType = "Task Dependent") Then
        Dim PredRaw() As String
        PredRaw() = Split(Application.WorksheetFunction.VLookup(Criteria, Data_range, 29, False), ", ")
        Dim H As String
        H = 0
        Dim Driv_Activity As String, Activity_Type As String, Status As String
        Driv_Activity = "No Drivers"
        
        Dim i As Long
            For i = LBound(PredRaw) To UBound(PredRaw)
                Dim Pred_ID As String
                PredRaw() = Split(Application.WorksheetFunction.VLookup(Criteria, Data_range, 29, False), ", ")
                Pred_ID = Left(PredRaw(i), InStr(1, PredRaw(i), ":", vbTextCompare) - 1)
                
                On Error Resume Next
                Activity_Type = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 9, False)
                Status = Application.VLookup(Pred_ID, Data_range, 28, False)
                
                If (Status <> "Completed") And (Activity_Type = "Task Dependent") Then
                    If InStr(1, PredRaw(i), "FS", vbTextCompare) <> 0 Then
                        Dim Pred_Date As String
                        Pred_Date = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 6, False)
                    ElseIf InStr(1, Pred_ID, "FF", vbTextCompare) <> 0 And (Len(PredRaw(i)) - 1 - InStr(1, PredRaw(i), "FF")) > 0 Then
                        Pred_Date = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 6, False) + Right(PredRaw(i), (Len(PredRaw(i)) - 2 - InStr(1, PredRaw(i), "FF")))
                    ElseIf InStr(1, Pred_ID, "FF", vbTextCompare) <> 0 Then
                        Pred_Date = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 6, False)
                    ElseIf InStr(1, Pred_ID, "SS", vbTextCompare) <> 0 And (Len(PredRaw(i)) - 1 - InStr(1, PredRaw(i), "SS")) > 0 Then
                        Pred_Date = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 5, False) + Right(PredRaw(i), (Len(PredRaw(i)) - 2 - InStr(1, PredRaw(i), "FF")))
                    ElseIf InStr(1, Pred_ID, "SS", vbTextCompare) <> 0 Then
                        Pred_Date = Application.WorksheetFunction.VLookup(Pred_ID, Data_range, 5, False)
                    End If
                End If
              
              If Pred_Date > H Then
                H = Pred_Date
                Driv_Activity = Pred_ID
              End If
            Next i
            'End of PredRaw
            If Driv_Activity <> "No Drivers" Then
                Driv_Activity = Pred_ID & "-->" & Driv_Activity
                Sheets(Sheet1.Name).Range("AE" & X).Value = Pred_ID
            End If
        Else
        End if 'End of main if blox
    Next X  'End of main for loop
End Sub

