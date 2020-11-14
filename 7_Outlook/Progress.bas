Private Const DefaultTitle = "Progress"
Private CurrentText As String
Private CurrentPercent As Single

Public Property Let Text(NewText As String)
  If NewText <> CurrentText Then
    CurrentText = NewText
    Me.Controls("UserText").Caption = CurrentText
    Call SizeToFit
  End If
End Property

Public Property Get Text() As String
  Text = CurrentText
End Property

Public Property Let Percent(NewPercent As Single)
  If NewPercent <> CurrentPercent Then
    CurrentPercent = Min(Max(NewPercent, 0#), 100#)
    Call UpdateProgress
  End If
End Property

Public Property Get Percent() As Single
  Percent = CurrentPercent
End Property

Public Sub Increment(ByVal NewPercent As Single, Optional ByVal NewText As String)
  Me.Percent = NewPercent
  If NewText <> "" Then Me.Text = NewText
  Call UpdateTitle
  Me.Repaint
End Sub

Private Sub UserForm_Initialize()
  Call SetupControls
  Call UpdateTitle
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then End 'Cancel = True
End Sub

Private Sub SetupControls()
  Dim i As Integer
  Dim Control As Object

  For i = Me.Controls.Count To 1 Step -1
    Me.Controls(i).Remove
  Next i

  Set Control = Me.Controls.Add("Forms.Label.1", "UserText", True)
  Control.Caption = ""
  Control.AutoSize = True
  Control.WordWrap = True
  Control.Font.Size = 8

  Set Control = Me.Controls.Add("Forms.Label.1", "ProgressFrame", True)
  Control.Caption = ""
  Control.Height = 16
  Control.SpecialEffect = fmSpecialEffectSunken

  Set Control = Me.Controls.Add("Forms.Label.1", "ProgressBar", True)
  Control.Caption = ""
  Control.Height = 14
  Control.BackStyle = fmBackStyleOpaque
  Control.BackColor = &HFF0000 ' Blue

  Call SizeToFit
End Sub

Private Sub SizeToFit()
  Me.Width = 240

  Me.Controls("UserText").Top = 6
  Me.Controls("UserText").Left = 6
  Me.Controls("UserText").AutoSize = False
  Me.Controls("UserText").Font.Size = 8
  Me.Controls("UserText").Width = Me.InsideWidth - 12
  Me.Controls("UserText").AutoSize = True

  Me.Controls("ProgressFrame").Top = Int(Me.Controls("UserText").Top + Me.Controls("UserText").Height) + 6
  Me.Controls("ProgressFrame").Left = 6
  Me.Controls("ProgressFrame").Width = Me.InsideWidth - 12
  Me.Controls("ProgressBar").Top = Me.Controls("ProgressFrame").Top + 1
  Me.Controls("ProgressBar").Left = Me.Controls("ProgressFrame").Left + 1

  Call UpdateProgress
  Me.Height = Me.Controls("ProgressFrame").Top + Me.Controls("ProgressFrame").Height + 6 + (Me.Height - Me.InsideHeight)
End Sub

Private Sub UpdateTitle()
  Me.Caption = DefaultTitle & " - " & Format(Int(CurrentPercent), "0") & "% Complete"
End Sub

Private Sub UpdateProgress()
  If CurrentPercent = 0 Then
    Me.Controls("ProgressBar").Visible = False
  Else
    Me.Controls("ProgressBar").Visible = True
    Me.Controls("ProgressBar").Width = Int((Me.Controls("ProgressFrame").Width - 2) * CurrentPercent / 100)
  End If
End Sub

Function Min(number1 As Single, number2 As Single) As Single
  If number1 < number2 Then Min = number1 Else Min = number2
End Function

Function Max(number1 As Single, number2 As Single) As Single
  If number1 > number2 Then Max = number1 Else Max = number2
End Function