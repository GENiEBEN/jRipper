Attribute VB_Name = "About"
Option Explicit

Private Type Globe
    lPos As Long
    xSli As Long


    X As Single
    Y As Single
    sOff As Single

    vDrag As Boolean
    bDrag As Boolean
    sDrag As Boolean
    lDown As Boolean
    bClick As Boolean
    lMin As Boolean
    mMin As Boolean
    sTrack As String

End Type

Public GL As Globe
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub KeyChoice(k As Integer)

    Select Case k
     Case 13
      'Call GetPlay(True)
     Case 27
      'Call StopPlay
     Case 37
      'Call GotoTime(-5)
     Case 39
      'Call GotoTime(5)
     Case 46
      'Call RemoveItem
    End Select

End Sub

Public Function Allow(c As Control, Button As Integer, X As Single, Y As Single) As Boolean

    On Error Resume Next
    If Button = 1 And GL.bClick Then
     If X >= 0 And X <= c.Width And Y >= 0 And Y <= c.Height Then
      Allow = True
     Else
      Allow = False
     End If
    End If

End Function

Public Sub SHLabels(Value As Boolean)

    With NIMP
     If Value Then .picPar.Top = 2880
     .tmrSc.Enabled = CBool(True - Value)
     .lblMAb.Visible = CBool(False - Value): .picPar.Visible = CBool(True - Value)
    End With

End Sub

Public Sub SHLabels_2(Value As Boolean)

    With AboutJR
     If Value Then .picPar.Top = 2850
     .tmrSc.Enabled = CBool(True - Value)
     .lblMAb.Visible = CBool(False - Value): .picPar.Visible = CBool(True - Value)
    End With

End Sub

