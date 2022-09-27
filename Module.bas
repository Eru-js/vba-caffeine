Attribute VB_Name = "Module1"
Private dateNext As Date
Public CheckBoolean As Boolean
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" (Point As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Integer, ByVal y As Integer) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP As Long = &H10
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Move_Cursor()
    Dim Hold As POINTAPI
    GetCursorPos Hold
    x2 = Hold.x
    y2 = Hold.y
    StartTime = Timer
    CheckBoolean = False
    
    If Worksheets("Sheet1").Range("A1").Value = "" Then
        Worksheets("Sheet1").Range("A1").Value = "12:00:30 AM"
    End If
    Worksheets("Sheet1").Range("A1").NumberFormat = "hh:mm:ss"

    SecondsToActivate = Worksheets("Sheet1").Range("A1").Value
    SecondsToActivate = Hour(SecondsToActivate) * 3600 + Minute(SecondsToActivate) * 60 + Second(SecondsToActivate)
    Debug.Print SecondsToActivate
    
    Do
    DoEvents
    
    GetCursorPos Hold
    x1 = Hold.x
    y1 = Hold.y
    
    'Reset time if cursor manually moved
    If x1 <> x2 Or y1 <> Hold.y Then
        StartTime = Timer
    End If
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print SecondsElapsed
    
    'Automate cursor if timer done then reset timer
    If SecondsElapsed >= SecondsToActivate Then
        'SetCursorPos Hold.x + 60, Hold.y + 60
        For i = 1 To 500
            For j = 1 To 100
                SetCursorPos x1 + j, y1
            Next j
            For j = 99 To 0 Step -1
                SetCursorPos x1 + j, y1
            Next j
        Next i
    
        mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&
        Sleep 100
        mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
        Sleep 100
        StartTime = Timer
        'dateNext = DateAdd("s", 60, Now)
        'Application.OnTime dateNext, "Move_Cursor"
    End If
    
    'Exit loop
    If CheckBoolean = True Then
        Exit Do
    End If
    
    GetCursorPos Hold
    x2 = Hold.x
    y2 = Hold.y
    
    Loop
End Sub

Sub Stop_Cursor()
    'Application.OnTime dateNext, "Move_Cursor", , False
    CheckBoolean = True
    MsgBox "Process has been stopped"
End Sub


