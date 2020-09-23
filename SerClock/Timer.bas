Attribute VB_Name = "mdlTimer"
Option Explicit

Private Declare Function SetTimer& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal nIDEvent As Long, _
                ByVal uElapse As Long, _
                ByVal lpTimerFunc As Long)
Private Declare Function KillTimer& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal nIDEvent As Long)

Private m_Timer     As Long

Public Function gfSetTimer(Optional lMilliSeconds& = &H36EE80) As Boolean
    Call gfKillTimer
    
    m_Timer = SetTimer(0, 0, lMilliSeconds&, AddressOf TimerProc)
    gfSetTimer = CBool(m_Timer)
End Function

Public Function TimerProc&(ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
    Dim D$
    
    Call gfKillTimer
    
    D$ = App.Path
    If Not Right$(D$, 1) = "\" Then D$ = D$ & "\"
    D$ = D$ & App.EXEName
    
    On Error GoTo ErrOccurred
    Call Shell(D$, vbNormalNoFocus)
    
    Unload frmClock
    'DoEvents
    'Load frmClock
ErrOccurred:
    
End Function

Public Function gfKillTimer() 'Optional hWnd& = 0&) ', Optional nIDEvent = 0&)
    If CBool(m_Timer) Then Call KillTimer(0&, m_Timer)
End Function
