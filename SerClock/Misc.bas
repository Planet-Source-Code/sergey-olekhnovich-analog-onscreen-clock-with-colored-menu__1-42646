Attribute VB_Name = "mdlMisc"
Option Explicit

Public Enum enumSaveSettings
    ssColor
    ssRadius
    ssPositionX
    ssPositionY
    ssMenuStyle
End Enum

Public Const conRegApp = "SergeyO"
Public Const conRegSec = "Analog Clock"
Public Const conRegRad = "Size":        Public Const conRegRadDef = 2018
Public Const conRegCol = "Color":       Public Const conRegColDef = 1999
Public Const conRegPsX = "PositionX":   Public Const conRegPsDefX = -1
Public Const conRegPsY = "PositionY":   Public Const conRegPsDefY = -1
Public Const conRegGrf = "Show Graphics"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const conMaxChar& = 256&
Private Const conMenuClassName1 = "#32768"
Private Const conMenuClassName2 = "ToolbarWindow32"
Private Const conMenuClassName3 = "MsoCommandBarPopup"
Private Declare Function WindowFromPoint& _
                Lib "user32" ( _
                ByVal xPoint As Long, _
                ByVal yPoint As Long)
Private Declare Function GetClassName& _
                Lib "user32" _
                Alias "GetClassNameA" ( _
                ByVal hWnd As Long, _
                ByVal lpClassName As String, _
                ByVal nMaxCount As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const WM_DRAWITEM = &H2B
Const WM_MEASUREITEM = &H2C
Const WM_INITMENU = &H116
Const WM_INITMENUPOPUP = &H117

Public gbDigital As Boolean
Private lpOriginalWinProc&

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_CONTEXTMENU = &H7B
Private Const WM_PAINT = &HF
Private Const WM_NCPAINT = &H85
Private Const WM_EXITMENULOOP = &H212

Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
Private Const WM_CREATE = &H1
Private Const WM_MOUSEMOVE = &H200
Private Const WM_MOUSEFIRST = &H200
Public Declare Function SetWindowLong& _
                Lib "user32" _
                Alias "SetWindowLongA" ( _
                ByVal hWnd&, _
                ByVal nIndex&, _
                ByVal dwNewLong&)
Public Declare Function CallWindowProc& _
                Lib "user32" _
                Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc&, _
                ByVal hWnd&, _
                ByVal Msg&, _
                ByVal wParam&, _
                ByVal lParam&)
Private Declare Function PostMessage& _
                Lib "user32" _
                Alias "PostMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Long)
'---------------------------------------------------------------------
Public Type PointAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'---------------------------------------------------------------------
Public Const WINDING& = 2
Public Const ALTERNATE& = 1
Public Const Pi# = 3.14159265
'---------------------------------------------------------------------
Public Const HWND_TOPMOST& = -1
Public Const SWP_NOMOVE& = &H2
Public Const SWP_NOSIZE& = &H1
Public Const SWP_NOACTIVATE& = &H10
'---------------------------------------------------------------------
Public Const RGN_AND& = 1
Public Const RGN_OR& = 2
Public Const RGN_XOR& = 3
Public Const RGN_DIFF& = 4
Public Const RGN_COPY& = 5
'---------------------------------------------------------------------
Declare Function CombineRgn& Lib "gdi32" (ByVal hDestRgn&, ByVal hSrcRgn1&, ByVal hSrcRgn2&, ByVal nCombineMode&)
Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&)
Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&)
Declare Function CreatePolygonRgn& Lib "gdi32" (lpPoint As PointAPI, ByVal nCount&, ByVal nPolyFillMode&)
'Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Declare Function GetDC& Lib "user32" (ByVal hWnd&)
Declare Function GetPixel& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&)
Declare Function ReleaseDC& Lib "user32" (ByVal hWnd&, ByVal hdc&)
Declare Function SetWindowPos& Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cX&, ByVal cY&, ByVal wFlags&)
Declare Function SetWindowRgn& Lib "user32" (ByVal hWnd&, ByVal hRgn&, ByVal bRedraw&)

Private Declare Function BeginPath& _
                Lib "gdi32" ( _
                ByVal hdc As Long)
Private Declare Function EndPath& _
                Lib "gdi32" ( _
                ByVal hdc As Long)
Private Declare Function PathToRegion& _
                Lib "gdi32" ( _
                ByVal hdc As Long)
Private Declare Function CreateSolidBrush& _
                Lib "gdi32" ( _
                ByVal crColor As Long)
'---------------------------------------------------------------------
Private Declare Function ReleaseCapture& _
                Lib "user32" ()
Private Declare Function SendMessage& _
                Lib "user32" _
                Alias "SendMessageA" ( _
                ByVal hWnd&, _
                ByVal wMsg&, _
                ByVal wParam&, _
                lParam As Any)
'---------------------------------------------------------------------
Const MF_BITMAP = 4
Const MF_CHECKED = 8

'---------------------------------------------------------------------
Private Const WM_USER = &H400
'---------------------------------------------------------------------

Public glRad&, glHandL&, glHandS&, glLeft&, glTop&, gtPC As PointAPI

Public Sub gfCreateClock()
    Dim P As PointAPI, aP(3) As PointAPI
    Dim dAngle#, i%, R&, Rtmp&, lTpPX&, lTpPY&
    Dim dTime As Date, lShift&, lSize&, sTime$
    
    Call gfInitializeValues(glRad)
    Let dTime = Now

    lShift = Round(glRad / 300, 0): If lShift < 2 Then lShift = 2
    
    R& = 0&
    lTpPX& = Screen.TwipsPerPixelX
    lTpPY& = Screen.TwipsPerPixelY
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   remove following remark to view digital clocks
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Rem gbDigital = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If gbDigital Then
        With frmClock
            '   get font size
            sTime = Format$(dTime, "Hh:Nn AMPM")
            lSize& = .TextWidth(sTime) * lTpPX& * 1.05
            .Font.Size = .Font.Size / (lSize / .Width)
            .Cls
            '.CurrentX = (.Width - .TextWidth(sTime) * lTpPX) / 2
            '.CurrentY = (.Height - .TextHeight(sTime) * lTpPY) / 2
            Call BeginPath(.hdc)

            frmClock.Print sTime$

            Call EndPath(.hdc)

            R = PathToRegion(.hdc)
        End With
    Else
        '   1st -   create dots
        For i = 0 To 11
            dAngle# = (30 * i) / 180 * Pi
            Call fGetPoint(glRad, dAngle#, P, lTpPX&, lTpPY&)
            
            Rtmp& = CreateEllipticRgn(P.X - lShift, P.Y - lShift, P.X + lShift, P.Y + lShift)
            If CBool(R&) Then Call CombineRgn(R&, R&, Rtmp&, RGN_OR) Else R& = Rtmp&
        Next
        
        '   2nd -   create big hand (minute)
        dAngle# = DatePart("n", dTime) / 60 * 2 * Pi
        Call fGetPoint(glHandL, dAngle#, aP(0), lTpPX&, lTpPY&)
        'Call fGetPoint(glHandL / 3, dAngle# + (150 / 180 * Pi), aP(1), lTpPX&, lTpPY&)
        'Call fGetPoint(glHandL / 5, dAngle# + (210 / 180 * Pi), aP(2), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 11, dAngle# + Pi / 2, aP(1), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 5, dAngle# + Pi, aP(2), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 11, dAngle# + Pi / 2 * 3, aP(3), lTpPX&, lTpPY&)
        
        Rtmp& = CreatePolygonRgn&(aP(0), 4, WINDING)
        Call CombineRgn(R&, R&, Rtmp&, RGN_OR)
        
        '   3rd -   create small hand (hour)
        dAngle# = (DatePart("h", dTime) Mod 12) / 12 * 2 * Pi + dAngle / 12
        Call fGetPoint(glHandS, dAngle#, aP(0), lTpPX&, lTpPY&)
        'Call fGetPoint(glHandS / 3, dAngle# + (140 / 180 * Pi), aP(1), lTpPX&, lTpPY&)
        'Call fGetPoint(glHandS / 5, dAngle# + (220 / 180 * Pi), aP(2), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 9, dAngle# + Pi / 2, aP(1), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 5, dAngle# + Pi, aP(2), lTpPX&, lTpPY&)
        Call fGetPoint(glRad / 9, dAngle# + Pi / 2 * 3, aP(3), lTpPX&, lTpPY&)
        
        Rtmp& = CreatePolygonRgn&(aP(0), 4, WINDING)
        Call CombineRgn(R&, R&, Rtmp&, RGN_OR)
        
        'dAngle# = (DatePart("h", dTime) Mod 12) / 12 * 2 * Pi + dAngle / 12
'        dAngle# = 0
'        Call fGetPoint(30, dAngle#, P, lTpPX&, lTpPY&)
'        Rtmp& = CreateEllipticRgn(P.X - lShift * 0.8, _
'                                  P.Y - lShift, _
'                                  P.X + lShift * 1.2, _
'                                  P.Y + lShift)
'
'        Call CombineRgn(R&, R&, Rtmp&, RGN_XOR)
    End If
    
    Call SetWindowRgn(frmClock.hWnd, R, True)
End Sub

Private Sub fGetPoint(lRad&, dAngle#, P As PointAPI, Optional lDivX& = 1, Optional lDivY& = 1)
    P.X = (gtPC.X + lRad& * Sin(dAngle#)) / lDivX&
    P.Y = (gtPC.Y - lRad& * Cos(dAngle#)) / lDivY&
End Sub

Public Sub gfInitializeValues(Optional lRadius& = 0&)
    Dim lSize&
    
    'If glRad <> lRadius& Then
        lSize& = lRadius& * 2 + 150
        If frmClock.Width <> lSize Then
            frmClock.Width = lRadius& * 2 + 150
            frmClock.Height = lRadius& * 2 + 150
        End If
    'End If
    glRad = lRadius&
    
    glHandL = glRad * 0.9
    glHandS = glRad * 0.7
    
    gtPC.X = glRad + 120
    gtPC.Y = glRad + 120
    
    
    'frmClock.BackColor = vbBlack
End Sub

Public Sub gfDragForm(F As Form)
    Const conShift! = 20!
    On Error Resume Next
    With F
        If ReleaseCapture() Then Call SendMessage(.hWnd, &HA1, 2, 0&)
        If .Right > Screen.Width Then .Left = Screen.Width - .Width - conShift!
        If .Bottom > Screen.Height Then .Top = Screen.Height - .Height - conShift!
        If .Left < 0 Then .Left = conShift!
        If .Top < 0 Then .Top = conShift!
    End With
End Sub

Public Sub SetTopMost(ByVal hWnd&)
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
End Sub

Public Sub Main()
    frmClock.Hide
End Sub

Public Sub gfCheckBack()
    Dim hdc&, lC1&, lC2&, lC3&, lC4&, lC&, lTpPX&, lTpPY&
    Static lColor&
    
    lTpPX& = Screen.TwipsPerPixelX
    lTpPY& = Screen.TwipsPerPixelY
    
    hdc = GetDC(0)
    With frmClock

        lC1 = GetAvgRGB(GetPixel(hdc, .Left \ lTpPX&, .Top \ lTpPY&))
        lC2 = GetAvgRGB(GetPixel(hdc, (.Left + .Width) \ lTpPX&, .Top \ lTpPY&))
        lC3 = GetAvgRGB(GetPixel(hdc, (.Left + .Width) \ lTpPX&, (.Top + .Height) \ lTpPY&))
        lC4 = GetAvgRGB(GetPixel(hdc, .Left \ lTpPX&, (.Top + .Height) \ lTpPY&))
        
        lC = (lC1 + lC2 + lC3 + lC4) \ 4
        
        lC& = IIf(lC < 130, vbBlack, vbWhite)
        
        If lC = lColor Then GoTo ExitSub
        .BackColor = lColor
        .Refresh
        lColor = lC
    End With
ExitSub:
    Call ReleaseDC(0&, hdc)
End Sub
Private Function GetAvgRGB%(ByVal Color&)
    Dim iRed&, iGreen&, iBlue&
    
    iBlue = Color And &HFF&
    iGreen = Color \ &H100& And &HFF&
    iRed = Color \ &H10000 And &HFF&
    
    GetAvgRGB% = (iBlue + iRed + iGreen) \ 3
End Function

Public Sub gfWindowHook(hWnd&)
    If CBool(lpOriginalWinProc) Then Call gfWindowUnhook(hWnd)
    lpOriginalWinProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf gfWinProc)
End Sub

Public Sub gfWindowUnhook(hWnd&)
    If CBool(lpOriginalWinProc&) Then
        Call SetWindowLong(hWnd&, GWL_WNDPROC, lpOriginalWinProc&)
        lpOriginalWinProc& = 0&
    End If
End Sub

Public Function gfWinProc&(ByVal hWnd&, ByVal uMsg&, ByVal wParam&, ByVal lParam&)
    gfWinProc& = False
    
    'If uMsg And WM_MOUSEMOVE Then Call gfTT_ForceShow
    
    Select Case uMsg&
        Case WM_CLOSE, WM_DESTROY
            Call gfWindowUnhook(hWnd&)
            Call gfMenuPopUpDestroy
            On Error Resume Next
            Unload frmClock
            gfWinProc& = True
        Case WM_MEASUREITEM
            Call gfMeasureMenu(lParam&)
        Case WM_DRAWITEM
            Call gfDrawMenu(hWnd&, lParam&)
        Case Else
            '
    End Select
    If lpOriginalWinProc Then
        gfWinProc& = CallWindowProc&(lpOriginalWinProc, hWnd&, _
                                     uMsg&, wParam&, lParam&)
    End If
End Function

Public Sub gfGetSettings()
    geColor = GetSetting(conRegApp, conRegSec, conRegCol, conRegColDef)
    geSize = GetSetting(conRegApp, conRegSec, conRegRad, conRegRadDef)
    
    glRad = (geSize - miSizeMenu + 1) * 200
    
    glPosX = GetSetting(conRegApp, conRegSec, conRegPsX, conRegPsDefX)
    glPosY = GetSetting(conRegApp, conRegSec, conRegPsY, conRegPsDefY)
    
    gbGraph = GetSetting(conRegApp, conRegSec, conRegGrf, False)
End Sub

Public Sub gfSaveSettings(eWhat As enumSaveSettings, lValue&)
    If eWhat = ssColor Then
        Call SaveSetting(conRegApp, conRegSec, conRegCol, Str(lValue))
    ElseIf eWhat = ssRadius Then
        Call SaveSetting(conRegApp, conRegSec, conRegRad, Str(lValue))
    ElseIf eWhat = ssPositionX Then
        Call SaveSetting(conRegApp, conRegSec, conRegPsX, Str(lValue))
    ElseIf eWhat = ssPositionY Then
        Call SaveSetting(conRegApp, conRegSec, conRegPsY, Str(lValue))
    ElseIf eWhat = ssMenuStyle Then
        Call SaveSetting(conRegApp, conRegSec, conRegGrf, Str(lValue))
    Else:   '   can't be
    End If
End Sub

Public Function gfMenu() As Boolean
    Dim lTpPX&, lTpPY&
    
    On Error GoTo ExitFalse
    
    lTpPX& = Screen.TwipsPerPixelX
    lTpPY& = Screen.TwipsPerPixelY
    With frmClock
        If fWinClassName(.Left \ lTpPX&, .Top \ lTpPY&) Then GoTo ExitTrue
        If fWinClassName((.Left + .Width) \ lTpPX&, .Top \ lTpPY&) Then GoTo ExitTrue
        If fWinClassName((.Left + .Width) \ lTpPX&, (.Top + .Height) \ lTpPY&) Then GoTo ExitTrue
        If fWinClassName(.Left \ lTpPX&, (.Top + .Height) \ lTpPY&) Then GoTo ExitTrue
    End With
ExitFalse:
    Exit Function
ExitTrue:
    gfMenu = True
End Function

Private Function fWinClassName(X&, Y&) As Boolean
    Dim D$, lW$
    
    On Error GoTo ExitFunction
    lW = WindowFromPoint(X, Y)
    D$ = Space$(conMaxChar)
    Call GetClassName(lW, D$, conMaxChar - 1)
    D$ = Left$(D$, InStr(1, D$, vbNullChar) - 1)
    fWinClassName = StrComp(D$, conMenuClassName1, vbTextCompare) = False Or _
                    StrComp(D$, conMenuClassName2, vbTextCompare) = False Or _
                    StrComp(D$, conMenuClassName3, vbTextCompare) = False
ExitFunction:
End Function
