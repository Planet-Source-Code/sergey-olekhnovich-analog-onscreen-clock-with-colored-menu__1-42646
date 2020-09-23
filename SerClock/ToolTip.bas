Attribute VB_Name = "mdlToolTip"
Option Explicit

Private Type InitCommonControlsExType
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type
Const ICC_LISTVIEW_CLASSES = &H1       ' listview, header
Const ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
Const ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
Const ICC_TAB_CLASSES = &H8            ' tab, tooltips
Const ICC_UPDOWN_CLASS = &H10          ' updown
Const ICC_PROGRESS_CLASS = &H20        ' progress
Const ICC_HOTKEY_CLASS = &H40          ' hotkey
Const ICC_ANIMATE_CLASS = &H80         ' animate
Const ICC_WIN95_CLASSES = &HFF
Const ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
Const ICC_USEREX_CLASSES = &H200       ' comboex
Const ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
Const ICC_INTERNET_CLASSES = &H800
Const ICC_PAGESCROLLER_CLASS = &H1000  ' page scroller
Const ICC_NATIVEFNTCTL_CLASS = &H2000  ' native font control

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (init As InitCommonControlsExType) As Boolean


Private Declare Function CreateWindowEx& _
                Lib "user32" _
                Alias "CreateWindowExA" ( _
                ByVal dwExStyle As Long, _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String, _
                ByVal dwStyle As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long, _
                ByVal hWndParent As Long, _
                ByVal hMenu As Long, _
                ByVal hInstance As Long, _
                lpParam As Any)
Private Declare Function SetWindowPos& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cX As Long, _
                ByVal cY As Long, _
                ByVal wFlags As Long)
Private Declare Function SendMessage& _
                Lib "user32" _
                Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any)
Private Declare Function GetClientRect& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                lpRect As RECT)
Private Declare Function DestroyWindow& _
                Lib "user32" ( _
                ByVal hWnd As Long)

' A RECT user defined type. This is used for setting the bounds of the tool tip window.
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' A TOOLINFO user defined type. This is used for setting 'all of the necessary
' flags when creating a tool tip window.
Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uid As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

' A constant used in conjunction with the CreateWindowEx 'API. It indicates to use the default value.
Private Const CW_USEDEFAULT = &H80000000

' Constants used with the SetWindowPosition API.
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1

' Constants for setting the style of the tool tip window.
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&

' A constant used with the SendMessage API to define 'private messages.
Private Const WM_USER = &H400

' Messages used for setting the duration time of tool 'tips.
' Not used here.
Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

' All of the flags for tool tip windows.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

' All of the available messages for tool tip windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

' Constants for setting the style of the tool tip window.
'
' Always tip, even if the parent window is not active.
Private Const TTS_ALWAYSTIP = &H1
'
' Use the balloon style tool tip. <used here>
Private Const TTS_BALLOON = &H40
'
' Win98 and up - do not use sliding tool tips.
Private Const TTS_NOANIMATE = &H10
'
' Win2K and up - do not fade in tool tips.
Private Const TTS_NOFADE = &H20
'
' Prevents windows from removing any ampersand characters 'in the tool tip
' string. Without this flag, Windows will automatically 'remove ampersand
' characters from the string. This is done to allow the 'same string to be
' used as the tool tip text, and as the caption of a 'control.
Private Const TTS_NOPREFIX = &H2


' The two different tool tip classes.
Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

' A long for containing the hwnd (window handle) of the 'tool tip window created in
' this demo. This would have to be an array of longs if 'we were to create
' tool tip windows for multiple windows/controls.
Dim hWndTT As Long


Public Sub gfTT_Create(hWndParent&, sTT$)
    Dim ti As TOOLINFO
    Dim R As RECT
    Dim lRV&, hInstance& ', sTTC$
    
    If Not CBool(hWndParent) Then Exit Sub
    'sTTC = StrConv(TOOLTIPS_CLASSA, vbUnicode)
    
    If Not CBool(hWndTT) Then If Not fInitializeCommonControls Then Exit Sub
    
    hWndTT = CreateWindowEx(WS_EX_TOPMOST, _
                            TOOLTIPS_CLASSA, _
                            vbNullString, _
                            WS_POPUP Or TTS_NOPREFIX Or TTS_BALLOON Or TTS_ALWAYSTIP, _
                            CW_USEDEFAULT, _
                            CW_USEDEFAULT, _
                            CW_USEDEFAULT, _
                            CW_USEDEFAULT, _
                            hWndParent&, _
                            0&, _
                            App.hInstance, _
                            0&)
    If Not CBool(hWndTT) Then
        MsgBox Err.LastDllError
        Exit Sub
    End If
    
'    Call SetWindowPos(hWndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
    Call GetClientRect(hWndParent&, R)
    
    '   Set all of the necessary info in the toolinfo UDT.
    With ti
        '   The size of the toolinfo UDT in bytes. Must be set!
        .cbSize = Len(ti)
        
        '   The flags that we want to pass to the tool tip. 'TTF_CENTERTIP is not
        '   necessary, but centers the tool tip to the window it is 'being applied to
        '   (when possible). TTF_SUBCLASS tells the tool tip window 'to subclass the
        '   window it is being applied to. This is the best route 'to take in VB, so
        '   subclassing by the developer is not necessary.
        .uFlags = TTF_CENTERTIP Or TTF_SUBCLASS
        '.uFlags = TTF_SUBCLASS
        
        '   The hwnd of the control having the tool tip applied.
        .hWnd = hWndParent&
        
        '   The instance of the app the tool tip applies to.
        .hinst = App.hInstance
        
        '   The ID (hwnd) of the tool tip window. Not necessary 'unless the window is
        '   created using the TTF_IDISHWND flag.
        ti.uid = 0&
        
        '   A pointer to the tool tip text.
        ti.lpszText = sTT
        
        '   The coordinates that specify the tool tip window's
        '   region of activation.
        With .RECT
            .Left = R.Left
            .Right = R.Right
            .Top = R.Top
            .Bottom = R.Bottom
        End With
    End With
    
    lRV = SendMessage(hWndTT, TTM_ADDTOOLA, 0&, ti)
    
    '   Send a message to the tool tip window telling it to set 'the maximum tip
    '   width, to allow line breaking.
    lRV = SendMessage(hWndTT, TTM_SETMAXTIPWIDTH, 0&, 80&)

    '   Send messages to the tool tip window telling it what 'it's fore and back
    '   colours are.
'    lRV = SendMessage(hWndTT, TTM_SETTIPBKCOLOR, RGB(255, 255, 255), 0)
    lRV = SendMessage(hWndTT, TTM_SETTIPTEXTCOLOR, RGB(0, 0, 128), 0&)

    '   Send a message to the tool tip window telling it to 'update itself
    '   (to reflect the new fore and back colours).
    lRV = SendMessage(hWndTT, TTM_UPDATETIPTEXTA, 0&, ti)
End Sub

Public Sub gfTT_ModifyText(sTT$)
    Dim ti As TOOLINFO
    If hWndTT Then
        ti.cbSize = Len(ti)
        If SendMessage(hWndTT, TTM_ENUMTOOLSA, 0, ti) Then
            ti.lpszText = sTT
            Call SendMessage(hWndTT, TTM_SETTOOLINFOA, 0, ti) ' &H141
      End If   ' TTM_ENUMTOOLSA
        
    End If
End Sub

Public Sub gfTT_ForceShow()
    If hWndTT Then Call SendMessage(hWndTT, TTM_POP, True, 0&)
End Sub

Public Sub gfTT_Destroy()
    If hWndTT& Then Call DestroyWindow(hWndTT&)
End Sub

Private Function fInitializeCommonControls() As Boolean
    'Const IE3_INSTALLED = True
    
    Dim initCC As InitCommonControlsExType
    
    On Error GoTo ExitFalse
    
    initCC.dwSize = Len(initCC)
    initCC.dwICC = ICC_BAR_CLASSES
    If InitCommonControlsEx(initCC) Then
        GoTo ExitTrue
    Else
        'Call InitCommonControls
        GoTo ExitFalse
    End If
    
ExitTrue:
    fInitializeCommonControls = True
    Exit Function
ExitFalse:
    If Err.Number Then Err.Clear
    fInitializeCommonControls = False
End Function
