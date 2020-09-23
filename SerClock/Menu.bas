Attribute VB_Name = "mdlMenu"
Option Explicit
'Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'Declare Function GetMenu Lib "user32" _
'                (ByVal hwnd As Long) As Long
'Declare Function GetSubMenu Lib "user32" _
'                (ByVal hMenu As Long, _
'                ByVal nPos As Long) As Long
'Declare Function GetMenuItemID Lib "user32" _
'                (ByVal hMenu As Long, _
'                ByVal nPos As Long) As Long
Private Declare Function SetMenuDefaultItem& _
                Lib "user32" ( _
                ByVal hMenu As Long, _
                ByVal uItem As Long, _
                ByVal fByPos As Long)
Public Enum genumSysColors
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum
Public Declare Function GetSysColor& _
                Lib "user32" ( _
                ByVal nIndex As genumSysColors)

Const RDW_INVALIDATE = &H1
Const BS_HATCHED = 2
Const HS_CROSS = 4
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public Enum enumMenuChecks
    michColors
    michSize
    michStyle
End Enum
Public Const conMENU_INCREMENT& = 1999
Public Enum enumMenuItemsID
    miGradient = 1000
    miAuto = conMENU_INCREMENT
    miQBColor0
    miQBColor1
    miQBColor2
    miQBColor3
    miQBColor4
    miQBColor5
    miQBColor6
    miQBColor7
    miQBColor8
    miQBColor9
    miQBColor10
    miQBColor11
    miQBColor12
    miQBColor13
    miQBColor14
    miQBColor15
    miSizeMenu
    miSizeMenu400
    miSizeMenu600
    miSizeMenu800
    miSizeMenu1000
    miShowText
    miShowGraph
    miExitProgram
    miClockOptions = 2000
End Enum

Public geColor   As enumMenuItemsID
Public geSize    As enumMenuItemsID
Public glPosX&, glPosY&
Public gbGraph As Boolean

Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function GetMenuItemInfo& _
                Lib "user32" _
                Alias "GetMenuItemInfoA" ( _
                ByVal hMenu As Long, _
                ByVal uItem As Long, _
                ByVal B As Boolean, _
                lpMII As MENUITEMINFO)
Private Declare Function SetMenuItemInfo& _
                Lib "user32" _
                Alias "SetMenuItemInfoA" ( _
                ByVal hMenu As Long, _
                ByVal uItem As Long, _
                ByVal fByPosition As Long, _
                lpMII As MENUITEMINFO)
Private Declare Function CreatePopupMenu& _
                Lib "user32" ()
Private Declare Function TrackPopupMenu& _
                Lib "user32" ( _
                ByVal hMenu As Long, _
                ByVal wFlags As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal nReserved As Long, _
                ByVal hWnd As Long, _
                ByVal lprc As Any)
Private Declare Function AppendMenu& _
                Lib "user32" _
                Alias "AppendMenuA" ( _
                ByVal hMenu As Long, _
                ByVal wFlags As Long, _
                ByVal wIDNewItem As enumMenuItemsID, _
                ByVal lpNewItem As Any)
Private Declare Function ModifyMenu& _
                Lib "user32" _
                Alias "ModifyMenuA" ( _
                ByVal hMenu As Long, _
                ByVal nPosition As Long, _
                ByVal wFlags As Long, _
                ByVal wIDNewItem As Long, _
                ByVal lpString As Any)
Private Declare Function DestroyMenu& _
                Lib "user32" ( _
                ByVal hMenu As Long)
Private Declare Function GetCursorPos& _
                Lib "user32" ( _
                lpPoint As POINTAPI)
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" ( _
                pDst As Any, _
                pSrc As Any, _
                ByVal ByteLen As Long)
Private Declare Function BitBlt& _
                Lib "gdi32" ( _
                ByVal hDestDC As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal nWidth As Long, _
                ByVal nHeight As Long, _
                ByVal hSrcDC As Long, _
                ByVal xSrc As Long, _
                ByVal ySrc As Long, _
                ByVal dwRop As Long)
Private Declare Function SetRect& _
                Lib "user32" ( _
                lpRect As RECT, _
                ByVal X1 As Long, _
                ByVal Y1 As Long, _
                ByVal X2 As Long, _
                ByVal Y2 As Long)
Private Declare Function DrawCaption& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hDC As Long, _
                pcRect As RECT, _
                ByVal un As Long)
Private Declare Function GetMenuItemRect& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hMenu As Long, _
                ByVal uItem As Long, _
                lprcItem As RECT)
Private Declare Function GetMenuItemCount& _
                Lib "user32" ( _
                ByVal hMenu As Long)
Private Declare Function GetPixel& _
                Lib "gdi32" ( _
                ByVal hDC As Long, _
                ByVal X As Long, _
                ByVal Y As Long)
Private Declare Function SetPixel& _
                Lib "gdi32" ( _
                ByVal hDC As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal crColor As Long)

Private Type MEASUREITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemWidth   As Long
    itemHeight  As Long
    itemData    As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemAction  As Long
    itemState   As Long
    hWndItem    As Long
    hDC         As Long
    rcItem      As RECT
    itemData    As Long
End Type

Private Const MF_APPEND = &H100&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_DEFAULT = &H1000&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_POPUP = &H10&
Private Const MF_REMOVE = &H1000&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_UNCHECKED = &H0&
Private Const MF_BITMAP = &H4&
Private Const MF_USECHECKBITMAPS = &H200&

Public Const MF_CHECKED = &H8&
Public Const MFT_RADIOCHECK = &H200&

Const TPM_RETURNCMD = &H100&

Public Const DC_GRADIENT = &H20
Public Const DC_ACTIVE = &H1
Public Const DC_ICON = &H4
Public Const DC_SMALLCAP = &H2
Public Const DC_TEXT = &H8

Public hMenu As Long
Public hSubMenu As Long
Public MP As POINTAPI, sMenu As Long
Public mnuHeight As Single


Private Function fGetResStr$(eID As enumMenuItemsID)
    On Error Resume Next
    
    fGetResStr$ = LoadResString(eID)
End Function

Private Sub fAppendMenu(lhMenu&, eID As enumMenuItemsID, Optional lFlag& = &H800&)
    Dim D$
    If CBool(eID) Then
        D$ = fGetResStr(eID)
        Call AppendMenu(lhMenu, lFlag, eID, D$)
    Else
        Call AppendMenu(lhMenu, lFlag&, 0&, 0&)
    End If
End Sub

Public Sub gfMenuPopUpDestroy()
    If CBool(hSubMenu) Then Call DestroyMenu(hSubMenu)
    If CBool(hMenu) Then Call DestroyMenu(hMenu)
End Sub
Public Sub gfMenuPopUpCreate(eColor As enumMenuItemsID, eSize As enumMenuItemsID)
    Const conFlag = MFT_RADIOCHECK Or MF_CHECKED
    
    'create the menu
    hMenu = CreatePopupMenu()
    hSubMenu = CreatePopupMenu()
    
    Call AppendMenu(hMenu, MF_OWNERDRAW Or MF_DISABLED, miGradient, 0&)  'SideBar
    
    Call AppendMenu(hMenu, MF_POPUP Or MF_MENUBARBREAK, hSubMenu, fGetResStr(miSizeMenu))
    Call fAppendMenu(hMenu, 0&) '   separator
    
    '   autocolor
    Call fAppendMenu(hMenu, miAuto, IIf(eColor = miAuto, conFlag, 0&))
    Call fAppendMenu(hMenu, 0&) '   separator
    
    '   colors
    Dim I&, j&, lFlag&
    If gbGraph Then
        With frmClock
            For I& = miQBColor0 To miQBColor15
                j = I - 2000
                lFlag = IIf(eColor = I, conFlag, 0&) Or MF_BITMAP
                lFlag = AppendMenu(hMenu, lFlag, I&, CLng(.picMenu(j).Picture))
            Next I&
        End With
    Else
        For I& = miQBColor0 To miQBColor15
            lFlag = IIf(eColor = I, conFlag, 0&)
            AppendMenu hMenu, lFlag, I, ByVal LoadResString(I)
        Next
    End If
    
    Call fAppendMenu(hMenu, 0&) '   separator
    
    Call fAppendMenu(hMenu, miShowText, IIf(gbGraph, 0&, conFlag))
    Call fAppendMenu(hMenu, miShowGraph, IIf(gbGraph, conFlag, 0&))
    
    Call fAppendMenu(hMenu, 0&) '   separator
       
    '   exit program
    Call fAppendMenu(hMenu, miExitProgram, 0&)
    
    '   submenus
    Call fAppendMenu(hSubMenu, miSizeMenu400, IIf(eSize = miSizeMenu400, conFlag, 0&))
    Call fAppendMenu(hSubMenu, miSizeMenu600, IIf(eSize = miSizeMenu600, conFlag, 0&))
    Call fAppendMenu(hSubMenu, miSizeMenu800, IIf(eSize = miSizeMenu800, conFlag, 0&))
    Call fAppendMenu(hSubMenu, miSizeMenu1000, IIf(eSize = miSizeMenu1000, conFlag, 0&))
       
    'Call DrawMenuBar(frmClock.hwnd)
    
    Call SetMenuDefaultItem(hMenu&, miExitProgram, False)
 End Sub

Public Sub gfMeasureMenu(ByRef lParam&)
    
    'It would appear that you cannot actually get measurements here,
    'you can only set them. There are no measurements until after the
    'Menu is drawn, but you only get a WM_MEASUREITEM message before the
    'initial WM_DRAWITEM.
    
    Dim MIS As MEASUREITEMSTRUCT
    'Load MIS with that in memory
    Call CopyMemory(MIS, ByVal lParam&, Len(MIS))
    
    MIS.itemWidth = 5   '(18 - 1) - 12. I don't know where the 12 comes
                        'from, but there always seems to be 12 pixels more than I want.
                        '18 is Small Titlebar height.
    
    'Return the updated MIS
    CopyMemory ByVal lParam&, MIS, Len(MIS)
    
End Sub

Public Sub gfDrawMenu(ByRef hWnd&, ByRef lParam&)
    
    Dim DIS As DRAWITEMSTRUCT, Rct As RECT, lRslt As Long
    
    Call CopyMemory(DIS, ByVal lParam&, Len(DIS))
    
    'since we can't measure in the MeasureMenu sub we'll do it here.
    'we cannot just get the sidebar height as it will only return
    'the height of an empty menu item. (i.e. 13). Maybe we can get the
    'height of the whole menu with some other API call that I don't know
    'about. I tried GetWindowRect.
    
    'String Menus
    Call GetMenuItemRect(hWnd, hMenu, 1, Rct)
    mnuHeight = (Rct.Bottom - Rct.Top) * (GetMenuItemCount(hMenu) - GetMenuItemCount(hSubMenu) - 1)
    'Separators
    GetMenuItemRect hWnd, hMenu, 3, Rct
    mnuHeight = mnuHeight + (Rct.Bottom - Rct.Top) * 4 ' 4 Seperators
    
    'set the size of our sidebar
    'SetRect rct, 0, 0, mnuHeight, 18
    
    'This is a bit of a copout, but it works
    'You could always use GradientFillRect and then draw rotated text
    'straight onto the sidebar, but this is much easier
    'you could use a hidden picturebox for this
    'Draw a form caption onto our userform, the length of our menu height
    
'    With frmClock
'    DrawCaption .hWnd, .hDC, Rct, DC_SMALLCAP Or DC_ACTIVE Or DC_TEXT Or DC_GRADIENT
'    End With
    Dim X!, Y!
    Dim nColor As Long
    
'''''''''''''''''''''''''''''''    'rotate our caption through 270 degrees
'''''''''''''''''''''''''''''''    'and paint onto menu
    Dim iBlue%, iRed%, iGreen%, lColor&, lTestColor&
    'nColor = GetSysColor(COLOR_ACTIVEBORDER)
    
    With frmClock.picM
        Debug.Print BitBlt(DIS.hDC, _
               2, mnuHeight - .Height - 35, _
               .Width, .Height, _
               .hDC, _
               0&, 0&, _
               vbSrcCopy)
    End With
    lTestColor& = GetPixel(DIS.hDC, 2, 2)
    For X = -50 To mnuHeight
        iBlue = 55 + 200 * (CDbl(X) / mnuHeight)
        For Y = 0 To 17
            lColor& = GetPixel(DIS.hDC, Y, mnuHeight - X)
            'Debug.Print Hex(lColor); ;
            If Not lColor = vbWhite Then
                Call SetPixel(DIS.hDC, Y, mnuHeight - X, RGB(iRed, iGreen, iBlue))
            End If
'''''''''''''''''''''''''''''''                'nColor = GetPixel(.hdc, X, Y)
'''''''''''''''''''''''''''''''                'SetPixel DIS.hdc, Y, mnuHeight - X, nColor
        Next Y
        'Debug.Print ""
    Next X
    
    'Call gfDrawText(DIS.hDC, frmClock.picM)
    
End Sub

Public Function gfMenuTrack&(hWnd&)
    Call GetCursorPos(MP)
    gfMenuTrack = TrackPopupMenu(hMenu, TPM_RETURNCMD, MP.X, MP.Y, 0&, hWnd&, 0&)
End Function

Public Sub gfMenuCaption()
    Call ModifyMenu(hMenu, miGradient, 0&, miGradient, ByVal CLng(frmClock.picM.Picture))
End Sub

Public Sub gfCheckUncheck(eItem As enumMenuItemsID, eWhat As enumMenuChecks)
    Dim I As enumMenuItemsID
    Dim lFlag& ', lRC&
    Const conFlag = MFT_RADIOCHECK Or MF_CHECKED
    
    'Exit Sub
    
    If eWhat = michColors Then
        For I = miAuto To miQBColor15
            lFlag = IIf(I = eItem, conFlag, 0&)
            If I = miAuto Then
                Call ModifyMenu(hMenu, I, lFlag, I, ByVal LoadResString(I))
            Else
                Call ModifyMenu(hMenu, I, lFlag Or MF_BITMAP, I, ByVal CLng(frmClock.picMenu(I - miQBColor0).Picture))
            End If
        Next
    ElseIf eWhat = michSize Then
        For I = miSizeMenu400 To miSizeMenu1000
            lFlag = IIf(I = eItem, conFlag, 0&)
            Call ModifyMenu(hMenu, I, lFlag, I, ByVal LoadResString(I))
        Next
    ElseIf eWhat = michStyle Then
        For I = miShowText To miShowGraph
            lFlag = IIf(I = eItem, conFlag, 0&)
            Call ModifyMenu(hMenu, I, lFlag, I, ByVal LoadResString(I))
        Next
    Else
        '   can't be
    End If
End Sub

Public Sub gfChangeMenuColors()
    Dim I&, lFlag&
    Const conFlag = MFT_RADIOCHECK Or MF_CHECKED
    
    For I = miQBColor0 To miQBColor15
        lFlag = IIf(geColor = I, conFlag, 0&)
        If gbGraph Then
            Call ModifyMenu(hMenu, I, lFlag Or MF_BITMAP, I, CLng(frmClock.picMenu(I - miQBColor0)))
        Else
            Call ModifyMenu(hMenu, I, lFlag, I, fGetResStr(I))
        End If
    Next
End Sub

Public Sub gfPopUpMenu()
    Dim P As POINTAPI
    
    Call GetCursorPos(P)
    Call frmClock.Form_MouseDown(vbRightButton, 0, _
                        CSng(P.X), _
                        CSng(P.Y))
    
End Sub
