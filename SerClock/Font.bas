Attribute VB_Name = "mdlFont"
Option Explicit

' API declarations.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long



Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Const ANSI_CHARSET = 0
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const CLIP_LH_ANGLES = 16
Private Const DEFAULT_QUALITY = 0
Private Const FF_DONTCARE = 0

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const WINDING = 2


' Weight values.
Private Enum FontWeights
    fwgt_Default = 0
    fwgt_Thin = 100
    fwgt_ExtraLight = 200
    fwgt_Light = 300
    fwgt_Normal = 400
    fwgt_Medium = 500
    fwgt_SemiBold = 600
    fwgt_Bold = 700
    fwgt_ExtraBold = 800
    fwgt_Heavy = 900
End Enum

' BorderStyle values:
Private Enum BorderStyles
    bs_None
    bs_Single
End Enum

' BackgroundStyle values:
Private Enum BackgroundStyles
    back_Transparent
    back_Opaque
End Enum


Private m_Caption           As String
Private m_Angle             As Single
Private m_FontName          As String
Private m_FontHeight        As Long
Private m_FontWidth         As Long
Private m_FontWeight        As FontWeights
Private m_FontItalic        As Boolean
Private m_FontUnderscore    As Boolean
Private m_FontStrikeOut     As Boolean
Private m_BorderStyle       As BorderStyles
Private m_BackgroundStyle   As BackgroundStyles
Private m_BackgroundColor   As Long 'OLE_COLOR
Private m_ForegroundColor   As Long 'OLE_COLOR


Private Function gfDrawText_Initialize()
    m_Caption = "SerClock Menu Options"
    m_Angle = 90!
    m_FontName = "Comic Sans MS"
    m_FontHeight = 16&
    m_FontWidth = 0
    m_FontWeight = fwgt_Bold
    m_FontItalic = False
    m_FontUnderscore = False
    m_FontStrikeOut = False
    m_BorderStyle = bs_None
    m_BackgroundStyle = back_Transparent
    m_BackgroundColor = 0
    m_ForegroundColor = vbWhite
End Function

Public Sub gfDrawText(hDC_&)
    Const PI = 3.14159265
    
    Dim hDC&
    Dim NewFont             As Long
    Dim OldFont             As Long
    Dim Escapement          As Long
    Dim TM                  As TEXTMETRIC
    Dim InternalLeading     As Long
    Dim TotalHeight         As Long
    Dim TextHeight          As Long
    Dim TextWidth           As Long
    Dim Theta               As Single
    Dim phi                 As Single
    Dim TextBoundWidth      As Single
    Dim TextBoundHeight     As Single
    Dim TotalBoundWidth     As Single
    Dim TotalBoundHeight    As Single
    Dim HeightDiff          As Single
    Dim WidthDiff           As Single
    Dim Pts(1 To 4)         As POINTAPI
    Dim X0                  As Single
    Dim Y0                  As Single
    Dim Rgn                 As Long
    
    Call gfDrawText_Initialize
    
    hDC = CreateCompatibleDC(hDC_)
    
    ' Convert the angle from degrees into degrees times 10.
    ' If the value is 0, use 360 degrees.
    Escapement = CLng(m_Angle * 10) Mod 3600
    If Escapement < 0 Then Escapement = Escapement + 3600
    If Escapement = 0 Then Escapement = 3600

    ' Create the new font.
    NewFont = CreateFont(m_FontHeight, m_FontWidth, _
                         Escapement, Escapement, _
                         m_FontWeight, m_FontItalic, m_FontUnderscore, m_FontStrikeOut, _
                         ANSI_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES, _
                         DEFAULT_QUALITY, FF_DONTCARE, m_FontName)

    ' Select the font.
    OldFont = SelectObject(hDC, NewFont)

    ' See how big the text is.
    GetTextMetrics hDC, TM
    With TM
        InternalLeading = .tmInternalLeading
        TotalHeight = .tmHeight
        TextHeight = TotalHeight - InternalLeading
        TextWidth = frmClock.TextWidth(m_Caption)
    End With

    ' Compute the bounding box.
    Theta = m_Angle * PI / 180
    phi = PI / 2 - Theta
    TextBoundWidth = Abs(TextHeight * Cos(phi)) + Abs(TextWidth * Cos(Theta))
    TextBoundHeight = Abs(TextHeight * Sin(phi)) + Abs(TextWidth * Sin(Theta))
    TotalBoundWidth = Abs(TotalHeight * Cos(phi)) + Abs(TextWidth * Cos(Theta))
    TotalBoundHeight = Abs(TotalHeight * Sin(phi)) + Abs(TextWidth * Sin(Theta))
    HeightDiff = TotalBoundHeight - TextBoundHeight
    WidthDiff = TotalBoundWidth - TextBoundWidth

    ' Find the text bounds.
    If Escapement <= 900 Then
        Pts(1).X = 0
        Pts(1).Y = TextWidth * Sin(Theta)
        Pts(2).X = TextWidth * Cos(Theta)
        Pts(2).Y = 0
        Pts(3).X = TextBoundWidth
        Pts(3).Y = TextHeight * Sin(phi)
        Pts(4).X = TextHeight * Cos(phi)
        Pts(4).Y = TextBoundHeight
        X0 = Pts(1).X - WidthDiff
        Y0 = Pts(1).Y - HeightDiff
    ElseIf Escapement <= 1800 Then
        Pts(1).X = -TextWidth * Cos(Theta)
        Pts(1).Y = TextBoundHeight
        Pts(2).X = 0
        Pts(2).Y = -TextHeight * Sin(phi)
        Pts(3).X = TextHeight * Cos(phi)
        Pts(3).Y = 0
        Pts(4).X = TextBoundWidth
        Pts(4).Y = TextWidth * Sin(Theta)
        X0 = Pts(1).X - WidthDiff
        Y0 = Pts(1).Y + HeightDiff
    ElseIf Escapement <= 2700 Then
        Pts(1).X = TextBoundWidth
        Pts(1).Y = -TextHeight * Sin(phi)
        Pts(2).X = -TextHeight * Cos(phi)
        Pts(2).Y = TextBoundHeight
        Pts(3).X = 0
        Pts(3).Y = -TextWidth * Sin(Theta)
        Pts(4).X = -TextWidth * Cos(Theta)
        Pts(4).Y = 0
        X0 = Pts(1).X + WidthDiff
        Y0 = Pts(1).Y + HeightDiff
    Else
        Pts(1).X = -TextHeight * Cos(phi)
        Pts(1).Y = 0
        Pts(2).X = TextBoundWidth
        Pts(2).Y = -TextWidth * Sin(Theta)
        Pts(3).X = TextWidth * Cos(Theta)
        Pts(3).Y = TextBoundHeight
        Pts(4).X = 0
        Pts(4).Y = TextHeight * Sin(phi)
        X0 = Pts(1).X + WidthDiff
        Y0 = Pts(1).Y - HeightDiff
    End If

    ' Size the control
'    Width = frmClock.ScaleX(TextBoundWidth, vbPixels, vbTwips)
'    Height = frmClock.ScaleY(TextBoundHeight, vbPixels, vbTwips)
'    UserControl.BackColor = m_BackgroundColor
'    UserControl.ForeColor = m_ForegroundColor

    ' See if the background should be transparent.
    If m_BackgroundStyle = vbTransparent Then
        ' Transparent. Make a region from the text.
        ' Reselect the font.
        SelectObject hDC, NewFont

        ' Draw the text and create a path.
        BeginPath hDC
        TextOut hDC, X0, Y0, m_Caption, Len(m_Caption)
        EndPath hDC

        ' Convert the path into a region.
        Rgn = PathToRegion(hDC)
    Else
        ' Opaque. Make a bounding box region.
        Rgn = CreatePolygonRgn(Pts(1), 4, WINDING)
    End If

'    ' Restrict the control to the region.
'    SetWindowRgn hWnd, Rgn, True

'    ' Erase the control.
'    Line (0, 0)-(ScaleWidth, ScaleHeight), m_BackgroundColor, BF
'    Picture = Image     ' This resets the font.

    ' Select the new font into the new hDC.
    OldFont = SelectObject(hDC, NewFont)

'    ' Print the caption.
'    CurrentX = X0
'    CurrentY = Y0
'    Print m_Caption
'    Picture = Image     ' This resets the font.

'    ' If we should draw a border, do so.
'    If m_BorderStyle = bs_Single Then
'        DrawWidth = 2
'        Polygon hDC, Pts(1), 4
'    End If

    ' Delete the rotated font.
    DeleteObject NewFont
End Sub


