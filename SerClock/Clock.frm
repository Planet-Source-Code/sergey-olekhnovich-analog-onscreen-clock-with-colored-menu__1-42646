VERSION 5.00
Begin VB.Form frmClock 
   BorderStyle     =   0  'None
   ClientHeight    =   6450
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10530
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   702
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2730
      Left            =   7350
      Picture         =   "Clock.frx":0442
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   3390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   15
      Top             =   5700
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   14
      Top             =   5340
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   13
      Top             =   4980
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   12
      Top             =   4620
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   11
      Top             =   4260
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   10
      Top             =   3900
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   9
      Top             =   3540
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   8
      Top             =   3180
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   7
      Top             =   2820
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   6
      Top             =   2460
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   5
      Top             =   2100
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   4
      Top             =   1740
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   3
      Top             =   1380
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   2
      Top             =   1020
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   1
      Top             =   660
      Width           =   2385
   End
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7650
      ScaleHeight     =   345
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   300
      Width           =   2385
   End
   Begin VB.Timer T 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9390
      Top             =   5970
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conFormatLongDate = "Long Date"
Private lMinute& ', bMenu As Boolean

Private Sub Form_Activate()
    Call gfTT_Create(Me.hWnd, Format$(Date, conFormatLongDate))
    Call gfDragForm(Me)
End Sub

Private Sub Form_DblClick()
    Const conMsg$ = "Do you want to quit the program?"
    Const conTitle$ = "Analog Clock"
    If MsgBox(conMsg, vbQuestion Or vbYesNo, conTitle) = vbYes Then
        Call gfWindowUnhook(Me.hWnd)
        Unload Me
    End If
End Sub

'Private Sub Form_KeyDown(iKeyCode%, iShift%)
'    If iKeyCode =  And iShift =  Then Call gfPopUpMenu
'End Sub

Private Sub Form_Load()
    Dim lV&
    
    Call fSetMousePointer
    
    lMinute = -1
    Me.ScaleMode = vbPixels
    
    Call gfGetSettings
    Call fLoadGraphics
    
    If geColor = miAuto Then
        Me.BackColor = vbBlack
    Else
        Me.BackColor = QBColor(geColor - miQBColor0)
    End If
    
    Call gfCreateClock
    
    Me.Show
    
    If glPosX = conRegPsDefX And glPosY = conRegPsDefY Then
        glPosX = Screen.Width - glRad * 3
        glPosY = Screen.Height - glRad * 4
        
        Call gfSaveSettings(ssPositionX, glPosX)
        Call gfSaveSettings(ssPositionY, glPosY)
    End If
    
    Me.Left = IIf(glPosX < 0, 0, glPosX)
    Me.Top = IIf(glPosY < 0, 0, glPosY)
    
    Call gfMenuPopUpCreate(geColor, geSize)
    Call gfWindowHook(Me.hWnd)
    
    T.Enabled = True
    
    Call gfSetTimer
End Sub

Public Sub Form_MouseDown(iButton%, iShift%, X!, Y!)
    Dim eR As enumMenuItemsID, sMsg$
    If iButton = vbRightButton Then
        'bMenu = True
        eR = gfMenuTrack(Me.hWnd)
        'bMenu = False
        Select Case eR
            Case enumMenuItemsID.miExitProgram
                Call Form_DblClick
            Case enumMenuItemsID.miAuto To enumMenuItemsID.miQBColor15
                Call gfCheckUncheck(eR, michColors)
                geColor = eR
                Call gfSaveSettings(ssColor, eR)
                If eR <> miAuto Then
                    Me.BackColor = QBColor(eR - enumMenuItemsID.miQBColor0)
                Else
                    Me.BackColor = vbBlack
                End If
            Case enumMenuItemsID.miSizeMenu400 To enumMenuItemsID.miSizeMenu1000
                geSize = eR
                Call gfCheckUncheck(eR, michSize)
                Call gfSaveSettings(ssRadius, eR)
                glRad = (eR - miSizeMenu + 1) * 200
                
                Call gfInitializeValues(glRad)
                Call gfCreateClock
                
                Call gfTT_Destroy
                Call gfTT_Create(Me.hWnd, Format$(Date, conFormatLongDate))
            Case enumMenuItemsID.miShowText
                If gbGraph Then
                    Call gfCheckUncheck(eR, michStyle)
                    gbGraph = False
                    Call gfChangeMenuColors
                    Call gfSaveSettings(ssMenuStyle, False)
                End If
            Case enumMenuItemsID.miShowGraph
                If Not gbGraph Then
                    Call gfCheckUncheck(eR, michStyle)
                    gbGraph = True
                    Call gfChangeMenuColors
                    Call gfSaveSettings(ssMenuStyle, True)
                End If
            Case Else
        End Select
    ElseIf iButton = vbLeftButton Then
        Call fSetMousePointer(False)
        Call gfDragForm(Me)
        Call fSetMousePointer(True)
        
        glPosX = Me.Left:   Call gfSaveSettings(ssPositionX, glPosX)
        glPosY = Me.Top:    Call gfSaveSettings(ssPositionY, glPosY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print "Over"
    'Debug.Print SendMessage(Me.hWnd, TVM_GETTOOLTIPS, 0&, 0&)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gfKillTimer
    Call gfTT_Destroy
    Call gfWindowUnhook(Me.hWnd)
End Sub

Public Sub T_Timer()
    Dim lM&
    
    If Not gfMenu Then Call SetTopMost(Me.hWnd)
    'If Not bMenu Then Call SetTopMost(Me.hwnd)
    lM = DatePart("n", Now)
    If lM <> lMinute Then Call gfCreateClock:    lMinute = lM: Call gfTT_ModifyText(Format$(Date, conFormatLongDate))
    
    If geColor = miAuto Then Call gfCheckBack
End Sub

Private Sub fLoadGraphics()
    Dim I%
    For I = 0 To 15
        Set picMenu(I).Picture = LoadResPicture(I + miQBColor0, vbResBitmap)
        picMenu(I).Refresh
    Next
End Sub

Private Sub fSetMousePointer(Optional bDefault As Boolean = True)
    On Error Resume Next
    Set Me.MouseIcon = LoadResPicture(IIf(bDefault, 101, 102), vbResCursor)
End Sub

Public Property Get Right!()
    Right! = Me.Left + Me.Width
End Property

Public Property Get Bottom!()
    Bottom! = Me.Top + Me.Height
End Property
