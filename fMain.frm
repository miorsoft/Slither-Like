VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slither Game"
   ClientHeight    =   9975
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   665
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1036
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicPanel 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   13440
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      Begin VB.HScrollBar hSnake 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CheckBox chkAI 
         Caption         =   "AI control"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "AI moves the player"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "START"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox chkBB 
         Caption         =   "Draw BB"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Debug Draw Bounding Boxes"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox chkJPG 
         Caption         =   "Save PNG Frames"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Save Frame to create a Video"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CheckBox chkBG 
         Caption         =   "BackGround"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Draw Background"
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   633
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   801
      TabIndex        =   0
      Top             =   120
      Width           =   12015
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub chkAI_Click()
AIcontrol = (chkAI.Value = vbChecked)
End Sub

Private Sub chkBB_Click()
    DrawBB = (chkBB.Value = vbChecked)
End Sub

Private Sub chkBG_Click()
    DoBackGround = (chkBG.Value = vbChecked)
End Sub

Private Sub chkJPG_Click()
    SaveFrames = (chkJPG.Value = vbChecked)
End Sub




Private Sub Form_Load()



RndM Timer

'MsgBox Cairo.CalcArc(2, 3)

    PIC.Width = PIC.Height * 4 / 3
    '    Stop

    PIC.Refresh

    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*") <> vbNullString Then Kill App.Path & "\Frames\*.*"

    InitRC
    InitResources

    Level = 1

    Set MultipleSounds = New cSounds

    DoBackGround = True

Open App.Path & "\LOG.txt" For Output As 1


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DoLOOP = False

End Sub

Private Sub Form_Resize()


    PIC.Width = Me.ScaleWidth '- 16
    PIC.Height = Me.ScaleHeight '- 16

    PIC.Width = (PIC.Width \ 8) * 8
    PIC.Height = (PIC.Height \ 8) * 8

PIC.Left = (Me.ScaleWidth - PIC.Width) \ 2
PIC.Top = (Me.ScaleHeight - PIC.Height) \ 2

    PicPanel.Top = PIC.Top
    PicPanel.Left = PIC.Left + PIC.Width - PicPanel.Width


    InitRC

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  DestroySoundManager


Close 1

    UnloadRC

End Sub



Private Sub Command1_Click()
    InitPool 6
    InitFOOD 6 * FoodXSnake                     '100

    Command1.Enabled = False

    MainLoop

End Sub

Private Sub hSnake_Change()
SNAKECAMERA = hSnake.Value

End Sub

Private Sub hSnake_Scroll()
SNAKECAMERA = hSnake.Value

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DoLOOP Then Snake(PLAYER).FASTspeed = -1

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePos.x = x - CenX
    MousePos.y = y - CenY

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DoLOOP Then Snake(PLAYER).FASTspeed = 0
End Sub
