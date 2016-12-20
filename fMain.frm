VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Slither Game"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14355
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
   ScaleHeight     =   665
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   957
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkBG 
      Caption         =   "BackGround"
      Height          =   495
      Left            =   13320
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkJPG 
      Caption         =   "Save Jpg Frames"
      Height          =   495
      Left            =   13320
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox chkBB 
      Caption         =   "Draw BB"
      Height          =   495
      Left            =   13320
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   615
      Left            =   13200
      TabIndex        =   1
      Top             =   120
      Width           =   1695
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

    PIC.Width = PIC.Height * 4 / 3

    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*") <> vbNullString Then Kill App.Path & "\Frames\*.*"

    InitRC
    InitResources

    Level = 1

    Set MultipleSounds = New clsSounds
    
DoBackGround = True



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DoLOOP = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  DestroySoundManager



    UnloadRC

End Sub



Private Sub Command1_Click()
    InitPool 6
    InitFOOD 100


    MainLoop

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Snake(PLAYER).FASTspeed = -1

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePos.x = x - CenX
    MousePos.y = y - CenY

End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Snake(PLAYER).FASTspeed = 0
End Sub
