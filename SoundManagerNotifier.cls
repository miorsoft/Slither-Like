VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundManagerNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Sub btnLoadAll_Click()
    Dim i As Integer
    On Error Resume Next
    For i = btnLoad.LBound To btnLoad.UBound
        If btnLoad(i).Enabled Then btnLoad_Click i
    Next
End Sub

Private Sub btnPlayAll_Click()
    Dim i As Integer
    On Error Resume Next
    For i = btnPlay.LBound To btnPlay.UBound
        If btnPlay(i).Enabled Then btnPlay_Click i
    Next
End Sub

Private Sub btnStopAll_Click()
    SoundManager.StopSound ALL_SOUND_BUFFERS
End Sub

Private Sub btnFreeAll_Click()
    SoundManager.FreeSound ALL_SOUND_BUFFERS
End Sub

Private Sub btnFree_Click(Index As Integer)
    SoundManager.FreeSound Index
End Sub

Private Sub btnLoad_Click(Index As Integer)
    SoundManager.LoadSoundFile Index, App.Path & "\Sounds\" & txtSoundFile(Index).Text
End Sub

Private Sub btnPlay_Click(Index As Integer)
    SoundManager.PlaySound Index
End Sub

Private Sub btnStop_Click(Index As Integer)
    SoundManager.StopSound Index
End Sub

Private Sub btnInstantSound_Click()
    ' Instantly load, play and free a sound using the first available buffer
    ' No notification is required
    SoundManager.LoadSoundFile SoundManager.FreeBuffer, App.Path & "\Sounds\Oshppis3.wav", BufferFlagInstant
End Sub

Private Sub Form_Load()
    Set SoundManager.Notifier = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Very, Very Important. This MUST be called or a crash is inevitable
    SoundManager.DestroySoundManager
End Sub

Private Sub SoundManagerNotifier_SoundLoaded(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Loaded"
        .ForeColor = vbBlue
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = True
    btnStop(BufferIndex).Enabled = False
End Sub

Private Sub SoundManagerNotifier_SoundPlayEnd(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Stopped"
        .ForeColor = vbBlue
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = True
    btnStop(BufferIndex).Enabled = False
End Sub

Private Sub SoundManagerNotifier_SoundPlayStart(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Playing"
        .ForeColor = vbGreen
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = False
    btnStop(BufferIndex).Enabled = True
End Sub

Private Sub SoundManagerNotifier_SoundUnloaded(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Empty"
        .ForeColor = RGB(127, 127, 127)
    End With
    
    btnFree(BufferIndex).Enabled = False
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = False
    btnStop(BufferIndex).Enabled = False
End Sub

