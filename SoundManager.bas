Attribute VB_Name = "SoundManager"
Option Explicit

' **********************************************************
' * File Information                                       *
' * ================                                       *
' * File        : SoundManager,bas                         *
' * Author      : grigri <grigri@shinyhappypixels.com>     *
' * Description : How to play multiple sound files         *
' *               simultaneously in VB6.                   *
' * Version     : 1.0                                      *
' **********************************************************
' * Version History                                        *
' * ===============                                        *
' * 22/10/06   v1.0  Initial Version                       *
' **********************************************************

' =========== API Declares (lots, I'm afraid) =============

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)


Private Const WHDR_DONE As Long = &H1
Private Const WHDR_PREPARED As Long = &H2

Private Const CALLBACK_WINDOW As Long = &H10000


Private Const WAVE_MAPPED As Long = &H4
Private Const WAVE_MAPPER As Long = -1&

Private Const MMSYSERR_BASE As Long = 0
Private Const MMSYSERR_ALLOCATED As Long = (MMSYSERR_BASE + 4)
Private Const MMSYSERR_BADDB As Long = (MMSYSERR_BASE + 14)
Private Const MMSYSERR_BADDEVICEID As Long = (MMSYSERR_BASE + 2)
Private Const MMSYSERR_BADERRNUM As Long = (MMSYSERR_BASE + 9)
Private Const MMSYSERR_DELETEERROR As Long = (MMSYSERR_BASE + 18)
Private Const MMSYSERR_ERROR As Long = (MMSYSERR_BASE + 1)
Private Const MMSYSERR_HANDLEBUSY As Long = (MMSYSERR_BASE + 12)
Private Const MMSYSERR_INVALFLAG As Long = (MMSYSERR_BASE + 10)
Private Const MMSYSERR_INVALHANDLE As Long = (MMSYSERR_BASE + 5)
Private Const MMSYSERR_INVALIDALIAS As Long = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_INVALPARAM As Long = (MMSYSERR_BASE + 11)
Private Const MMSYSERR_KEYNOTFOUND As Long = (MMSYSERR_BASE + 15)
Private Const MMSYSERR_LASTERROR As Long = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_MOREDATA As Long = (MMSYSERR_BASE + 21)
Private Const MMSYSERR_NODRIVER As Long = (MMSYSERR_BASE + 6)
Private Const MMSYSERR_NODRIVERCB As Long = (MMSYSERR_BASE + 20)
Private Const MMSYSERR_NOERROR As Long = 0
Private Const MMSYSERR_NOMEM As Long = (MMSYSERR_BASE + 7)
Private Const MMSYSERR_NOTENABLED As Long = (MMSYSERR_BASE + 3)
Private Const MMSYSERR_NOTSUPPORTED As Long = (MMSYSERR_BASE + 8)
Private Const MMSYSERR_READERROR As Long = (MMSYSERR_BASE + 16)
Private Const MMSYSERR_VALNOTFOUND As Long = (MMSYSERR_BASE + 19)
Private Const MMSYSERR_WRITEERROR As Long = (MMSYSERR_BASE + 17)




Private Type WAVEHDR
    lpData     As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser     As Long
    dwFlags    As Long
    dwLoops    As Long
    lpNext     As Long
    Reserved   As Long
End Type

Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutOpen Lib "winmm.dll" (ByRef lphWaveOut As Long, ByVal uDeviceID As Long, ByRef lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long

'-------------

Private Const MMIO_ALLOCBUF As Long = &H10000
Private Const MMIO_COMPAT As Long = &H0
Private Const MMIO_CREATE As Long = &H1000
Private Const MMIO_CREATELIST As Long = &H40
Private Const MMIO_CREATERIFF As Long = &H20
Private Const MMIO_DEFAULTBUFFER As Long = 8192
Private Const MMIO_DELETE As Long = &H200
Private Const MMIO_DENYNONE As Long = &H40
Private Const MMIO_DENYREAD As Long = &H30
Private Const MMIO_DENYWRITE As Long = &H20
Private Const MMIO_DIRTY As Long = &H10000000
Private Const MMIO_EMPTYBUF As Long = &H10
Private Const MMIO_EXCLUSIVE As Long = &H10
Private Const MMIO_EXIST As Long = &H4000
Private Const MMIO_FHOPEN As Long = &H10
Private Const MMIO_FINDCHUNK As Long = &H10
Private Const MMIO_FINDLIST As Long = &H40
Private Const MMIO_FINDPROC As Long = &H40000
Private Const MMIO_FINDRIFF As Long = &H20
Private Const MMIO_GETTEMP As Long = &H20000
Private Const MMIO_GLOBALPROC As Long = &H10000000
Private Const MMIO_INSTALLPROC As Long = &H10000
Private Const MMIO_OPEN_VALID As Long = &H3FFFF
Private Const MMIO_PARSE As Long = &H100
Private Const MMIO_PUBLICPROC As Long = &H10000000
Private Const MMIO_READ As Long = &H0
Private Const MMIO_READWRITE As Long = &H2
Private Const MMIO_REMOVEPROC As Long = &H20000
Private Const MMIO_RWMODE As Long = &H3
Private Const MMIO_SHAREMODE As Long = &H70
Private Const MMIO_TOUPPER As Long = &H10
Private Const MMIO_UNICODEPROC As Long = &H1000000
Private Const MMIO_VALIDPROC As Long = &H11070000
Private Const MMIO_WRITE As Long = &H1
Private Const MMIOERR_BASE As Long = 256
Private Const MMIOERR_ACCESSDENIED As Long = (MMIOERR_BASE + 12)
Private Const MMIOERR_CANNOTCLOSE As Long = (MMIOERR_BASE + 4)
Private Const MMIOERR_CANNOTEXPAND As Long = (MMIOERR_BASE + 8)
Private Const MMIOERR_CANNOTOPEN As Long = (MMIOERR_BASE + 3)
Private Const MMIOERR_CANNOTREAD As Long = (MMIOERR_BASE + 5)
Private Const MMIOERR_CANNOTSEEK As Long = (MMIOERR_BASE + 7)
Private Const MMIOERR_CANNOTWRITE As Long = (MMIOERR_BASE + 6)
Private Const MMIOERR_CHUNKNOTFOUND As Long = (MMIOERR_BASE + 9)
Private Const MMIOERR_FILENOTFOUND As Long = (MMIOERR_BASE + 1)
Private Const MMIOERR_INVALIDFILE As Long = (MMIOERR_BASE + 16)
Private Const MMIOERR_NETWORKERROR As Long = (MMIOERR_BASE + 14)
Private Const MMIOERR_OUTOFMEMORY As Long = (MMIOERR_BASE + 2)
Private Const MMIOERR_PATHNOTFOUND As Long = (MMIOERR_BASE + 11)
Private Const MMIOERR_SHARINGVIOLATION As Long = (MMIOERR_BASE + 13)
Private Const MMIOERR_TOOMANYOPENFILES As Long = (MMIOERR_BASE + 15)
Private Const MMIOERR_UNBUFFERED As Long = (MMIOERR_BASE + 10)
Private Const MMIOM_CLOSE As Long = 4
Private Const MMIOM_OPEN As Long = 3
Private Const MMIOM_READ As Long = MMIO_READ
Private Const MMIOM_RENAME As Long = 6
Private Const MMIOM_SEEK As Long = 2
Private Const MMIOM_USER As Long = &H8000
Private Const MMIOM_WRITE As Long = MMIO_WRITE
Private Const MMIOM_WRITEFLUSH As Long = 5

Private Type MMCKINFO
    ckid       As Long
    ckSize     As Long
    fccType    As Long
    dwDataOffset As Long
    dwFlags    As Long
End Type

Private Type MMIOINFO
    dwFlags    As Long
    fccIOProc  As Long
    pIOProc    As Long
    wErrorRet  As Long
    htask      As Long
    cchBuffer  As Long
    pchBuffer  As String
    pchNext    As String
    pchEndRead As String
    pchEndWrite As String
    lBufOffset As Long
    lDiskOffset As Long
    adwInfo(4) As Long
    dwReserved1 As Long
    dwReserved2 As Long
    hmmio      As Long
End Type

Private Declare Function mmioAdvance Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByRef lpckParent As Any, ByVal uFlags As Long) As Long
Private Declare Function mmioFlush Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioGetInfo Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioInstallIOProc Lib "winmm.dll" Alias "mmioInstallIOProcA" (ByVal fccIOProc As Long, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function mmioInstallIOProcA Lib "winmm.dll" (ByVal fccIOProc As String, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, ByRef lpmmioinfo As Any, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByRef pch As Any, ByVal cch As Long) As Long
Private Declare Function mmioRename Lib "winmm.dll" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, ByRef lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long
Private Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function mmioSendMessage Lib "winmm.dll" (ByVal hmmio As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function mmioSetBuffer Lib "winmm.dll" (ByVal hmmio As Long, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioSetInfo Lib "winmm.dll" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioWrite Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Private Declare Function mmsystemGetVersion Lib "winmm.dll" () As Long

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels  As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize     As Integer
End Type

Private Const SEEK_SET As Long = 0

'------------ Window Handling Declarations (needed for the callback window)
Private Const MM_WOM_CLOSE As Long = &H3BC
Private Const MM_WOM_DONE As Long = &H3BD
Private Const MM_WOM_OPEN As Long = &H3BB
Private Const WM_DESTROY As Long = &H2
Private Const WM_CLOSE As Long = &H10


Private Const SS_SIMPLE As Long = &HB&
Private Const WS_POPUP As Long = &H80000000

Private Const GWL_WNDPROC As Long = -4


Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' ============ Non-API Declares; Internal Values ============

Private Const MAX_BUFFER_COUNT As Long = 64    ' 32

Public Enum SoundBufferFlags
    BufferFlagNone = 0
    BufferFlagAutoPlay = 1
    BufferFlagFreeWhenDone = 2
    BufferFlagNoNotify = 4

    ' This one's just for convenience
    BufferFlagInstant = BufferFlagAutoPlay Or BufferFlagFreeWhenDone Or BufferFlagNoNotify
End Enum

Public Enum SoundBufferStatus
    BufferError = -1
    BufferEmpty = 0
    BufferLoaded = 1
    BufferPlaying = 2
End Enum

Private Type SoundBufferInfo
    hWaveOut   As Long
    hdr        As WAVEHDR
    buf()      As Byte
    status     As SoundBufferStatus
    flags      As SoundBufferFlags
End Type

Public Const ALL_SOUND_BUFFERS As Long = -1

Private Buffers(1 To MAX_BUFFER_COUNT) As SoundBufferInfo

Private hCallbackWnd As Long
Private pfnOldWindowProc As Long

'''Public Notifier As SoundManagerNotifier '-------------------------------------<<<<<<<

Public Sub DestroySoundManager()
' Do not forget to call this when you're done.
    FreeSound ALL_SOUND_BUFFERS
    If hCallbackWnd <> 0 Then
        SetWindowLong hCallbackWnd, GWL_WNDPROC, pfnOldWindowProc
        DestroyWindow hCallbackWnd
    End If
End Sub

Private Function FindIndexFromHandle(ByVal hWaveOut As Long) As Long
' This should be optimized into a fast lookup routine, but
' for the small amount of buffers here it doesn't matter.
' (returns 0 if not found)
    Dim BufferIndex As Long
    For BufferIndex = 1 To MAX_BUFFER_COUNT
        If Buffers(BufferIndex).hWaveOut = hWaveOut Then
            FindIndexFromHandle = BufferIndex
            Exit Function
        End If
    Next
End Function

Public Function FreeBuffer() As Long
' Find the first free buffer (returns 0 if none found)


    Dim Index  As Long
    For Index = 1 To MAX_BUFFER_COUNT
        If Buffers(Index).status = BufferEmpty Then  '''' On XP works perfectlty
            '   If Buffers(Index).status <> BufferPlaying Then    '''' On Vista a little better (it seems not not works CALLBACK)
            FreeBuffer = Index
            Exit Function
        End If
    Next


End Function

Public Function SoundStatus(ByVal BufferIndex As Long) As SoundBufferStatus
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then
        SoundStatus = BufferError
        Exit Function
    End If
    SoundStatus = Buffers(BufferIndex).status
End Function

Public Function LoadSoundFile(ByVal BufferIndex As Long, ByVal FileName As String, Optional flags As SoundBufferFlags = BufferFlagNone) As Boolean
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then Exit Function

    ' Free any sound currently in the buffer
    FreeSound BufferIndex

    Dim InputHandle As Long
    Dim DataChunk As MMCKINFO
    Dim ParentChunk As MMCKINFO
    Dim InputChunk As MMCKINFO
    Dim EmptyChunk As MMCKINFO
    Dim MinSize As Long
    Dim WaveFCC As Long
    Dim RiffFCC As Long
    Dim WaveFormat As WAVEFORMATEX

    MinSize = Len(WaveFormat) - 2

    WaveFCC = mmioStringToFOURCC("WAVE", 0)
    RiffFCC = mmioStringToFOURCC("RIFF", 0)

    InputHandle = mmioOpen(FileName, ByVal 0&, MMIO_ALLOCBUF Or MMIO_READ)
    If InputHandle = 0 Then
        MsgBox "Cannot open file"
        InputHandle = 0
        Exit Function
    End If

    If mmioDescend(InputHandle, ParentChunk, ByVal 0&, 0) <> 0 Then
        MsgBox "Cannot descend"
        GoTo CLEARUP_AND_EXIT
    End If

    If ParentChunk.ckid <> RiffFCC Or ParentChunk.fccType <> WaveFCC Then
        MsgBox "Incorrect format"
        GoTo CLEARUP_AND_EXIT
    End If

    InputChunk.ckid = mmioStringToFOURCC("fmt", 0)

    If mmioDescend(InputHandle, InputChunk, ParentChunk, MMIO_FINDCHUNK) <> 0 Then
        MsgBox "Could not find fmt chunk"
        GoTo CLEARUP_AND_EXIT
    End If

    If InputChunk.ckSize < MinSize Then
        MsgBox "Not enough data, only " & InputChunk.ckSize & ", wanted " & MinSize
        GoTo CLEARUP_AND_EXIT
    End If

    If mmioRead(InputHandle, WaveFormat, LenB(WaveFormat)) < MinSize Then
        MsgBox "Not enough data read"
        GoTo CLEARUP_AND_EXIT
    End If

    If mmioSeek(InputHandle, ParentChunk.dwDataOffset + 4&, SEEK_SET) = -1 Then
        MsgBox "Could not seek"
        GoTo CLEARUP_AND_EXIT
    End If

    DataChunk = EmptyChunk

    DataChunk.ckid = mmioStringToFOURCC("data", 0)

    If mmioDescend(InputHandle, DataChunk, ParentChunk, MMIO_FINDCHUNK) <> 0 Then
        MsgBox "Could not descend into data"
        GoTo CLEARUP_AND_EXIT
    End If

    ' Make sure we have a callback window
    If hCallbackWnd = 0 Then
        If CreateCallbackWindow = False Then
            MsgBox "Cant CreateCallbackWindow"
            GoTo CLEARUP_AND_EXIT
        End If
    End If


    With Buffers(BufferIndex)
        ReDim .buf(0 To DataChunk.ckSize - 1)

        If mmioRead(InputHandle, .buf(0), DataChunk.ckSize) <> DataChunk.ckSize Then
            MsgBox "Could not read full buffer"
            GoTo CLEARUP_AND_EXIT
        End If

        Call waveOutOpen(.hWaveOut, WAVE_MAPPER, WaveFormat, hCallbackWnd, 0, CALLBACK_WINDOW)
        ' Prep the header
        .hdr.lpData = VarPtr(.buf(0))
        .hdr.dwBufferLength = UBound(.buf) - LBound(.buf) + 1



        Call waveOutPrepareHeader(.hWaveOut, .hdr, LenB(.hdr))

        .status = BufferLoaded
        .flags = flags

        LoadSoundFile = True

        ' Send notification if needed
        '<<<<  If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundLoaded(BufferIndex)

        ' Check for automatic playback
        If flags And BufferFlagAutoPlay Then
            PlaySound BufferIndex

        End If
    End With

CLEARUP_AND_EXIT:
    If InputHandle <> 0 Then
        Call mmioClose(InputHandle, 0)
        InputHandle = 0
    End If
End Function

Public Sub FreeSound(ByVal BufferIndex As Long)
' Handle the "all buffers" flag
    If BufferIndex = ALL_SOUND_BUFFERS Then
        For BufferIndex = 1 To MAX_BUFFER_COUNT
            If Buffers(BufferIndex).status <> BufferEmpty Then FreeSound BufferIndex
        Next
        Exit Sub
    End If

    If Buffers(BufferIndex).status = BufferEmpty Then Exit Sub

    ' If the sound is currently playing then we need to stop it first
    StopSound BufferIndex

    With Buffers(BufferIndex)
        ' Unprepare the header
        waveOutUnprepareHeader .hWaveOut, .hdr, LenB(.hdr)
        ' Close the handle
        Call waveOutClose(.hWaveOut)
        .hWaveOut = 0
        ' Erase the buffer
        Erase .buf
        ZeroMemory .hdr, LenB(.hdr)
        ' Set the status to empty
        .status = BufferEmpty

        Debug.Print "Sound " & BufferIndex & " Freed"
        '<<<<     If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundUnloaded(BufferIndex)
    End With
End Sub

Public Sub StopSound(ByVal BufferIndex As Long)
' Handle the "all buffers" flag
    If BufferIndex = ALL_SOUND_BUFFERS Then
        For BufferIndex = 1 To MAX_BUFFER_COUNT
            StopSound BufferIndex
        Next
        Exit Sub
    End If

    With Buffers(BufferIndex)
        Debug.Print .status
        If .status = BufferPlaying Then waveOutReset .hWaveOut
    End With
End Sub

Public Function PlaySound(ByVal BufferIndex As Long) As Boolean



' Check we've got a valid index
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then Exit Function

    StopSound BufferIndex
    With Buffers(BufferIndex)
        ' The sound must be loaded and not currently playing to be played
        If .status <> BufferLoaded Then Exit Function

        ' Ensure we have a valid handle loaded
        If .hWaveOut = 0 Then Exit Function

        ' Write the data to the sound device
        Call waveOutWrite(.hWaveOut, .hdr, LenB(.hdr))

        ' Update status
        .status = BufferPlaying

        ' Notify if required
        '<<<       If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundPlayStart(BufferIndex)
    End With

    ' All done!
    PlaySound = True
End Function

Private Function CreateCallbackWindow() As Boolean
    If hCallbackWnd <> 0 Then Exit Function
    hCallbackWnd = CreateWindowEx(0, "STATIC", "Soundmanager Window", WS_POPUP Or SS_SIMPLE, 0, 0, 100, 20, 0, 0, App.hInstance, ByVal 0&)
    If hCallbackWnd = 0 Then Exit Function
    pfnOldWindowProc = SetWindowLong(hCallbackWnd, GWL_WNDPROC, AddressOf CallbackWindowProc)

    CreateCallbackWindow = True
End Function

Private Function CallbackWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim BufferIndex As Long


    Select Case uMsg
        '    Case MM_WOM_OPEN
        '    Case MM_WOM_CLOSE
    Case MM_WOM_DONE
        BufferIndex = FindIndexFromHandle(wParam)
        If BufferIndex <> 0 Then
            With Buffers(BufferIndex)
                '<<<               If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundPlayEnd(BufferIndex)
                .status = BufferLoaded

                ' Automatic Free?
                If .flags And BufferFlagFreeWhenDone Then
                    FreeSound BufferIndex
                End If


            End With
        End If
    End Select
    CallbackWindowProc = CallWindowProc(pfnOldWindowProc, hWnd, uMsg, wParam, lParam)
End Function




Public Sub MyPlaySound(S As String)



'    Dim BufferIndex As Long
    LoadSoundFile SoundManager.FreeBuffer, App.Path & "\snd\" & S, BufferFlagInstant

    '    For BufferIndex = 1 To MAX_BUFFER_COUNT
    '        If Buffers(BufferIndex).status <> BufferEmpty Then
    '            If Buffers(BufferIndex).status <> BufferPlaying Then FreeSound BufferIndex
    '        End If
    '    Next

End Sub
