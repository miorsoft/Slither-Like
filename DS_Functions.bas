Attribute VB_Name = "DS_Functions"
Option Explicit

'Author: TheTrick

Private Type CHUNK
    id         As Long
    szData     As Long
End Type

Private Type curBuffer
    b(15)      As Currency
End Type

Private Type mp3Const
    bitrate(1, 15) As Integer
    smprate(2, 3) As Long
End Type

Private Type LARGE_INTEGER
    lowpart    As Long
    highpart   As Long
End Type

Private Type MPEGLAYER3WAVEFORMAT
    wFormatTag As Integer
    nChannels  As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize     As Integer
    wID        As Integer
    fdwFlags   As Long
    nBlockSize As Integer
    nFramesPerBlock As Integer
    nCodecDelay As Integer
End Type

Private Type ACMSTREAMHEADER
    cbStruct   As Long
    fdwStatus  As Long
    lpdwUser   As Long
    lppbSrc    As Long
    cbSrcLength As Long
    cbSrcLengthUsed As Long
    lpdwSrcUser As Long
    lppbDst    As Long
    cbDstLength As Long
    cbDstLengthUsed As Long
    lpdwDstUser As Long
    dwDriver(9) As Long
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingW" (ByVal hFile As Long, lpFileMappingAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long
Private Declare Function memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare Function SetFilePointerEx Lib "kernel32" (ByVal hFile As Long, ByVal liDistanceToMoveL As Long, ByVal liDistanceToMoveH As Long, ByRef lpNewFilePointer As LARGE_INTEGER, ByVal dwMoveMethod As Long) As Long
Private Declare Function acmStreamClose Lib "msacm32" (ByVal has As Long, ByVal fdwClose As Long) As Long
Private Declare Function acmStreamConvert Lib "msacm32" (ByVal has As Long, ByRef pash As ACMSTREAMHEADER, ByVal fdwConvert As Long) As Long
Private Declare Function acmStreamMessage Lib "msacm32" (ByVal has As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function acmStreamOpen Lib "msacm32" (phas As Any, ByVal had As Long, pwfxSrc As Any, pwfxDst As Any, pwfltr As Any, dwCallback As Any, dwInstance As Any, ByVal fdwOpen As Long) As Long
Private Declare Function acmStreamPrepareHeader Lib "msacm32" (ByVal has As Long, ByRef pash As ACMSTREAMHEADER, ByVal fdwPrepare As Long) As Long
Private Declare Function acmStreamReset Lib "msacm32" (ByVal has As Long, ByVal fdwReset As Long) As Long
Private Declare Function acmStreamSize Lib "msacm32" (ByVal has As Long, ByVal cbInput As Long, ByRef pdwOutputBytes As Long, ByVal fdwSize As Long) As Long
Private Declare Function acmStreamUnprepareHeader Lib "msacm32" (ByVal has As Long, ByRef pash As ACMSTREAMHEADER, ByVal fdwUnprepare As Long) As Long

Private Const OPEN_EXISTING As Long = 3
Private Const PAGE_READONLY As Long = 2&
Private Const FILE_SHARE_READ As Long = &H1
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_MAP_READ As Long = &H4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const RIFF_SIGNATURE As Long = &H46464952
Private Const WAVE_SIGNATURE As Long = &H45564157
Private Const FMT_SIGNATURE As Long = &H20746D66
Private Const DATA_SIGNATURE As Long = &H61746164
Private Const FILE_END As Long = 2
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const MPEGLAYER3_FLAG_PADDING_OFF As Long = 2
Private Const WAVE_FORMAT_MPEGLAYER3 As Long = &H55
Private Const WAVE_FORMAT_PCM As Long = 1
Private Const MPEGLAYER3_WFX_EXTRA_BYTES As Long = 12
Private Const MPEGLAYER3_ID_MPEG As Long = 1
Private Const ACM_STREAMSIZEF_SOURCE As Long = &H0
Private Const ACM_STREAMCONVERTF_BLOCKALIGN As Long = &H4

Private isMp3Init As Boolean
Private Constants As mp3Const

' // Create a sound buffer from specified audio file.
' // Support WAV, MP3
Public Function DSCreateSoundBufferFromFile(ByVal ds As DirectSound8, _
                                            ByRef strFileName As String, _
                                            ByRef bufDesc As DSBUFFERDESC) As IDirectSoundBuffer
    Dim hFile  As Long
    Dim hMap   As Long
    Dim lpData As Long
    Dim errNum As Long
    Dim size   As LARGE_INTEGER

    hFile = CreateFile(StrPtr(strFileName), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If hFile = INVALID_HANDLE_VALUE Then
        If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
            Err.Raise 53
        Else
            Err.Raise 75
        End If
        Exit Function
    End If
    ' // Get file size
    SetFilePointerEx hFile, 0, 0, size, FILE_END

    If size.highpart <> 0 Or size.lowpart < 0 Then
        Err.Raise 7
        Exit Function
    End If

    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, 0, 0, 0)
    CloseHandle hFile
    If hMap = 0 Then
        Err.Raise 5
        Exit Function
    End If

    lpData = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0)
    CloseHandle hMap
    If lpData = 0 Then
        Err.Raise 5
        Exit Function
    End If

    On Error Resume Next
    Set DSCreateSoundBufferFromFile = DSCreateSoundBufferFromMemory(ds, lpData, size.lowpart, bufDesc)
    errNum = Err.Number
    On Error GoTo 0

    UnmapViewOfFile lpData

    If errNum Then Err.Raise errNum

End Function

' // Create a sound buffer from specified audio file in memory.
' // Support WAV, MP3
Public Function DSCreateSoundBufferFromMemory(ByVal ds As DirectSound8, _
                                              ByVal lpData As Long, _
                                              ByVal szData As Long, _
                                              ByRef bufDesc As DSBUFFERDESC) As IDirectSoundBuffer
    Dim chkData As CHUNK
    Dim subChnk As CHUNK
    Dim chkType As Long
    Dim lpFmt  As Long
    Dim szFmt  As Long
    Dim lpDat  As Long
    Dim szDat  As Long
    Dim size   As Long
    Dim ptr    As Long
    Dim ret    As Long
    Dim hdr(9) As Byte

    ' // Check size
    If szData < 4 Then GoTo ERROR_OUTOFMEMORY
    ' // Check RIFF
    If IsBadReadPtr(ByVal lpData, szData) Then GoTo ERROR_OUTOFMEMORY
    GetMem4 ByVal lpData, chkData

    If chkData.id = RIFF_SIGNATURE Then
        ' // Wave
        If IsBadReadPtr(ByVal lpData, 8) Then GoTo ERROR_OUTOFMEMORY
        GetMem8 ByVal lpData, chkData
        ' // Check size
        If chkData.szData > szData Then GoTo ERROR_OUTOFMEMORY
        lpData = lpData + 8
        If IsBadReadPtr(ByVal lpData, chkData.szData) Or chkData.szData < 4 Then GoTo ERROR_OUTOFMEMORY
        GetMem4 ByVal lpData, chkType
        If chkType <> WAVE_SIGNATURE Then GoTo ERROR_FORMAT

        lpData = lpData + 4
        chkData.szData = chkData.szData - 4
        ' // Find chunks
        Do While (chkData.szData >= 8) And Not (lpFmt > 0 And lpDat > 0)

            GetMem8 ByVal lpData, subChnk
            lpData = lpData + 8

            If subChnk.szData > chkData.szData - 8 Then GoTo ERROR_OUTOFMEMORY

            Select Case subChnk.id
            Case FMT_SIGNATURE

                If lpFmt Then GoTo ERROR_FORMAT
                lpFmt = lpData
                szFmt = subChnk.szData

            Case DATA_SIGNATURE

                If lpDat Then GoTo ERROR_FORMAT
                lpDat = lpData
                szDat = subChnk.szData

            End Select
            lpData = lpData + subChnk.szData + (subChnk.szData And 1)

            chkData.szData = chkData.szData - 8 - subChnk.szData

        Loop

        If lpFmt <> 0 And lpDat <> 0 And szFmt > 0 And szDat > 0 Then

            bufDesc.dwSize = Len(bufDesc)
            bufDesc.dwBufferBytes = szDat
            bufDesc.lpwfxFormat = lpFmt

            ds.CreateSoundBuffer bufDesc, DSCreateSoundBufferFromMemory, ByVal 0&

            DSCreateSoundBufferFromMemory.Lock 0, 0, ptr, szDat, 0, 0, DSBLOCK_ENTIREBUFFER
            memcpy ByVal ptr, ByVal lpDat, szDat
            DSCreateSoundBufferFromMemory.Unlock ptr, szDat, 0, 0

        Else: GoTo ERROR_FORMAT
        End If

    Else
        ' // Expecting MP3
        If Not isMp3Init Then Mp3Init

        If szData >= 128 Then
            ' // Skip ID3V1 tag
            memcpy hdr(0), ByVal lpData + szData - 128, 3

            If hdr(0) = &H54 And hdr(1) = &H41 And hdr(2) = &H47 Then

                szData = szData - 128

            End If

        End If

        If szData >= 10 Then
            ' // Skip ID3V2 tags from beginning
            memcpy hdr(0), ByVal lpData, 10

            If hdr(0) = &H49 And hdr(1) = &H44 And hdr(2) = &H33 Then

                ' // Footer present
                If hdr(5) And &H10 Then
                    szData = szData - 10
                End If

                size = hdr(6) * &H200000
                size = size Or (hdr(7) * &H4000&)
                size = size Or (hdr(8) * &H80&)
                size = size Or hdr(9)
                size = size + 10

                lpData = lpData + size
                szData = szData - size

            Else
                ' // Skip ID3V2 tags from end
                memcpy hdr(0), ByVal lpData + szData - 10, 10

                If hdr(2) = &H49 And hdr(1) = &H44 And hdr(0) = &H33 Then

                    szData = szData - 10

                    size = hdr(6) * &H200000
                    size = size Or (hdr(7) * &H4000&)
                    size = size Or (hdr(8) * &H80&)
                    size = size Or hdr(9)
                    size = size + 10

                    szData = szData - size

                End If

            End If

        End If

        If szData < 4 Then GoTo ERROR_OUTOFMEMORY

        ' // Find a frame sync
        Do

            GetMem4 ByVal lpData, hdr(0)

            If hdr(0) = &HFF And (hdr(1) And &HE0) = &HE0 Then
                Dim vers As Long
                Dim layer As Long
                Dim bitrate As Long
                Dim smprate As Long
                Dim padding As Long
                Dim channel As Long
                Dim format As MPEGLAYER3WAVEFORMAT

                vers = (hdr(1) And &H18) \ 8
                If vers = 1 Then GoTo ERROR_FORMAT

                layer = (hdr(1) And &H6) \ 2
                If layer <> 1 Then GoTo ERROR_FORMAT    ' Only Layer 3

                If vers = 3 Then
                    bitrate = Constants.bitrate(0, (hdr(2) And &HF0) \ &H10)
                Else
                    bitrate = Constants.bitrate(1, (hdr(2) And &HF0) \ &H10)
                End If

                If vers = 3 Then
                    smprate = Constants.smprate(0, (hdr(2) And &HC) \ &H4)
                ElseIf vers = 2 Then
                    smprate = Constants.smprate(1, (hdr(2) And &HC) \ &H4)
                Else
                    smprate = Constants.smprate(2, (hdr(2) And &HC) \ &H4)
                End If

                padding = (hdr(2) And &H2) \ 2
                channel = -(((hdr(3) And &HC0) \ 64) <> 3) + 1

                If vers = 3 Then
                    size = Int(144000 * bitrate / smprate) + padding
                Else
                    size = Int(72000 * bitrate / smprate) + padding
                End If

                With format
                    .wFormatTag = WAVE_FORMAT_MPEGLAYER3
                    .cbSize = MPEGLAYER3_WFX_EXTRA_BYTES
                    .nChannels = channel
                    .nAvgBytesPerSec = bitrate * 128
                    .wBitsPerSample = 0
                    .nBlockAlign = 1
                    .nSamplesPerSec = smprate
                    .nFramesPerBlock = 1
                    .nCodecDelay = 0
                    .fdwFlags = MPEGLAYER3_FLAG_PADDING_OFF
                    .wID = MPEGLAYER3_ID_MPEG
                    .nBlockSize = size
                End With

                Exit Do

            End If

            lpData = lpData + 1
            szData = szData - 1

        Loop While szData >= 4

        If szData > 0 And format.wFormatTag = WAVE_FORMAT_MPEGLAYER3 Then
            ' // Try to convert
            Dim hStream As Long
            Dim dstFormat As WAVEFORMATEX
            Dim Buffer() As Byte
            Dim acmHdr As ACMSTREAMHEADER
            Dim outSize As Long
            Dim index As Long

            With dstFormat
                .cbSize = Len(dstFormat)
                .nChannels = format.nChannels
                .nSamplesPerSec = format.nSamplesPerSec
                .wBitsPerSample = 16
                .nBlockAlign = (.wBitsPerSample \ 8) * .nChannels
                .nAvgBytesPerSec = .nBlockAlign * .nSamplesPerSec
                .wFormatTag = WAVE_FORMAT_PCM
            End With
            ' // Open conversion stream
            ret = acmStreamOpen(hStream, 0, format, dstFormat, ByVal 0&, ByVal 0&, ByVal 0&, 0)
            If ret Then GoTo ERROR_FORMAT

            Do While szData > 0

                ' // Calc output buffer size
                ret = acmStreamSize(hStream, szData, szDat, ACM_STREAMSIZEF_SOURCE)
                If ret Then
                    acmStreamClose hStream, 0
                    GoTo ERROR_OUTOFMEMORY
                End If

                outSize = outSize + szDat

                ReDim Preserve Buffer(outSize - 1)

                ' // Calc header
                With acmHdr
                    .cbStruct = Len(acmHdr)
                    .lppbDst = VarPtr(Buffer(index))
                    .lppbSrc = lpData
                    .cbDstLength = szDat
                    .cbSrcLength = szData
                End With

                ' // Prepare header
                ret = acmStreamPrepareHeader(hStream, acmHdr, 0)
                If ret Then
                    acmStreamClose hStream, 0
                    GoTo ERROR_OUTOFMEMORY
                End If
                ' // Convert to PCM
                ret = acmStreamConvert(hStream, acmHdr, ACM_STREAMCONVERTF_BLOCKALIGN)
                acmStreamUnprepareHeader hStream, acmHdr, 0

                If ret Then
                    acmStreamClose hStream, 0
                    GoTo ERROR_OUTOFMEMORY
                End If

                szData = szData - acmHdr.cbSrcLengthUsed
                lpData = lpData + acmHdr.cbSrcLengthUsed
                index = index + acmHdr.cbDstLengthUsed

            Loop

            acmStreamClose hStream, 0

            outSize = index

            bufDesc.dwSize = Len(bufDesc)
            bufDesc.dwBufferBytes = outSize
            bufDesc.lpwfxFormat = VarPtr(dstFormat)

            ds.CreateSoundBuffer bufDesc, DSCreateSoundBufferFromMemory, ByVal 0&

            DSCreateSoundBufferFromMemory.Lock 0, 0, ptr, outSize, 0, 0, DSBLOCK_ENTIREBUFFER
            memcpy ByVal ptr, Buffer(0), outSize
            DSCreateSoundBufferFromMemory.Unlock ptr, outSize, 0, 0

        Else: GoTo ERROR_FORMAT
        End If

    End If

    Exit Function

ERROR_OUTOFMEMORY:
    Err.Raise 7: Exit Function
ERROR_FORMAT:
    Err.Raise 5: Exit Function

End Function

Private Sub Mp3Init()
    Dim b      As curBuffer

    b.b(0) = 450377142658.6656@: b.b(1) = 900743977448.248@: b.b(2) = 1351114248211.6672@
    b.b(3) = 1801487954948.9248@: b.b(4) = 2702228496423.3344@: b.b(5) = 3602975909897.8496@
    b.b(6) = 4503737067267.712@: b.b(7) = 18941235272.0895@: b.b(8) = 4735201446.045@
    b.b(9) = 10307921515.2@: b.b(10) = 13743895348.4@: b.b(11) = 3435973838.4@

    memcpy Constants.bitrate(0, 1), b.b(0), 96

    isMp3Init = True

End Sub
