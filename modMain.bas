Attribute VB_Name = "modMain"
Option Explicit

Private Scores() As Long
Private ScoresIdx() As Long



Public MultipleSounds As cSounds

Public Snake() As clsSnake
Public NSnakes As Long
Public MinFoodForLevelChange As Long

Public InvNSnakes As Double

Public DoLOOP As Boolean

Public MousePos As geoVector2D

Public CNT  As Long

Public Level As Long

Public Camera As geoVector2D
Public CameraBB As tBB


Public Const PLAYER As Long = 0

Public Const STARTLENGTH As Long = 5 '



Public Const SoundINTRO As String = "intropm.wav"
Public Const SoundPlayerChomp As String = "wakawaka.wav"
Public Const SoundEnemyChomp As String = "apple-crunch-17.wav"

'Public Const SoundEnenmyKilledByMe As String = "eatghost.wav"
'Public Const SoundEnenmyKilledByMe As String = "reitanna__son-of-a-bitch.wav"
Public Const SoundEnenmyKilledByMe As String = "manuts__death-5.wav"

Public Const SoundEnenmyKilled As String = "uohm.wav"


Public Const SoundPlayerDeath As String = "death.wav"

Public DrawBB As Long
Public SaveFrames As Long
Private Frame As Long
Private Const JPGframeRate As Long = 3    ''''75/3= 25 FPS ' Multiple of 3  ( cnt mod 3)

Public DoBackGround As Long


Public ZOOM As Double
Public invZOOM As Double

Public AIcontrol As Boolean

Private StrScore As String
Private invMaxScore As Double


Public Sub InitPool(ByVal NoSnakes As Long)

    Dim I   As Long
    NSnakes = NoSnakes
    InvNSnakes = 1 / NSnakes
    MinFoodForLevelChange = NSnakes * 2
    

    ReDim Snake(NSnakes)
    ReDim Scores(NSnakes)
    ReDim ScoresIdx(NSnakes)

    For I = 0 To NSnakes
        If Snake(I) Is Nothing Then Set Snake(I) = New clsSnake
        Snake(I).Init Rnd * MaxW, Rnd * maxH, I, STARTLENGTH    '+ Rnd * 30
    Next


End Sub





Public Sub MainLoop()
    Dim I         As Long
    Dim pTime     As Double
    Dim pTime2    As Double
    Dim FPS       As Long
    Dim pCnt      As Long
Dim J As Long

    Dim StrCaption As String



    Dim ZOOMtoGO  As Double



    DoLOOP = True

    Level = 1
    'fMain.Caption = "Level: " & Level & "  Snakes: " & NSnakes & "  Food: " & NFood


    MultipleSounds.PlaySound SoundINTRO

    Timing = 0
    pTime = Timing
    pTime2 = Timing




    Do

        If Timing - pTime2 > 1 Then
            FPS = CNT - pCnt
            pCnt = CNT
            pTime2 = Timing
        End If




        If Timing - pTime > 0.01333 Then    '75 FPS computed
            'If Timing - pTime > 0.01666 Then    '60 FPS computed

            pTime = Timing

            For I = 0 To NSnakes
                Snake(I).MOVE
            Next

            FoodMoveAndCheckEaten       '------------------------------------


            If CNT Mod JPGframeRate = 0 Then    '' 75 / 3 FPS =25 FPS Drawn
                'If CNT Mod 3 = 0 Then    '' 60 / 3 FPS =20 FPS Drawn

                '                CheckCollisionsOnlyPlayer
                CheckCollisionsALLtoALL

                With vbDrawCC
                    .SetSourceColor 0
                    .Paint
                    .Save

                    'ZOOMtoGO = 30# * Snake(PLAYER).InvDiam
                    'ZOOMtoGO = 28# * Snake(PLAYER).InvDiam '---2nd video
                    'ZOOMtoGO = 0.0625 + 25# * Snake(PLAYER + 1).InvDiam


                    'ZOOMtoGO = 0.0625 + 25# * Snake(PLAYER).InvDiam   'ok

                    ZOOMtoGO = 0.05 + 10 * Snake(PLAYER).InvDiam ^ 0.7


                    ZOOM = ZOOM * 0.98 + ZOOMtoGO * 0.02
                    invZOOM = 1# / ZOOM

                    .TranslateDrawings -Camera.x * ZOOM + CenX, -Camera.y * ZOOM + CenY

                    .ScaleDrawings ZOOM, ZOOM

                    '-...................................................
                    If DoBackGround Then
                        ' USE BACKGOUND --->>> Slow with ZOOM
                        '.RenderSurfaceContent "BK", wMinX, wMinY, , , CAIRO_FILTER_FAST

                        'Lower Res
                        .RenderSurfaceContent "BK", wMinX, wMinY, (wMaxX - wMinX), (wMaxY - wMinY), CAIRO_FILTER_FAST

                        '                        .Rectangle wMinX, wMinY, wMaxX - wMinX, wMaxY - wMinY
                        '                        .Fill True, Cairo.cr
                    Else
                        .SetSourceColor &H404040
                        .Rectangle wMinX, wMinY, wMaxX - wMinX, wMaxY - wMinY
                        .Fill
                    End If



                    DrawFOOD            '--------------------------------


                    ' Camera = Snake(PLAYER + 1).GetHEADPos
                    Camera = Snake(PLAYER).GetHEADPos

                    '                    CameraBB.minX = Camera.x - CenX
                    '                    CameraBB.maxX = Camera.x + CenX
                    '                    CameraBB.minY = Camera.y - CenY
                    '                    CameraBB.maxY = Camera.y + CenY

                    CameraBB.minX = Camera.x - CenX * invZOOM
                    CameraBB.maxX = Camera.x + CenX * invZOOM
                    CameraBB.minY = Camera.y - CenY * invZOOM
                    CameraBB.maxY = Camera.y + CenY * invZOOM




                    For I = 0 To NSnakes
                        Snake(I).DRAW DrawBB
                    Next

                    .Restore

                    .TextOut 5, 5, StrCaption
                    .DrawText MaxW - 300, 5, 400, 1000, StrScore

For I = 0 To NSnakes
J = ScoresIdx(I)
.SetSourceRGBA Snake(J).ColorR, Snake(J).ColorG, Snake(J).ColorB, 0.5
.Rectangle MaxW - 305, 5 + (I + 2) * 15, 90 * Scores(I) * invMaxScore, 14
.Fill
Next
.SetSourceRGBA 0, 1, 0, 0.3
.Rectangle MaxW - 305, 2, 90 * (1 - (NFood - MinFoodForLevelChange) * FoodDiv), 31
.Fill





                    If SaveFrames Then  'Recorder Red DOT
                        .SetSourceRGBA 1, 0, 0, (1# + Sin(CNT * 0.01333 * PI2))
                        .Ellipse MaxW - 20, 30, 18, 18
                        .Fill
                    End If


                End With


                vbDRAW.Srf.DrawToDC PicHDC



                If SaveFrames Then
                    ' If CNT Mod JPGframeRate = 0 Then    'Multiple of 4 JPGframeRate
                    If DoLOOP Then
                        vbDRAW.Srf.WriteContentToJpgFile App.Path & "\Frames\" & format(Frame, "00000") & ".jpg", 100
                        Frame = Frame + 1
                    End If
                    ' End If
                End If

            End If


            DoEvents
            CNT = CNT + 1


            If NFood <= MinFoodForLevelChange Then  '5 'Next Level '
                InitPool NSnakes * 1.18    '1.2
                InitFOOD NSnakes * FoodXSanke
                Level = Level + 1
                StrCaption = "Level: " & Level & "       Snakes: " & NSnakes & "       Food: " & NFood & "        FPS: " & FPS \ JPGframeRate & "       Score: " & Snake(PLAYER).GetSize & "                                   By MiorSoft"
                MultipleSounds.PlaySound SoundINTRO
            End If


            If CNT Mod 100 = 0 Then
                StrCaption = " Level: " & Level & "       Snakes: " & NSnakes & "       Food: " & NFood & "        FPS: " & FPS \ JPGframeRate & "       Score: " & Snake(PLAYER).GetSize & "                                   By MiorSoft"
                UpdateSCORESString
            End If



        End If

    Loop While DoLOOP

End Sub

Public Function ClampLong(V As Double, Min As Long, maX As Long) As Long

    ClampLong = V
    If V < Min Then
        ClampLong = Min
    ElseIf V > maX Then
        ClampLong = maX
    End If


End Function

Private Sub UpdateSCORESString()
    Dim I As Long, J As Long
    Dim SW        As Long
    Dim TMP&
    Dim S         As String


    For I = 0 To NSnakes
        Scores(I) = Snake(I).GetSize * 10
        ScoresIdx(I) = I
    Next


AG: '- SORT------------------
    SW = 0
    For I = 0 To NSnakes - 1
        For J = I + 1 To NSnakes
            If Scores(I) < Scores(J) Then    'SWAP
                TMP = Scores(I): Scores(I) = Scores(J): Scores(J) = TMP
                TMP = ScoresIdx(I): ScoresIdx(I) = ScoresIdx(J): ScoresIdx(J) = TMP
                SW = SW + 1
            End If
        Next
    Next
    If SW Then GoTo AG


    StrScore = "SCORES:" & vbCrLf & vbCrLf
    For I = 0 To NSnakes
        If ScoresIdx(I) = 0& Then
            S = "PLYR"
        Else
            S = ScoresIdx(I)
        End If
        StrScore = StrScore & S & Space(11 - Len(S) - Len(CStr(Scores(I)))) & Scores(I) & vbCrLf
    Next



invMaxScore = 1 / Scores(0)



End Sub

