Attribute VB_Name = "modMain"
Option Explicit

Public MultipleSounds As clsSounds

Public Snake() As clsSnake
Public NSnakes As Long

Public DoLOOP As Boolean

Public MousePos As geoVector2D

Public CNT  As Long

Public Level As Long

Public Camera As geoVector2D
Public CameraBB As tBB


Public Const PLAYER As Long = 0

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


Public Sub InitPool(ByVal NoSnakes As Long)

    Dim I   As Long
    NSnakes = NoSnakes

    ReDim Snake(NSnakes)

    For I = 0 To NSnakes
        If Snake(I) Is Nothing Then Set Snake(I) = New clsSnake
        Snake(I).Init Rnd * MaxW, Rnd * maxH, I, 5    '+ Rnd * 30
    Next


End Sub





Public Sub MainLoop()
    Dim I   As Long
    Dim pTime As Double



    Dim ZOOMtoGO As Double



    DoLOOP = True

    Level = 1
    fMain.Caption = "Level: " & Level & "  Snakes: " & NSnakes & "  Food: " & NFood
    MultipleSounds.PlaySound SoundINTRO

    Timing = 0
    pTime = Timing

    vbDRAW.CC.AntiAlias = CAIRO_ANTIALIAS_GRAY




    Do

        If Timing - pTime > 0.01333 Then    '75 FPS computed
            pTime = Timing

            For I = 0 To NSnakes
                Snake(I).MOVE
            Next

            FoodMoveAndCheckEaten    '------------------------------------


            If CNT Mod 3 = 0 Then    '' 75 / 3 FPS =25 FPS Drawn

                '                CheckCollisionsOnlyPlayer
                CheckCollisionsALLtoALL

                With vbDRAW.CC
                    .SetSourceColor 0
                    .Paint
                    .Save

                    ZOOMtoGO = 30# / Snake(PLAYER).Diam
                    ZOOM = ZOOM * 0.98 + ZOOMtoGO * 0.02
                    invZOOM = 1# / ZOOM

                    .TranslateDrawings -Camera.x * ZOOM + CenX, -Camera.y * ZOOM + CenY


                    .ScaleDrawings ZOOM, ZOOM

                    '-...................................................
                    If DoBackGround Then
                        ' USE BACKGOUND --->>> Slow with ZOOM
                        .RenderSurfaceContent "BK", wMinX, wMinY, , , CAIRO_FILTER_FAST
                        '                        .Rectangle wMinX, wMinY, wMaxX - wMinX, wMaxY - wMinY
                        '                        .Fill True, Cairo.cr
                    Else
                        .SetSourceColor &H404040
                        .Rectangle wMinX, wMinY, wMaxX - wMinX, wMaxY - wMinY
                        .Fill
                    End If



                    DrawFOOD    '--------------------------------


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
                End With


                vbDRAW.Srf.DrawToDC PicHDC



                If SaveFrames Then


                    If CNT Mod JPGframeRate = 0 Then    'Multiple of 3 JPGframeRate
                        If DoLOOP Then
                            vbDRAW.Srf.WriteContentToJpgFile App.Path & "\Frames\" & format(Frame, "00000") & ".jpg", 100
                            Frame = Frame + 1
                        End If
                    End If
                End If

            End If


            DoEvents
            CNT = CNT + 1

            If NFood < 5 Then   'Next Level
                InitPool NSnakes * 1.2
                InitFOOD NSnakes * 20
                Level = Level + 1
                fMain.Caption = "Level: " & Level & "  Snakes: " & NSnakes & "  Food: " & NFood
                MultipleSounds.PlaySound SoundINTRO
            End If


            If CNT Mod 100 = 0 Then
                fMain.Caption = "Level: " & Level & "  Snakes: " & NSnakes & "  Food: " & NFood
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
