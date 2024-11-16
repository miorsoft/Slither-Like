Attribute VB_Name = "modMain"
Option Explicit

Private Scores() As Long
Private ScoresIdx() As Long



Public MultipleSounds As cSounds

Public Snake() As clsSnake
Public NSnakes As Long
Public MaxNSnakes As Long
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

Public DoDrawFlags As Long
Public DoDrawBB As Long
Public DoDrawMAP As Long
Public SaveFrames As Long
Private Frame As Long
Private Const JPGframeRate As Long = 3    ''''75/3= 25 FPS ' Multiple of 3  ( cnt mod 3)

Public DoBackGround As Long


Public ZOOM As Double
Public invZOOM As Double

Public AIcontrol As Boolean

Private StrScore As String
Private invMaxScore As Double

Public LIFES As Long
Public PlayerScore As Long
Public LevelCNT As Long

Public SNAKECAMERA As Long

Public MAPsrf As cCairoSurface
Public MapCC As cCairoContext
Public Const MapScale As Double = 0.05

'********************
'***** MANIFEST *****
'********************
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

Private Sub Main()

    Dim iccex As InitCommonControlsExStruct, hMod As Long
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all known values
    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_ALL_CLASSES    ' you really should customize this value from the available constants
    End With
    On Error Resume Next ' error? Requires IEv3 or above
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    On Error GoTo 0
    '... show your main form next (i.e., Form1.Show)
    fMain.Show
    If hMod Then FreeLibrary hMod


'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.

End Sub
'********************
'********************
'********************

Public Sub InitPool(ByVal NoSnakes As Long, Optional NEWGAME As Boolean = True)

    Dim I      As Long

    If NEWGAME Then
        LIFES = 20 ' 5 '3
        PlayerScore = STARTLENGTH * 10
        Level = 1
        NoSnakes = 6
    Else    'BONUS LIfE
        If ScoresIdx(0) = PLAYER Then LIFES = LIFES + 1
    End If

    LevelCNT = CNT + 300

    If Level = 1 Then Print #1, "--------------------"
    Print #1, Level & "  Lifes " & LIFES & "   Score " & PlayerScore


    NSnakes = NoSnakes
    InvNSnakes = 1 / NSnakes
    MinFoodForLevelChange = NSnakes * (FoodXSnake / 34) * 2 '2

    If NSnakes > MaxNSnakes Then
        ReDim Snake(NSnakes)
        ReDim Scores(NSnakes)
        ReDim ScoresIdx(NSnakes)
        MaxNSnakes = NSnakes
    End If

    For I = 0 To NSnakes
        If Snake(I) Is Nothing Then Set Snake(I) = New clsSnake
        Snake(I).Init RndM * MaxW, RndM * maxH, I, STARTLENGTH    '+ rndm * 30
    Next

If SNAKECAMERA > NoSnakes Then fMain.hSnake.Value = 0
fMain.hSnake.Min = 0
fMain.hSnake.maX = NoSnakes


End Sub





Public Sub MainLoop()
    Dim I         As Long
    Dim pTime     As Double
    Dim pTime2    As Double
    Dim FPS       As Long
    Dim pCnt      As Long
    Dim J         As Long

    Dim StrCaption As String

    Dim ZOOMtoGO  As Double

    Dim FH        As Double


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
                Snake(I).MOVE2
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

                    ZOOMtoGO = 0.05 + 10 * Snake(SNAKECAMERA).InvDiam ^ 0.7


                    'ZOOM = ZOOM * 0.98 + ZOOMtoGO * 0.02
                    ZOOM = ZOOM * 0.995 + ZOOMtoGO * 0.005    '--2024

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



                    '                    Camera = Snake(PLAYER).GetHEADPos
                    Camera = Snake(SNAKECAMERA).GetHEADPos


                    '                    CameraBB.minX = Camera.x - CenX
                    '                    CameraBB.maxX = Camera.x + CenX
                    '                    CameraBB.minY = Camera.y - CenY
                    '                    CameraBB.maxY = Camera.y + CenY

                    CameraBB.minX = Camera.x - CenX * invZOOM
                    CameraBB.maxX = Camera.x + CenX * invZOOM
                    CameraBB.minY = Camera.y - CenY * invZOOM
                    CameraBB.maxY = Camera.y + CenY * invZOOM


                    For I = 0 To NSnakes
                        Snake(I).DRAW DoDrawBB
                    Next
                    If DoDrawFlags Then
                        For I = 0 To NSnakes
                            Snake(I).DRAWFlag
                        Next
                    End If

                    .Restore


                    If CNT < LevelCNT Then .DrawTextCell CenX - 50, CenY - 30, 100, 60, "LEVEL " & CStr(Level) & vbCrLf & vbCrLf & "Lifes " & LIFES, , vbCenter, , 444, 0.9, 111

                    .TextOut 5, 5, StrCaption
                    .DrawText MaxW - 300, 5, 400, 1000, StrScore

                    For I = 0 To NSnakes
                        J = ScoresIdx(I)
                        .SetSourceRGBA Snake(J).ColorR, Snake(J).ColorG, Snake(J).ColorB, 0.5
                        .Rectangle MaxW - 305, 5 + (I + 2) * 15, 90 * Scores(I) * invMaxScore, 14
                        .Fill
                        If (CNT - Snake(J).DyingTime) < 160 Then
                            .SetLineWidth 2
                            .SetSourceRGB 1, 0.25, 0.25
                            .Rectangle MaxW - 305, 5 + (I + 2) * 15, 90, 14
                            .Stroke
                        End If

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


If DoDrawMAP Then
If (CNT Mod 6&) = 0& Then UpdateMAPsrf
vbDrawCC.RenderSurfaceContent MAPsrf, 10, 25
End If

                vbDRAW.Srf.DrawToDC PicHDC



                If SaveFrames Then
                    ' If CNT Mod JPGframeRate = 0 Then    'Multiple of 4 JPGframeRate
                    If DoLOOP Then
                        vbDRAW.Srf.WriteContentToPngFile App.Path & "\Frames\" & format(Frame, "00000") & ".png"    ', 100
                        Frame = Frame + 1
                    End If
                    ' End If
                End If

            End If


            DoEvents
            CNT = CNT + 1


            If NFood <= MinFoodForLevelChange Then    '5 'Next Level '
                InitPool NSnakes + 1, False    ' * 1.18    '1.2
                InitFOOD NSnakes * FoodXSnake
                Level = Level + 1
                StrCaption = " Lifes: " & LIFES & "     SCORE: " & PlayerScore & "     Level: " & Level & "     Snakes: " & NSnakes + 1 & "     Food: " & NFood & "      FPS: " & FPS \ JPGframeRate & "     Length: " & Snake(PLAYER).GetSize & "                                   By MiorSoft"
                MultipleSounds.PlaySound SoundINTRO
            End If



            If (CNT And 15&) = 0& Then
                StrCaption = " Lifes: " & LIFES & "     SCORE: " & PlayerScore & "     Level: " & Level & "     Snakes: " & NSnakes + 1 & "     Food: " & NFood & "      FPS: " & FPS \ JPGframeRate & "     Length: " & Snake(PLAYER).GetSize & "                                   By MiorSoft"
                UpdateSCORESString
                
            End If
           


        End If

    Loop While DoLOOP

End Sub

Public Function ClampLong(ByVal V As Double, ByVal Min As Long, ByVal maX As Long) As Long

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

Debug.Print
AG: '- SORT SCORES------------------
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
    Debug.Print SW
'    If SW Then GoTo AG


    StrScore = "Level SCORES:" & vbCrLf & vbCrLf
    For I = 0 To NSnakes
        If ScoresIdx(I) = 0& Then
            S = "PLYR"
        Else
            S = ScoresIdx(I)
        End If
        StrScore = StrScore & S & Space(11 - Len(S) - Len(CStr(Scores(I)))) & Scores(I) & vbCrLf
    Next



    invMaxScore = 1# / Scores(0)



End Sub

Private Function UpdateMAPsrf()
    Dim I         As Long
    Dim P         As geoVector2D

    '    MapCC.Operator = CAIRO_OPERATOR_OUT
    '    MapCC.SetSourceRGBA 0, 0, 0, 0.8
    '    MapCC.Paint
    '    MapCC.Operator = CAIRO_OPERATOR_OVER

    Dim O         As Long
    O = MapCC.Operator
    MapCC.Operator = CAIRO_OPERATOR_CLEAR: MapCC.Paint: MapCC.Operator = O
    
    'BackGround
    MapCC.SetSourceRGBA 0.28, 0.28, 0.28, 0.82
    'MapCC.Paint
    MapCC.RoundedRect 0, 0, MAPsrf.Width, MAPsrf.Height, 6: MapCC.Fill
    
    'FOOD
    MapCC.SetSourceRGBA 0.1, 0.6, 0.1, 0.55
    For I = 0 To NFood
        MapCC.Rectangle (FOOD(I).POS.x - wMinX) * MapScale, (FOOD(I).POS.y - wMinY) * MapScale, 1, 1
    Next
    MapCC.Stroke

    'SNAKES
    For I = NSnakes To 0 Step -1
        Snake(I).DrawToMAP MapCC, MapScale
    Next

    'Visible Screen
    P = Snake(PLAYER).GetHEADPos
    MapCC.SetSourceRGBA 1, 1, 1, 0.06
    MapCC.RoundedRect (P.x - wMinX) * MapScale - MaxW * 0.5 * invZOOM * MapScale, (P.y - wMinY) * MapScale - maxH * 0.5 * invZOOM * MapScale, MaxW * MapScale * invZOOM, maxH * MapScale * invZOOM, 4
    MapCC.Fill

End Function
