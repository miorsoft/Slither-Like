Attribute VB_Name = "modFOOD"
Option Explicit

Public Type tPosAndVel
    POS           As geoVector2D
    Vel           As geoVector2D

End Type

'Public FOOD() As tPosAndVel
'Public FOODcolor() As Long
'Public FoodAge() As Double
'Public FootRotSign() As Long


Public Type tFOOD
    POS           As geoVector2D
    Vel           As geoVector2D
    Color         As Long
    foodWHITE     As Double
    RotSign       As Long
    Born          As Long
    fromSnake     As Long
    fromSnakeIconName As String
End Type

Public FOOD()     As tFOOD


Public NFood      As Long
Private MaxFood   As Long
Public FOODLEVEL  As Long

Public FoodDiv    As Double


Public Const FoodSize As Double = 9
Public Const FoodXSnake As Long = 35    '34 ' 30 '25

Public Const FoodLengthValue As Double = 1

Public COStable(-360 To 360) As Double
Public SINtable(-360 To 360) As Double




Public Sub InitFOOD(HowMuch As Long)
    Dim I         As Long




    Dim P         As geoVector2D
    For I = 1 To HowMuch
        P.x = wMinX + RndM * (wMaxX - wMinX)
        P.y = wMinY + RndM * (wMaxY - wMinY)
        AddFoodParticle P, False, -1
    Next



    For I = -360 To 360
        COStable(I) = Cos(I / 360 * PI2)    '* 0.002
        SINtable(I) = Sin(I / 360 * PI2)    '* 0.002
    Next


    MaxFood = HowMuch
    FOODLEVEL = HowMuch
    FoodDiv = 1 / (MaxFood - MinFoodForLevelChange)


    Set MAPsrf = Cairo.CreateSurface((wMaxX - wMinX) * MapScale, (wMaxY - wMinY) * MapScale, ImageSurface)
    Set MapCC = MAPsrf.CreateContext
    MapCC.AntiAlias = CAIRO_ANTIALIAS_FAST
    MapCC.SetLineCap CAIRO_LINE_CAP_ROUND
    MapCC.SetLineJoin CAIRO_LINE_JOIN_ROUND
    

End Sub
Private Sub RemoveFood(wF As Long)
    Dim I         As Long

    '    NFood = NFood - 1
    '    For I = wF To NFood
    '        FOOD(I) = FOOD(I + 1)
    '        FoodAge(I) = FoodAge(I + 1)
    '    Next

    '---With No Loop '--2024
    FOOD(wF) = FOOD(NFood)
'    FoodAge(wF) = FoodAge(NFood)
    NFood = NFood - 1




End Sub

Public Sub AddFoodParticle(POS As geoVector2D, ByVal IsWhite As Boolean, ByVal fromSnake As Long)

    NFood = NFood + 1
    If NFood > MaxFood Then
        MaxFood = NFood + 20
        ReDim Preserve FOOD(MaxFood)
        '        ReDim Preserve FoodAge(MaxFood)
        '        ReDim Preserve FootRotSign(MaxFood)
    End If

    With FOOD(NFood)
        .Vel.x = 0
        .Vel.y = 0
        .POS = POS

        .RotSign = Int(RndM * 2)
        If .RotSign = 0 Then .RotSign = -1

        If IsWhite Then .foodWHITE = 1 Else: .foodWHITE = 0
        .Born = CNT
        .fromSnake = fromSnake
        .fromSnakeIconName = "FoodIcon" & CStr(fromSnake)
                

        InitFoodIcon .fromSnake, NFood
        
    End With
End Sub


'Private Sub FoodToRNDPosition(wF As Long)
'    With FOOD(wF)
'        .POS.x = wMinX + RndM * (wMaxX - wMinX)
'        .POS.y = wMinY + RndM * (wMaxY - wMinY)
'    End With
'End Sub
Public Sub DrawFOOD()
    Dim I      As Long

    For I = 0 To NFood
        With FOOD(I)
            If InsideBB(CameraBB, .POS) Then
'                vbDrawCC.RenderSurfaceContent "FoodIcon", .POS.x - FoodSize, .POS.y - FoodSize, , , CAIRO_FILTER_FAST, 0.75
                      vbDrawCC.RenderSurfaceContent .fromSnakeIconName, .POS.x - FoodSize, .POS.y - FoodSize, , , CAIRO_FILTER_FAST, 0.75
           
                If .foodWHITE > 0 Then vbDrawCC.RenderSurfaceContent "FoodIconLight", .POS.x - FoodSize * 2, .POS.y - FoodSize * 2, , , CAIRO_FILTER_FAST, .foodWHITE
            End If

'            .foodWHITE = .foodWHITE - 0.0015    '--2024
'            If .foodWHITE < 0# Then .foodWHITE = 0#
'            .foodWHITE = .foodWHITE * 0.996
            .foodWHITE = .foodWHITE * 0.9965
            
        End With
    Next

End Sub

Public Sub FoodMoveAndCheckEaten()
    Dim I      As Long
    Dim J      As Double
    Dim HeadPosition As geoVector2D
    Dim Hvel   As geoVector2D

    Dim D      As Double
    Dim dx     As Double
    Dim dy     As Double
    Dim vD     As Double
    Dim vDx    As Double
    Dim vDy    As Double


    Dim GrabR  As Double



    'For I = 0 To NFood
    Do

        With FOOD(I)

            For J = 0 To NSnakes
                If InsideBB(Snake(J).getBB, FOOD(I).POS) Then    'when there's a lot of food skip check far away
                    HeadPosition = Snake(J).GetHEADPos


                    'HeadPosition = VectorSUM(HeadPosition, VectorMUL(Snake(J).GetHEADVel, Snake(J).MySIZE * 3))

                    dx = HeadPosition.x - .POS.x
                    dy = HeadPosition.y - .POS.y
                    D = dx * dx + dy * dy

                    'GrabR = Snake(J).MySIZE * 10
                    GrabR = Snake(J).Diam

                    GrabR = GrabR * GrabR
                    If D < GrabR Then
                        vD = 0.01 * Sqr(GrabR) / (Sqr(D) * 1)
                        vDx = dx * vD
                        vDy = dy * vD
                        .Vel.x = .Vel.x + vDx
                        .Vel.y = .Vel.y + vDy
                        Snake(J).TongueOut = 1
                    End If


                    'GrabR = Snake(J).DIAM * 0.7
                    GrabR = Snake(J).Radius + FoodSize * 0.5
                    GrabR = GrabR * GrabR
                    If D < GrabR Then
                        If J = SNAKECAMERA Then
                            If Snake(SNAKECAMERA).IsDying = 0 Then
                                ' MultipleSounds.playsound "eatfruit.wav"

                                MultipleSounds.PlaySound SoundPlayerChomp, 0, -1500
                                PlayerScore = PlayerScore + 10

                            End If
                        Else
                            HeadPosition = Snake(SNAKECAMERA).GetHEADPos
                            dx = HeadPosition.x - .POS.x
                            dy = HeadPosition.y - .POS.y
                            D = Sqr(dx * dx + dy * dy)
                            'MultipleSounds.PlaySound SoundEnemyChomp, ClampLong(-dx * 3, -10000, 10000), ClampLong(-D * 0.8, -10000, 0)
                            MultipleSounds.PlaySound SoundEnemyChomp, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 1, -10000, 0)

                        End If
                        'Snake(J).fLength = Snake(J).fLength + 1
                        Snake(J).SetSize = Snake(J).GetSize + FoodLengthValue
                        '   FoodToRNDPosition I
                        RemoveFood I
                    End If
                End If

            Next


            .POS = VectorSUM(.POS, .Vel)
            .Vel = VectorMUL(.Vel, 0.992)

            '---- Food move/animation  '--2024
            Dim SC As geoVector2D
            Dim A As Long
            A = (I * 137.52 + CNT * 1.33 * .RotSign) Mod 360
            SC.x = COStable(A) * 0.0027 '0.0025
            SC.y = SINtable(A) * 0.0027 '0.0025
            .Vel = VectorSUM(.Vel, SC)
            '------------


            If .POS.x < wMinX Then .POS.x = wMinX: .Vel.x = -.Vel.x
            If .POS.y < wMinY Then .POS.y = wMinY: .Vel.y = -.Vel.y
            If .POS.x > wMaxX Then .POS.x = wMaxX: .Vel.x = -.Vel.x
            If .POS.y > wMaxY Then .POS.y = wMaxY: .Vel.y = -.Vel.y


        End With

        '    Next
I = I + 1
Loop While I <= NFood

End Sub

Public Sub CreateFoodFromDeadSnake(wS As Long)
    Dim I         As Long



    '    MinFoodForLevelChange = MinFoodForLevelChange + (Snake(wS).Ntokens - 1) * 0.5 - (STARTLENGTH - 1)
    '    FoodDiv = 1 / (MaxFood - MinFoodForLevelChange)

    For I = 0 To Snake(wS).Ntokens '- 2
        '
        '        NFood = NFood + 1
        '        If NFood > MaxFood Then
        '            MaxFood = NFood + 20
        '            ReDim Preserve FOOD(MaxFood)
        '            ReDim Preserve FoodAge(MaxFood)
        '        End If
        '        With FOOD(NFood)
        '            .POS = Snake(wS).GetTokenPos(I)
        '            .Vel.x = (RndM * 2 - 1) * 0.125
        '            .Vel.y = (RndM * 2 - 1) * 0.125
        '        End With
        '        FoodAge(NFood) = 1

        AddFoodParticle Snake(wS).GetTokenPos(I), True, wS '-1
        With FOOD(NFood)
            .Vel.x = (RndM * 2 - 1) * 0.125
            .Vel.y = (RndM * 2 - 1) * 0.125
        End With

    Next
End Sub


Public Function PointToNearestFood(Head As tPosAndVel, ByVal SnakeIDX As Long) As geoVector2D
    Dim I      As Long
    Dim J      As Long
    Dim D      As Double
    Dim dx     As Double
    Dim dy     As Double
    Dim MIND   As Double
    Dim Direct As Double
    Dim Avoid  As Double

    Dim HX#, HY#, HVX#, HVY#
    HX = Head.POS.x
    HY = Head.POS.y
    HVX = Head.Vel.x
    HVY = Head.Vel.y
    Dim SFFM#
    SFFM = Snake(SnakeIDX).SearchForFOODMODE

    MIND = 1E+32

    For I = 0 To NFood
        With FOOD(I)
            dx = .POS.x - HX
            dy = .POS.y - HY
            D = dx * dx + dy * dy

            Direct = -Sgn(dx * HVX + dy * HVY)    '''' Consider nearer the ones in front

            ''        Direct = Direct + 2#
            '        Direct = Direct + 1.65 '--2024

            Direct = Direct + SFFM          '--2024

            ''D = D * Direct * (1# - FoodAge(I) * 0.95)
            'D = D * Direct * (1# - FoodAge(I) * 0.98)
            D = D * Direct * (1# - .foodWHITE * 0.99999)    '--2024


            If .fromSnake = SnakeIDX Then 'Ignore fresh-Trail ones
                Avoid = 300 - (CNT - .Born) '300=4 secs .. 200
                If Avoid > 0 Then
                    D = D * (1 + 5 * Avoid * 0.005)
                End If
            End If



            If D < MIND Then
                MIND = D
                J = I
            End If
        End With
    Next

    PointToNearestFood = VectorNormalize(VectorSUB(FOOD(J).POS, Head.POS))

End Function



