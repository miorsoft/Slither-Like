Attribute VB_Name = "modFOOD"
Option Explicit

Public Type tPosAndVel
    POS     As geoVector2D
    Vel     As geoVector2D
    RotSign As Long
End Type

Public FOOD() As tPosAndVel
Public FOODcolor() As Long
Public FoodAge() As Double


Public NFood As Long
Public MaxFood As Long


Public Const FoodSize As Double = 9

Public Const FoodLengthValue As Double = 1

Public COStable(-360 To 360) As Double
Public SINtable(-360 To 360) As Double


Public Sub InitFOOD(HowMuch As Long)
    Dim I         As Long

    NFood = HowMuch
    MaxFood = NFood

    ReDim FOOD(NFood)
    ReDim FoodAge(NFood)

    For I = 0 To NFood

        With FOOD(I)
            .POS.x = wMinX + Rnd * (wMaxX - wMinX)
            .POS.y = wMinY + Rnd * (wMaxY - wMinY)

            FoodAge(I) = 0

            .RotSign = Int(Rnd * 2)
            If .RotSign = 0 Then .RotSign = -1    '1

        End With
    Next



    For I = -360 To 360
        COStable(I) = Cos(I / 360 * PI2)    '* 0.002
        SINtable(I) = Sin(I / 360 * PI2)    '* 0.002
    Next


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
    FoodAge(wF) = FoodAge(NFood)
    NFood = NFood - 1




End Sub

Public Sub AddFoodParticle(POS As geoVector2D, IsWhite As Long)

    NFood = NFood + 1
    If NFood > MaxFood Then
        MaxFood = NFood + 20

        ReDim Preserve FOOD(MaxFood)
        ReDim Preserve FoodAge(MaxFood)
    End If

    With FOOD(NFood)
        .Vel.x = 0
        .Vel.y = 0
        .POS = POS
                    .RotSign = Int(Rnd * 2)
            If .RotSign = 0 Then .RotSign = -1    '1
        
    End With

If IsWhite Then FoodAge(NFood) = 1 Else: FoodAge(NFood) = 0

End Sub


Private Sub FoodToRNDPosition(wF As Long)
    With FOOD(wF)
        .POS.x = wMinX + Rnd * (wMaxX - wMinX)
        .POS.y = wMinY + Rnd * (wMaxY - wMinY)
    End With
End Sub
Public Sub DrawFOOD()
    Dim I   As Long

    ' vbDRAW.CC.SetSourceColor vbGreen


    For I = 0 To NFood
        With FOOD(I)
            'vbDRAW.CC.Ellipse .Pos.X, .Pos.y, 5, 5
            'vbDRAW.CC.Fill

            If InsideBB(CameraBB, FOOD(I).POS) Then
                vbDrawCC.RenderSurfaceContent "FoodIcon", .POS.x - FoodSize, .POS.y - FoodSize, , , CAIRO_FILTER_FAST, 0.75
       '        If FoodAge(I) > 0 Then vbDrawCC.RenderSurfaceContent "FoodIconLight", .POS.x - FoodSize * 3, .POS.y - FoodSize * 3, , , CAIRO_FILTER_FAST, FoodAge(I)
                If FoodAge(I) > 0 Then vbDrawCC.RenderSurfaceContent "FoodIconLight", .POS.x - FoodSize * 2, .POS.y - FoodSize * 2, , , CAIRO_FILTER_FAST, FoodAge(I)
                
            End If
        End With
'        FoodAge(I) = FoodAge(I) - 0.002
        FoodAge(I) = FoodAge(I) - 0.0015 '--2024
        If FoodAge(I) < 0# Then FoodAge(I) = 0#
    Next



End Sub

Public Sub FoodMoveAndCheckEaten()
    Dim I   As Long
    Dim J   As Double
    Dim HeadPosition As geoVector2D
    Dim Hvel As geoVector2D

    Dim D   As Double
    Dim dx  As Double
    Dim dy  As Double
    Dim vD  As Double
    Dim vDx As Double
    Dim vDy As Double


    Dim GrabR As Double


    For I = 0 To NFood
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
                        If J = PLAYER Then
                            If Snake(PLAYER).IsDying = 0 Then
                                ' MultipleSounds.playsound "eatfruit.wav"

                                MultipleSounds.PlaySound SoundPlayerChomp, 0, -1000
                                
                            End If
                        Else
                            HeadPosition = Snake(PLAYER).GetHEADPos
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
SC.x = COStable(A) * 0.0025
SC.y = SINtable(A) * 0.0025
.Vel = VectorSUM(.Vel, SC)
'------------

            
            If .POS.x < wMinX Then .POS.x = wMinX: .Vel.x = -.Vel.x
            If .POS.y < wMinY Then .POS.y = wMinY: .Vel.y = -.Vel.y
            If .POS.x > wMaxX Then .POS.x = wMaxX: .Vel.x = -.Vel.x
            If .POS.y > wMaxY Then .POS.y = wMaxY: .Vel.y = -.Vel.y


        End With
    Next


End Sub

Public Sub CreateFoodFromDeadSnake(wS As Long)
    Dim I         As Long
    For I = 0 To Snake(wS).Ntokens - 2  '1
        NFood = NFood + 1
        If NFood > MaxFood Then
            MaxFood = NFood + 20
            ReDim Preserve FOOD(MaxFood)
            ReDim Preserve FoodAge(MaxFood)

        End If

        With FOOD(NFood)
            .POS = Snake(wS).GetTokenPos(I)
            .Vel.x = (Rnd * 2 - 1) * 0.125
            .Vel.y = (Rnd * 2 - 1) * 0.125
        End With
        FoodAge(NFood) = 1
    Next
End Sub


Public Function PointToNearestFood(Head As tPosAndVel, ByVal SnakeIDX As Long) As geoVector2D
    Dim I         As Long
    Dim J         As Long
    Dim D         As Double
    Dim dx        As Double
    Dim dy        As Double
    Dim MIND      As Double
    Dim Direct    As Double

    Dim HX#, HY#, HVX#, HVY#
    HX = Head.POS.x
    HY = Head.POS.y
    HVX = Head.Vel.x
    HVY = Head.Vel.y
Dim SFFM#
SFFM = Snake(SnakeIDX).SearchForFOODMODE

    MIND = 1E+32

    For I = 0 To NFood
        dx = FOOD(I).POS.x - HX
        dy = FOOD(I).POS.y - HY
        D = dx * dx + dy * dy

        Direct = -Sgn(dx * HVX + dy * HVY)    '''' Consider nearer the ones in front

        ''        Direct = Direct + 2#
        '        Direct = Direct + 1.65 '--2024

        Direct = Direct + SFFM  '--2024

        ''D = D * Direct * (1# - FoodAge(I) * 0.95)
        'D = D * Direct * (1# - FoodAge(I) * 0.98)
        D = D * Direct * (1# - FoodAge(I) * 0.99)    '--2024

        If D < MIND Then
            MIND = D
            J = I
        End If
    Next

    PointToNearestFood = VectorNormalize(VectorSUB(FOOD(J).POS, Head.POS))

End Function

