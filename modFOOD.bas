Attribute VB_Name = "modFOOD"
Option Explicit

Public Type tPosAndVel
    POS     As geoVector2D
    Vel     As geoVector2D
End Type

Public FOOD() As tPosAndVel
Public FOODcolor() As Long


Public NFood As Long

Public Const FoodSize As Double = 9

Public Const FoodLengthValue As Double = 1



Public Sub InitFOOD(HowMuch As Long)
    Dim I   As Long

    NFood = HowMuch

    ReDim FOOD(NFood)
    For I = 0 To NFood

        With FOOD(I)
            .POS.x = wMinX + Rnd * (wMaxX - wMinX)
            .POS.y = wMinY + Rnd * (wMaxY - wMinY)
        End With
    Next

End Sub
Private Sub RemoveFood(wF As Long)
    Dim I   As Long

    NFood = NFood - 1
    For I = wF To NFood
        FOOD(I) = FOOD(I + 1)
    Next

End Sub

Public Sub AddFoodParticle(POS As geoVector2D)

    NFood = NFood + 1
    ReDim Preserve FOOD(NFood)
    With FOOD(NFood)
        .Vel.x = 0
        .Vel.y = 0
        .POS = POS
    End With


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
                vbDRAW.CC.RenderSurfaceContent "FoodIcon", .POS.x - FoodSize, .POS.y - FoodSize, , , CAIRO_FILTER_FAST, 0.75
            End If
        End With
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
                    GrabR = Snake(J).Diam * 0.5 + FoodSize * 0.5
                    GrabR = GrabR * GrabR
                    If D < GrabR Then
                        If J = PLAYER Then
                            If Snake(PLAYER).IsDying = 0 Then
                                ' MultipleSounds.playsound "eatfruit.wav"

                                MultipleSounds.PlaySound SoundPlayerChomp
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
            If .POS.x < wMinX Then .POS.x = wMinX: .Vel.x = -.Vel.x
            If .POS.y < wMinY Then .POS.y = wMinY: .Vel.y = -.Vel.y
            If .POS.x > wMaxX Then .POS.x = wMaxX: .Vel.x = -.Vel.x
            If .POS.y > wMaxY Then .POS.y = wMaxY: .Vel.y = -.Vel.y


        End With
    Next


End Sub

Public Sub CreateFoodFromDeadSnake(wS As Long)
    Dim I   As Long
    For I = 0 To Snake(wS).Ntokens - 2    '1
        NFood = NFood + 1
        ReDim Preserve FOOD(NFood)
        With FOOD(NFood)
            .POS = Snake(wS).GetTokenPos(I)
            .Vel.x = (Rnd * 2 - 1) * 0.125
            .Vel.y = (Rnd * 2 - 1) * 0.125
        End With
    Next
End Sub


Public Function PointToNearestFood(HeadPos As geoVector2D) As geoVector2D
    Dim I   As Long
    Dim J   As Long
    Dim D   As Double
    Dim dx  As Double
    Dim dy  As Double
    Dim MIND As Double

    MIND = 1E+32

    For I = 0 To NFood
        dx = FOOD(I).POS.x - HeadPos.x
        dy = FOOD(I).POS.y - HeadPos.y
        D = dx * dx + dy * dy
        If D < MIND Then
            MIND = D
            J = I
        End If
    Next

    PointToNearestFood = VectorNormalize(VectorSUB(FOOD(J).POS, HeadPos))

End Function

Public Function AvoidEnemy(Idx As Long, POS As geoVector2D, Vel As geoVector2D) As geoVector2D
    Dim I   As Long
    Dim J   As Long

    Dim TPleft As geoVector2D
    Dim TPRight As geoVector2D
    Dim TP  As geoVector2D

    Dim C   As Double
    Dim S   As Double
    Dim A   As Double

    Dim tEsc As geoVector2D
    Dim Dmin As Double
    Dim D1  As Double
    Dim D2  As Double
    Dim Diam As Double

    Diam = Snake(Idx).Diam


    A = Atan2(Vel.x, Vel.y)

    TPleft.x = POS.x - Cos(A - 0.5) * Diam
    TPleft.y = POS.y - Sin(A - 0.5) * Diam
    TPRight.x = POS.x - Cos(A + 0.5) * Diam
    TPRight.y = POS.y - Sin(A + 0.5) * Diam

    Dmin = 1E+28


    Diam = (Diam + 30) * 8
    Diam = Diam * Diam

    'If Idx = PLAYER Then Stop

    For I = 0 To NSnakes
        If I <> Idx Then

            For J = 1 To Snake(I).Ntokens - 1


                TP = Snake(I).GetTokenPos(J)


                If Sgn((TP.x - POS.x) * Vel.x) + Sgn((TP.y - POS.y) * Vel.y) > 1 Then

                    D1 = DistFromPointSQU(TP, TPleft)
                    D2 = DistFromPointSQU(TP, TPRight)

                    If (D1 < Diam) Or (D2 < Diam) Then


                        If D1 < Dmin Or D2 < Dmin Then
                            If D1 < D2 Then
                                tEsc.x = Cos(A - 0.25) * 8
                                tEsc.y = Sin(A - 0.25) * 8
                            Else
                                tEsc.x = Cos(A + 0.25) * 8
                                tEsc.y = Sin(A + 0.25) * 8
                            End If
                            If D1 < Dmin Then Dmin = D1 Else: Dmin = D2
                        End If
                    End If
                End If

            Next
        End If
    Next

    AvoidEnemy = tEsc




End Function



