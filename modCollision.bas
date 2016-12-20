Attribute VB_Name = "modCollision"
Option Explicit
Public Type tBB
    minX    As Double
    minY    As Double
    maxX    As Double
    maxY    As Double
End Type

Public Function InsideBB(BB As tBB, P As geoVector2D) As Boolean
    InsideBB = True

    If P.x < BB.minX Then InsideBB = False: Exit Function
    If P.y < BB.minY Then InsideBB = False: Exit Function
    If P.x > BB.maxX Then InsideBB = False: Exit Function
    If P.y > BB.maxY Then InsideBB = False: Exit Function

End Function

Public Function BBOverlapping(BB1 As tBB, BB2 As tBB) As Boolean
    BBOverlapping = True

    If BB1.maxX < BB2.minX Then BBOverlapping = False: Exit Function
    If BB1.maxY < BB2.minY Then BBOverlapping = False: Exit Function
    If BB1.minX > BB2.maxX Then BBOverlapping = False: Exit Function
    If BB1.minY > BB2.maxY Then BBOverlapping = False: Exit Function

End Function



'Public Sub CheckCollisionsOnlyPlayer()
'    Dim R      As Double
'    Dim R2     As Double
'    Dim I      As Long
'    Dim J      As Long
'    Dim TokenPosition     As geoVector2D
'    Dim Dx     As Double
'    Dim Dy     As Double
'    Dim D      As Double
'    Dim MIND   As Double
'    Dim BB     As tBB
'
'
'    Dim HeadPosI   As geoVector2D
'
'
'    'PLAYER to ENEMY
'    HeadPosI = snake(player).GetHEADPos
'    R = snake(player).DIAM * 0.5
'
'    For I = 1 To NSnakes
'
'        If InsideBB(Snake(I).getBB, HeadPosI) Then
'
'            R2 = Snake(I).DIAM * 0.5
'
'            For J = 0 To Snake(I).Ntokens - 1
'
'                TokenPosition = Snake(I).GetTokenPos(J)
'
'                Dx = HeadPosI.x - TokenPosition.x
'                Dy = HeadPosI.y - TokenPosition.y
'                D = Dx * Dx + Dy * Dy
'                MIND = R + R2
'                MIND = MIND * MIND
'
'                If D < MIND Then
'                    'Player Dead
'                    If snake(player).IsDying = 0 Then MultipleSounds.playsound SoundPlayerDeath
'
'                    snake(player).Kill: Exit For
'                End If
'
'            Next
'
'        End If
'    Next
'    '----------------------------------
'    'ENEMY to PLAYER
'    BB = snake(player).getBB
'    R2 = snake(player).DIAM * 0.5
'    For I = 1 To NSnakes
'        HeadPosI = Snake(I).GetHEADPos
'
'        If InsideBB(BB, HeadPosI) Then
'
'            R = Snake(I).DIAM * 0.5
'
'            For J = 0 To snake(player).Ntokens - 1
'                TokenPosition = snake(player).GetTokenPos(J)
'                Dx = HeadPosI.x - TokenPosition.x
'                Dy = HeadPosI.y - TokenPosition.y
'
'                D = Dx * Dx + Dy * Dy
'                MIND = R + R2
'                MIND = MIND * MIND
'
'                If D < MIND Then
'
'                    'If Snake(I).IsDying = 0 Then MultipleSounds.playsound SoundEnenmyKilledByMe
'                    If Snake(I).IsDying = 0 Then MultipleSounds.playsound SoundEnenmyKilledByMe
'
'                    Snake(I).Kill
'                End If
'            Next
'        End If
'
'    Next
'
'
'    '----------------------------------
'
'
'End Sub

Public Sub CheckCollisionsALLtoALL()
    Dim Ri  As Double
    Dim Rj  As Double
    Dim I   As Long
    Dim J   As Long

    Dim K   As Long


    Dim TokenPosition As geoVector2D
    Dim dx  As Double
    Dim dy  As Double
    Dim D   As Double
    Dim MIND As Double
    Dim BB  As tBB


    Dim HeadPosI As geoVector2D
    Dim HeadPosJ As geoVector2D

    For I = 0 To NSnakes
        Snake(I).UpdateBB
    Next

    For I = 0 To NSnakes - 1
        HeadPosI = Snake(I).GetHEADPos
        Ri = Snake(I).Diam * 0.5

        For J = I + 1 To NSnakes

            If InsideBB(Snake(J).getBB, HeadPosI) Then

                Rj = Snake(J).Diam * 0.5
                MIND = Ri + Rj
                MIND = MIND * MIND

                For K = 0 To Snake(J).Ntokens - 1

                    TokenPosition = Snake(J).GetTokenPos(K)

                    dx = HeadPosI.x - TokenPosition.x
                    dy = HeadPosI.y - TokenPosition.y
                    D = dx * dx + dy * dy

                    If D < MIND Then
                        'Dead Snake
                        If I = PLAYER Then
                            If Snake(I).IsDying = 0 Then MultipleSounds.PlaySound SoundPlayerDeath
                        Else
                            If Snake(I).IsDying = 0 Then

                                HeadPosI = Snake(PLAYER).GetHEADPos
                                dx = HeadPosI.x - TokenPosition.x
                                dy = HeadPosI.y - TokenPosition.y
                                D = Sqr(dx * dx + dy * dy)
                                'MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 3, -10000, 10000), ClampLong(-D * 0.8, -10000, 0)
                                MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 1, -10000, 0)

                                HeadPosI = Snake(I).GetHEADPos
                            End If

                        End If
                        Snake(I).Kill: Exit For
                    End If
                Next
            End If

            HeadPosJ = Snake(J).GetHEADPos
            Rj = Snake(J).Diam * 0.5

            If InsideBB(Snake(I).getBB, HeadPosJ) Then
                ' Ri = Snake(I).DIAM * 0.5
                MIND = Ri + Rj
                MIND = MIND * MIND

                For K = 0 To Snake(I).Ntokens - 1

                    TokenPosition = Snake(I).GetTokenPos(K)

                    dx = HeadPosJ.x - TokenPosition.x
                    dy = HeadPosJ.y - TokenPosition.y
                    D = dx * dx + dy * dy


                    If D < MIND Then
                        'Player Dead
                        If J = PLAYER Then
                            If Snake(J).IsDying = 0 Then MultipleSounds.PlaySound SoundPlayerDeath
                        Else
                            If I = PLAYER Then
                                If Snake(J).IsDying = 0 Then

                                    HeadPosJ = Snake(PLAYER).GetHEADPos
                                    dx = HeadPosJ.x - TokenPosition.x
                                    dy = HeadPosJ.y - TokenPosition.y
                                    D = Sqr(dx * dx + dy * dy)
                                    'MultipleSounds.PlaySound SoundEnenmyKilledByMe, ClampLong(-dx * 3, -10000, 10000), ClampLong(-D * 0.8, -10000, 0)
                                    MultipleSounds.PlaySound SoundEnenmyKilledByMe, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 1, -10000, 0)

                                    HeadPosJ = Snake(J).GetHEADPos
                                End If

                            Else
                                If Snake(J).IsDying = 0 Then
                                    HeadPosI = Snake(PLAYER).GetHEADPos
                                    dx = HeadPosI.x - TokenPosition.x
                                    dy = HeadPosI.y - TokenPosition.y
                                    D = Sqr(dx * dx + dy * dy)
                                    'MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 3, -10000, 10000), ClampLong(-D * 0.8, -10000, 0)
                                    MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 1, -10000, 0)

                                    HeadPosI = Snake(I).GetHEADPos
                                End If
                            End If
                            Snake(J).Kill: Exit For
                        End If
                    End If
                Next
            End If
        Next
    Next



    '----------------------------------


End Sub




Public Function NewSnakePosition(Idx As Long) As geoVector2D

    Dim BB  As tBB
    Dim POS As geoVector2D
    Dim PlayerHeadPOS As geoVector2D
    Dim dx  As Double
    Dim dy  As Double
    Dim C   As Long
    Dim InsBB As Boolean


    If Idx = PLAYER Then

        Do
            InsBB = False
            POS.x = wMinX + (wMaxX - wMinX) * Rnd
            POS.y = wMinY + (wMaxY - wMinY) * Rnd
            For C = 1 To NSnakes
                BB = Snake(C).getBB
                If InsideBB(BB, POS) Then InsBB = True: Exit For
            Next
        Loop While InsBB

    Else
        ' BB = Snake(PLAYER).getBB
        PlayerHeadPOS = Snake(PLAYER).GetHEADPos

        Do
            POS.x = wMinX + (wMaxX - wMinX) * Rnd
            POS.y = wMinY + (wMaxY - wMinY) * Rnd
            dx = POS.x - PlayerHeadPOS.x
            dy = POS.y - PlayerHeadPOS.y
            'Loop While InsideBB(BB, POS) Or ((dx * dx + dy * dy) < 40000)
        Loop While (dx * dx + dy * dy) < 40000


    End If

    NewSnakePosition = POS


End Function
