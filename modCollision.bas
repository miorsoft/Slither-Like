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


    If BB1.maxX < BB2.minX Then Exit Function
    If BB1.maxY < BB2.minY Then Exit Function
    If BB1.minX > BB2.maxX Then Exit Function
    If BB1.minY > BB2.maxY Then Exit Function
    
    BBOverlapping = True

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
'    R = snake(player).radius
'
'    For I = 1 To NSnakes
'
'        If InsideBB(Snake(I).getBB, HeadPosI) Then
'
'            R2 = Snake(I).radius
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
'    R2 = snake(player).radius
'    For I = 1 To NSnakes
'        HeadPosI = Snake(I).GetHEADPos
'
'        If InsideBB(BB, HeadPosI) Then
'
'            R = Snake(I).radius
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
    Dim Ri     As Double
    Dim Rj     As Double
    Dim I      As Long
    Dim J      As Long

    Dim K      As Long


    Dim TokenPosition As geoVector2D
    Dim dx     As Double
    Dim dy     As Double
    Dim D      As Double
    Dim MIND   As Double
    Dim BB     As tBB


    Dim HeadPosI As geoVector2D
    Dim HeadPosJ As geoVector2D

    For I = 0 To NSnakes
        Snake(I).UpdateBB
    Next

    For I = 0 To NSnakes - 1
        HeadPosI = Snake(I).GetHEADPos
        Ri = Snake(I).Radius

        For J = I + 1 To NSnakes

            If InsideBB(Snake(J).getBB, HeadPosI) Then

                Rj = Snake(J).Radius
                MIND = Ri + Rj
                MIND = MIND * MIND

                For K = 0 To Snake(J).Ntokens '- 1

                    TokenPosition = Snake(J).GetTokenPos(K)

                    dx = HeadPosI.x - TokenPosition.x
                    dy = HeadPosI.y - TokenPosition.y
                    D = dx * dx + dy * dy

                    If D < MIND Then
                        'Dead Snake
                        If I = SNAKECAMERA Then
                            If Snake(I).IsDying = 0 Then MultipleSounds.PlaySound SoundPlayerDeath
                        Else
                            If Snake(I).IsDying = 0 Then

                                HeadPosI = Snake(SNAKECAMERA).GetHEADPos
                                dx = HeadPosI.x - TokenPosition.x
                                dy = HeadPosI.y - TokenPosition.y
                                D = Sqr(dx * dx + dy * dy)
                                MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 0.5, -10000, 0)

                                HeadPosI = Snake(I).GetHEADPos
                            End If

                        End If
                        Snake(I).Kill
                         Exit For
                    End If
                Next
            End If

            HeadPosJ = Snake(J).GetHEADPos
            Rj = Snake(J).Radius

            If InsideBB(Snake(I).getBB, HeadPosJ) Then
                ' Ri = Snake(I).radius
                MIND = Ri + Rj
                MIND = MIND * MIND

                For K = 0 To Snake(I).Ntokens '- 1

                    TokenPosition = Snake(I).GetTokenPos(K)

                    dx = HeadPosJ.x - TokenPosition.x
                    dy = HeadPosJ.y - TokenPosition.y
                    D = dx * dx + dy * dy

                    If D < MIND Then
                        'Player Dead
                        If J = SNAKECAMERA Then
                            If Snake(J).IsDying = 0 Then MultipleSounds.PlaySound SoundPlayerDeath
                        Else
                            If I = SNAKECAMERA Then
                                If Snake(J).IsDying = 0 Then
                                    HeadPosJ = Snake(SNAKECAMERA).GetHEADPos
                                    dx = HeadPosJ.x - TokenPosition.x
                                    dy = HeadPosJ.y - TokenPosition.y
                                    D = Sqr(dx * dx + dy * dy)
                                    MultipleSounds.PlaySound SoundEnenmyKilledByMe, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 1, -10000, 0)

                                    HeadPosJ = Snake(J).GetHEADPos
                                End If

                            Else
                                If Snake(J).IsDying = 0 Then
                                    HeadPosI = Snake(SNAKECAMERA).GetHEADPos
                                    dx = HeadPosI.x - TokenPosition.x
                                    dy = HeadPosI.y - TokenPosition.y
                                    D = Sqr(dx * dx + dy * dy)
                                    MultipleSounds.PlaySound SoundEnenmyKilled, ClampLong(-dx * 2, -10000, 10000), ClampLong(-D * 0.5, -10000, 0)

                                    HeadPosI = Snake(I).GetHEADPos
                                End If
                            End If
'                            Snake(J).Kill: Exit For
                        End If
                        Snake(J).Kill: Exit For
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
            POS.x = wMinX + (wMaxX - wMinX) * RndM
            POS.y = wMinY + (wMaxY - wMinY) * RndM
            For C = 1 To NSnakes
                BB = Snake(C).getBB
                If InsideBB(BB, POS) Then InsBB = True: Exit For
            Next
        Loop While InsBB

    Else
        ' BB = Snake(PLAYER).getBB
        PlayerHeadPOS = Snake(PLAYER).GetHEADPos

        Do
            POS.x = wMinX + (wMaxX - wMinX) * RndM
            POS.y = wMinY + (wMaxY - wMinY) * RndM
            dx = POS.x - PlayerHeadPOS.x
            dy = POS.y - PlayerHeadPOS.y
            'Loop While InsideBB(BB, POS) Or ((dx * dx + dy * dy) < 40000)
        Loop While (dx * dx + dy * dy) < 40000


    End If

    NewSnakePosition = POS


End Function



Public Function AvoidEnemy(ByVal Idx As Long, POS As geoVector2D, Vel As geoVector2D) As geoVector2D
    Dim I   As Long
    Dim J   As Long

    Dim TPleft As geoVector2D
    Dim TPRight As geoVector2D
    Dim TP  As geoVector2D

    Dim C   As Double
    Dim S   As Double
    Dim A   As Double

    Dim EscapeDirection As geoVector2D
    Dim Dmin As Double
    Dim D1  As Double
    Dim D2  As Double
    Dim Diam As Double

    Diam = Snake(Idx).Diam

    A = Atan2(Vel.x, Vel.y)
    'TPleft.x = POS.x - Cos(A - 0.5) * Diam
    'TPleft.y = POS.y - Sin(A - 0.5) * Diam
    'TPRight.x = POS.x - Cos(A + 0.5) * Diam
    'TPRight.y = POS.y - Sin(A + 0.5) * Diam
    'USING TABLE---------------------------------------
Dim CC#, SS#
Dim tbA As Long
If A < 0.5 Then A = A + PI2
tbA = (A - 0.5) * 360# * InvPI2
TPleft.x = POS.x - COStable(tbA) * Diam
TPleft.y = POS.y - SINtable(tbA) * Diam
If (A + 0.5) > PI2 Then A = A - PI2
tbA = (A + 0.5) * 360# * InvPI2
TPRight.x = POS.x - COStable(tbA) * Diam
TPRight.y = POS.y - SINtable(tbA) * Diam
'---------------------------------------
If A < 0# Then A = A + PI2



    Dmin = 1E+28

    'Diam = (Diam + 30) * 3    ' 8 '''' Distance Sense
    'Diam = Diam * Diam

    For I = 0 To NSnakes
        If I <> Idx Then

Diam = Snake(Idx).Diam * 2.5 + 1 * Snake(I).Diam '--2024
Diam = Diam * Diam

            For J = 0 To Snake(I).Ntokens '- 1


                TP = Snake(I).GetTokenPos(J)

                If Sgn((TP.x - POS.x) * Vel.x + (TP.y - POS.y) * Vel.y) > 0 Then    'Correct!

                    D1 = DistFromPointSQU(TP, TPleft)
                    D2 = DistFromPointSQU(TP, TPRight)

                    If (D1 < Diam) Or (D2 < Diam) Then
                        If D1 < Dmin Or D2 < Dmin Then
                            If D1 < D2 Then
                                'EscapeDirection.x = Cos(A - 0.25) * 8#
                                'EscapeDirection.y = Sin(A - 0.25) * 8#
If A < 0.25 Then A = A + PI2
tbA = (A - 0.25) * 360# * InvPI2
EscapeDirection.x = COStable(tbA) * 8#
EscapeDirection.y = SINtable(tbA) * 8#
                            Else
                                'EscapeDirection.x = Cos(A + 0.25) * 8#
                                'EscapeDirection.y = Sin(A + 0.25) * 8#
If (A + 0.25) > PI2 Then A = A - PI2
tbA = (A + 0.25) * 360# * InvPI2
EscapeDirection.x = COStable(tbA) * 8#
EscapeDirection.y = SINtable(tbA) * 8#
                                
                            End If
                            If D1 < Dmin Then Dmin = D1 Else: Dmin = D2
                        End If
                    End If
                End If

            Next
        End If
    Next

    AvoidEnemy = EscapeDirection


End Function




