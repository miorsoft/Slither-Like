VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private HEAD   As tPosAndVel

Private Token() As tPosAndVel
Private Const MaxNtok As Long = 199

Public Ntokens As Long


'Public MySIZE  As Double

Public tokIDX  As Long
Public MyIDX   As Long

Private Color  As Long
Private ColorR As Double
Private ColorG As Double
Private ColorB As Double

Private DirChange As geoVector2D
Private EyEANG As Double
Private oldANG As Double
Public Diam    As Double


Public IsDying As Long
Private DyingTime As Long

Private BB     As tBB

Public TongueOut As Double
Public TongueLen As Double

Private fLength As Double

Public FASTspeed As Long
Public FASTspeedCOUNT As Long


Public Property Get GetSize() As Double
    GetSize = fLength
End Property
Public Property Let SetSize(V As Double)
    fLength = V
    'DIAM = 5 * (fLength + 10) ^ (0.4)
    Diam = 5 * (fLength + 15) ^ (0.45)

End Property


Friend Property Get GetHEADPos() As geoVector2D
    GetHEADPos = Token(tokIDX).POS
End Property

Friend Property Get GetHEADVel() As geoVector2D
    GetHEADVel = Token(tokIDX).Vel
End Property

Friend Property Let SetHEADVel(NV As geoVector2D)
    Token(tokIDX).Vel = NV
End Property

Friend Property Get GetTokenPos(wT As Long) As geoVector2D
    GetTokenPos = Token(wT).POS
End Property


Friend Function getBB() As tBB
    getBB = BB
End Function
'Friend Function getBB() As tBB
'    Dim I      As Long
'
'    getBB.minX = wMaxX * 2
'    getBB.minY = wMaxY * 2
'    getBB.maxX = wMinX * 2
'    getBB.maxY = wMinY * 2
'
'
'    For I = 0 To Ntokens
'        With Token(I).Pos
'            If .x < getBB.minX Then getBB.minX = .x
'            If .y < getBB.minY Then getBB.minY = .y
'            If .x > getBB.maxX Then getBB.maxX = .x
'            If .y > getBB.maxY Then getBB.maxY = .y
'        End With
'    Next
'
'End Function
Public Sub UpdateBB()
    Dim I      As Long
    BB.minX = wMaxX * 2
    BB.minY = wMaxY * 2
    BB.maxX = wMinX * 2
    BB.maxY = wMinY * 2

    For I = 0 To Ntokens - 1
        With Token(I).POS
            If .x - Diam < BB.minX Then BB.minX = .x - Diam
            If .y - Diam < BB.minY Then BB.minY = .y - Diam
            If .x + Diam > BB.maxX Then BB.maxX = .x + Diam
            If .y + Diam > BB.maxY Then BB.maxY = .y + Diam
        End With
    Next

End Sub
Public Sub Init(x As Double, y As Double, Idx As Long, Leng As Double)
    Dim I      As Long

    'fLength = Leng
    SetSize = Leng

    Ntokens = fLength

    ReDim Token(MaxNtok)
    MyIDX = Idx

    '    If MyIDX <> PLAYER Then
    '        MySIZE = (1 - Rnd * 0.65) * 4.5
    '    Else
    '        MySIZE = 3
    '    End If
    '''    DIAM = MySIZE * 7

    tokIDX = 1

    x = wMinX + (wMaxX - wMinX) * Rnd
    y = wMinY + (wMaxY - wMinY) * Rnd


    For I = 0 To Ntokens
        Token(I).POS.x = x
        Token(I).POS.y = y
    Next

    HEAD.POS.x = x
    HEAD.POS.y = y
    HEAD.Vel.x = Rnd * 2 - 1
    HEAD.Vel.y = Rnd * 2 - 1
    'Color = RGB(120 + Rnd * 135, 120 + Rnd * 135, 120 + Rnd * 135)
    ColorR = 0.45 + Rnd * 0.5
    ColorG = 0.45 + Rnd * 0.5
    ColorB = 0.45 + Rnd * 0.5

    UpdateBB



End Sub

Private Sub RandomizeMe()
    Dim x      As Double
    Dim y      As Double

    Dim POS    As geoVector2D

    Dim I      As Long

    IsDying = 0    ' False

    'fLength = 2
    SetSize = 2

    Ntokens = fLength


    ReDim Token(MaxNtok)

    ''    MySIZE = (1 - Rnd * 0.65) * 4.5
    ''    'DIAM = MySIZE * 7

    tokIDX = 1


    POS = NewSnakePosition(MyIDX)
    x = POS.x
    y = POS.y


    For I = 0 To Ntokens
        Token(I).POS.x = x
        Token(I).POS.y = y
    Next

    HEAD.POS.x = x
    HEAD.POS.y = y
    HEAD.Vel.x = Rnd * 2 - 1
    HEAD.Vel.y = Rnd * 2 - 1
    'Color = RGB(120 + Rnd * 135, 120 + Rnd * 135, 120 + Rnd * 135)
    ColorR = 0.45 + Rnd * 0.5
    ColorG = 0.45 + Rnd * 0.5
    ColorB = 0.45 + Rnd * 0.5

    UpdateBB


End Sub


Public Sub DRAW(DrawBoundingBox As Long)
    Dim ANG    As Double
    Dim A      As Double
    Dim A2     As Double


    Dim Reye   As Double
    Dim Reye2  As Double

    Dim DEye   As Double

    Dim I      As Long
    Dim J      As Long
    Dim P      As geoVector2D

    Dim X1     As Double
    Dim Y1     As Double
    Dim x2     As Double
    Dim Y2     As Double
    Dim dx     As Double
    Dim dy     As Double

    Dim C      As Double
    Dim S      As Double
    Dim K      As Double
    Dim InvNotkens
    Dim CADD   As Double

Dim StepD As Long



    If BBOverlapping(CameraBB, BB) = False Then Exit Sub


    Reye = Diam * 0.25
    Reye2 = Reye * 0.85
    DEye = Reye * 1.65    '2


'If MyIDX = PLAYER Then Stop

StepD = 1 + Diam * 0.0125
'If StepD <> 1 Then Stop



    With vbDRAW.CC

        If IsDying Then

            .SetSourceRGBA Rnd, Rnd, Rnd, 0.33
            For I = tokIDX + 1 To tokIDX + Ntokens - 1 Step StepD
                J = I Mod Ntokens
                P = Token(J).POS
                .Ellipse P.x, P.y, Diam, Diam
                .Fill
            Next

            If (CNT - DyingTime > 104) Then
                CreateFoodFromDeadSnake MyIDX
                RandomizeMe
            End If

        Else
            '.SetSourceColor Color, 0.9
            '.SetSourceRGBA ColorR, ColorG, ColorB, 0.9

            InvNotkens = 1# / (Ntokens)
            K = 0
            For I = tokIDX + 1 To tokIDX + Ntokens - 1 Step StepD   'Senza -1 disegna anche testa

                J = I Mod Ntokens
                P = Token(J).POS
                K = K + StepD

                CADD = Cos((-K * InvNotkens) * PI2 * 5#) * 0.2
                
                .SetSourceRGBA ColorR + CADD, ColorG + CADD, ColorB + CADD    ', 0.9

                .Ellipse P.x, P.y, Diam, Diam
                .Fill


            Next
        End If




        '........................................
        'Drawhead EYES
        ANG = Atan2(HEAD.Vel.x, HEAD.Vel.y)

        '--------Tongue

        TongueLen = TongueLen * 0.3 + TongueOut * 0.7
        TongueOut = TongueOut * 0.6
        C = Cos(ANG) * Diam * TongueLen * 0.45
        S = Sin(ANG) * Diam * TongueLen * 0.45
        X1 = HEAD.POS.x + C
        Y1 = HEAD.POS.y + S
        .SetSourceRGBA 1, 0.6, 0.6, 0.4
        .Ellipse X1, Y1, Diam * 0.67, Diam * 0.67
        .Fill
        .SetSourceRGBA ColorR + CADD, ColorG + CADD, ColorB + CADD
        .Ellipse HEAD.POS.x, HEAD.POS.y, Diam, Diam
        .Fill
        '---------------------------------

        If MyIDX = PLAYER Then
            EyEANG = Atan2(DirChange.x, DirChange.y)
        Else
            EyEANG = oldANG
        End If
        If IsDying Then EyEANG = (CNT + DyingTime) * 0.3

        .SetSourceRGB 1, 1, 0.8    'EYEBALLS
        A = ANG - PIh * 0.5
        C = Cos(A) * Reye
        S = Sin(A) * Reye
        X1 = HEAD.POS.x + C
        Y1 = HEAD.POS.y + S
        .Ellipse X1, Y1, DEye, DEye
        x2 = HEAD.POS.x - S
        Y2 = HEAD.POS.y + C
        .Ellipse x2, Y2, DEye, DEye
        .Fill

        .SetSourceRGBA 0, 0, 0, 0.5    'eye contour
        .Ellipse X1, Y1, DEye, DEye
        .Stroke
        x2 = HEAD.POS.x - S
        Y2 = HEAD.POS.y + C
        .Ellipse x2, Y2, DEye, DEye
        .Stroke



        .SetSourceColor 0    'PUPILS
        dx = Cos(EyEANG) * (Reye * 0.5)
        dy = Sin(EyEANG) * (Reye * 0.5)
        .Ellipse X1 + dx, Y1 + dy, Reye2, Reye2
        .Ellipse x2 + dx, Y2 + dy, Reye2, Reye2
        .Fill
        '........................................


        If DrawBoundingBox Then
            .Rectangle BB.minX, BB.minY, BB.maxX - BB.minX, BB.maxY - BB.minY
            .Stroke
        End If

    End With












End Sub


Public Sub MOVE()
    Dim I      As Long
    Dim J      As Long
    Dim J2     As Long


    If IsDying Then Exit Sub

    tokIDX = tokIDX + 1
    tokIDX = tokIDX Mod Ntokens


    If MyIDX = PLAYER Then    'Player
        DirChange = VectorMUL(MousePos, 0.002)    '.004
        
    ''   DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, HEAD.POS, HEAD.Vel))
        
        HEAD.Vel = VectorNormalize(VectorSUM(HEAD.Vel, DirChange))
        
       
        
    Else    'Other Snakes

        oldANG = Atan2(HEAD.Vel.x, HEAD.Vel.y)

        If Cos((CNT + MyIDX * 50500) * 0.0075) < -0.0625 Then
            DirChange.x = Rnd * 2 - 1
            DirChange.y = Rnd * 2 - 1
            DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, HEAD.POS, HEAD.Vel))
        Else
            DirChange = PointToNearestFood(HEAD.POS)
            DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, HEAD.POS, HEAD.Vel))
        End If

        DirChange = VectorMUL(DirChange, 0.15)    '.3
        HEAD.Vel = VectorNormalize(VectorSUM(HEAD.Vel, DirChange))

If FASTspeed Then
If Rnd < 0.025 Then FASTspeed = False
Else
If Rnd < 0.00625 Then FASTspeed = True
End If


    End If


    '    HEAD.POS = VectorSUM(HEAD.POS, VectorMUL(HEAD.Vel, MySIZE))


    If FASTspeed And fLength > 3 Then
        FASTspeedCOUNT = FASTspeedCOUNT + 1
        HEAD.POS = VectorSUM(HEAD.POS, VectorMUL(HEAD.Vel, 6))
    Else
        HEAD.POS = VectorSUM(HEAD.POS, VectorMUL(HEAD.Vel, 3.5))
    End If



    If FASTspeedCOUNT > 20 Then
        FASTspeedCOUNT = 0
        AddFoodParticle Token((tokIDX + 1) Mod Ntokens).POS
        SetSize = GetSize - FoodLengthValue
    End If




    '    If HEAD.Pos.X < wMinX Then HEAD.Pos.X = wMaxX
    '    If HEAD.Pos.Y < wMinY Then HEAD.Pos.Y = wMaxY
    '    If HEAD.Pos.X > wMaxX Then HEAD.Pos.X = wMinX
    '    If HEAD.Pos.Y > wMaxY Then HEAD.Pos.Y = wMinY

    If HEAD.POS.x < wMinX Then HEAD.POS.x = wMinX: HEAD.Vel.x = -HEAD.Vel.x
    If HEAD.POS.y < wMinY Then HEAD.POS.y = wMinY: HEAD.Vel.y = -HEAD.Vel.y
    If HEAD.POS.x > wMaxX Then HEAD.POS.x = wMaxX: HEAD.Vel.x = -HEAD.Vel.x
    If HEAD.POS.y > wMaxY Then HEAD.POS.y = wMaxY: HEAD.Vel.y = -HEAD.Vel.y



    Token(tokIDX).POS = HEAD.POS


    For I = tokIDX + Ntokens - 1 To tokIDX + 1 Step -1
        J = I Mod Ntokens
        J2 = J - 1
        If J2 < 0 Then J2 = J2 + Ntokens
        If J2 <> tokIDX Then Token(J).POS = VectorSUM(VectorMUL(Token(J).POS, 0.18), VectorMUL(Token(J2).POS, 0.82))

    Next
    If fLength \ 1 < Ntokens Then
        Ntokens = Ntokens - 1
    End If

    ' fLength = fLength + 0.002 * Rnd
    If fLength \ 1 > Ntokens Then
        Ntokens = Ntokens + 1
        If Ntokens > MaxNtok Then
            Ntokens = MaxNtok
        Else
            'Riordina
            For I = Ntokens - 1 To tokIDX + 1 Step -1
                J = I Mod Ntokens
                J2 = I - 1 Mod Ntokens
                Token(J).POS = Token(J2).POS
            Next
        End If
    End If


End Sub

Public Sub Kill()

    If Not (IsDying) Then
        IsDying = -1
        DyingTime = CNT
    End If

End Sub

