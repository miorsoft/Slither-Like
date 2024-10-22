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



Private Head As tPosAndVel

Private Token() As tPosAndVel
Private Const MaxNtok As Long = 999 ' 199

Public Ntokens As Long
Private InvNtokens As Double
Private InvNtokens2 As Double

'Public MySIZE  As Double

Public tokIDX As Long
Public MyIDX As Long

Private Color As Long
Public ColorR As Double
Public ColorG As Double
Public ColorB As Double

Private DirChange As geoVector2D
Private EyEANG As Double
Private oldANG As Double
Public Diam As Double
Public InvDiam As Double
Public Radius As Double




Public IsDying As Long
Private DyingTime As Long

Private BB  As tBB

Public TongueOut As Double
Public TongueLen As Double

Private fLength As Double

Public FASTspeed As Long
Public FASTspeedCOUNT As Long

Private Const SPEED As Double = 2.5 ' 3 '3.5

Public SearchForFOODMODE As Double




Public Property Get GetSize() As Double
    GetSize = fLength
End Property
Public Property Let SetSize(V As Double)
    fLength = V
    'DIAM = 5 * (fLength + 10) ^ (0.4)
    'Diam = 5 * (fLength + 15) ^ (0.45)
    
    'Diam = 5 * (fLength + 15) ^ (0.5) '------ 2nd video
    
    'Diam = 5# * (fLength + 16#) ^ (0.5) ''' OK
    
    Diam = 1 + 4.5 * (fLength) ^ 0.5
    
    Radius = Diam * 0.5
    
InvDiam = 1# / Diam

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
Friend Sub UpdateBB()
    Dim I   As Long
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
Friend Sub Init(x As Double, y As Double, Idx As Long, Leng As Double)
    Dim I   As Long

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

    Head.POS.x = x
    Head.POS.y = y
    Head.Vel.x = Rnd * 2 - 1
    Head.Vel.y = Rnd * 2 - 1
    'Color = RGB(120 + Rnd * 135, 120 + Rnd * 135, 120 + Rnd * 135)
    ColorR = 0.45 + Rnd * 0.5
    ColorG = 0.45 + Rnd * 0.5
    ColorB = 0.45 + Rnd * 0.5

    UpdateBB



End Sub

Private Sub RandomizeMe()
    Dim x   As Double
    Dim y   As Double

    Dim POS As geoVector2D

    Dim I   As Long

    IsDying = 0    ' False

    'fLength = 2
    SetSize = 8 '3

    Ntokens = fLength
InvNtokens = 1 / Ntokens
InvNtokens2 = 1 / (Ntokens - 1)

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

    Head.POS.x = x
    Head.POS.y = y
    Head.Vel.x = Rnd * 2 - 1
    Head.Vel.y = Rnd * 2 - 1
    'Color = RGB(120 + Rnd * 135, 120 + Rnd * 135, 120 + Rnd * 135)
    ColorR = 0.45 + Rnd * 0.5
    ColorG = 0.45 + Rnd * 0.5
    ColorB = 0.45 + Rnd * 0.5

    UpdateBB

SearchForFOODMODE = 1.2 + Rnd * 0.8

End Sub


Friend Sub DRAW(DrawBoundingBox As Long)
    Dim ANG As Double
    Dim A   As Double
    Dim A2  As Double


    Dim Reye As Double
    Dim Reye2 As Double

    Dim DEye As Double

    Dim I   As Long
    Dim J   As Long
    Dim P   As geoVector2D

    Dim X1  As Double
    Dim Y1  As Double
    Dim x2  As Double
    Dim Y2  As Double
    Dim dx  As Double
    Dim dy  As Double

    Dim C   As Double
    Dim S   As Double
    Dim K   As Double
    '  Dim InvNtokens
    Dim CADD As Double

    Dim StepD As Long



    If BBOverlapping(CameraBB, BB) = False Then Exit Sub


    Reye = Diam * 0.25
    Reye2 = Reye * 0.85
    DEye = Reye * 1.65    '2


    'If MyIDX = PLAYER Then Stop

    'StepD = 1 + Diam * 0.0125

    StepD = 1 + Diam * Diam * 0.0004 '0.00045
 


    With vbDrawCC

        If IsDying Then
StepD = 1

            .SetSourceRGBA Rnd, Rnd, Rnd, 0.33
            For I = tokIDX + 1 To tokIDX + Ntokens - 1 Step StepD
                J = I Mod Ntokens
                P = Token(J).POS
                '.Ellipse P.x, P.y, Diam, Diam
                .Arc P.x, P.y, Radius
                
                .Fill
            Next

            If (CNT - DyingTime > 40) Then '104 '' Blinking time
                CreateFoodFromDeadSnake MyIDX
                RandomizeMe
            End If

        Else


            '            InvNtokens = 1# / (Ntokens)
            '            K = 0
            '            For I = tokIDX + 1 To tokIDX + Ntokens - 1 Step StepD   'Senza -1 disegna anche testa
            '
            '                J = I Mod Ntokens
            '                P = Token(J).POS
            '                K = K + StepD
            '
            '                CADD = Cos((-K * InvNtokens) * PI2 * 5#) * 0.2
            '
            '                .SetSourceRGBA ColorR + CADD, ColorG + CADD, ColorB + CADD    ', 0.9
            '
            '                .Ellipse P.x, P.y, Diam, Diam
            '                .Fill
            '            Next

            .SetLineWidth Diam

            K = 0
            I = tokIDX + StepD * 1 + 1
            J = I Mod Ntokens
            P = Token(J).POS
            K = K + StepD
            
            .MoveTo P.x, P.y

            For I = tokIDX + StepD * 1 + 2 To tokIDX + Ntokens Step StepD    'Senza -1 disegna anche testa

                'J = I Mod Ntokens
                DivRem I, Ntokens, J
                
                
                P = Token(J).POS
                K = K + StepD
  
                'To Avoid Head->Queue Line:
                If J <> (tokIDX + StepD) Mod Ntokens Then
                
                            '    'CADD = Cos((-K * InvNtokens) * PI2 * 5#) * 0.2 - FASTspeed * 0.25
                
                       'CADD = Cos((K * InvNtokens2) * PI2 * 3.5) * 0.2 - FASTspeed * 0.25
                Dim AA&
                AA = ((K * InvNtokens2) * 360 * 3.5) Mod 360
                CADD = COStable(AA) * 0.2 - FASTspeed * 0.25

                    .SetSourceRGB ColorR + CADD, ColorG + CADD, ColorB + CADD
                    .LineTo P.x, P.y
                    .Stroke
                End If

                .MoveTo P.x, P.y

            Next
                       
            .LineTo Head.POS.x, Head.POS.y
            .Stroke

        End If




        '........................................
        'Drawhead EYES

        '.SetLineWidth 1 * Diam
        .SetLineWidth 0.05 * Diam

        ANG = Atan2(Head.Vel.x, Head.Vel.y)

        '--------Tongue

        TongueLen = TongueLen * 0.3 + TongueOut * 0.7
        TongueOut = TongueOut * 0.6
        C = Cos(ANG) * Diam * TongueLen * 0.45
        S = Sin(ANG) * Diam * TongueLen * 0.45
        X1 = Head.POS.x + C
        Y1 = Head.POS.y + S
        .SetSourceRGBA 1, 0.6, 0.6, 0.4
        '.Ellipse X1, Y1, Diam * 0.67, Diam * 0.67
        .Arc X1, Y1, Diam * 0.335
        .Fill
        'Draw Head
        .SetSourceRGB ColorR + CADD, ColorG + CADD, ColorB + CADD
        '.Ellipse Head.POS.x, Head.POS.y, Diam, Diam
        .Arc Head.POS.x, Head.POS.y, Radius
        .Fill
        '---------------------------------

'        If MyIDX = PLAYER Then
'            EyEANG = Atan2(DirChange.x, DirChange.y)
            EyEANG = EyEANG + 0.25 * AngleDIFF(EyEANG, Atan2(DirChange.x, DirChange.y)) '--2024
'        Else
'            EyEANG = oldANG
'        End If
        If IsDying Then EyEANG = (CNT + DyingTime) * 0.3

        .SetSourceRGB 1, 1, 0.85    'EYEBALLS
        A = ANG - PIh * 0.5
        C = Cos(A) * Reye
        S = Sin(A) * Reye
        X1 = Head.POS.x + C
        Y1 = Head.POS.y + S
        '.Ellipse X1, Y1, DEye, DEye
        .Arc X1, Y1, DEye * 0.5
        
        x2 = Head.POS.x - S
        Y2 = Head.POS.y + C
        '.Ellipse x2, Y2, DEye, DEye
        .Arc x2, Y2, DEye * 0.5
        .Fill

        .SetSourceRGBA 0, 0, 0, 0.55    'eye contour
        '.Ellipse X1, Y1, DEye, DEye
         .Arc X1, Y1, DEye * 0.5
        .Stroke
        x2 = Head.POS.x - S
        Y2 = Head.POS.y + C
        '.Ellipse x2, Y2, DEye, DEye
         .Arc x2, Y2, DEye * 0.5
        .Stroke



        .SetSourceColor 0    'PUPILS
        dx = Cos(EyEANG) * (Reye * 0.5)
        dy = Sin(EyEANG) * (Reye * 0.5)
        '.Ellipse X1 + dx, Y1 + dy, Reye2, Reye2
        .Arc X1 + dx, Y1 + dy, Reye2 * 0.5
        '.Ellipse x2 + dx, Y2 + dy, Reye2, Reye2
        .Arc x2 + dx, Y2 + dy, Reye2 * 0.5
        .Fill
        '........................................


        If DrawBoundingBox Then
            .Rectangle BB.minX, BB.minY, BB.maxX - BB.minX, BB.maxY - BB.minY
            .Stroke
        End If

    End With












End Sub


Friend Sub MOVE()
    Dim I   As Long
    Dim J   As Long
    Dim J2  As Long


    If IsDying Then Exit Sub

    tokIDX = (tokIDX + 1) Mod Ntokens


    If MyIDX = PLAYER And (Not (AIcontrol)) Then 'Player
'        DirChange = VectorMUL(MousePos, 0.01 * InvDiam)  '.001
        DirChange = VectorMUL(VectorNormalize(MousePos), 3.25 * InvDiam)

        ''   DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, HEAD.POS, HEAD.Vel))

        Head.Vel = VectorNormalize(VectorSUM(Head.Vel, DirChange))

    Else    'Other Snakes (AI)

        oldANG = Atan2(Head.Vel.x, Head.Vel.y)


    '    'If Cos((CNT + MyIDX * 50500) * 0.0075) < -0.98 Then  ''' Ramble -.75 -0.0625
        If Cos((CNT + MyIDX * 50500) * 0.0075) < -1# * (NSnakes - MyIDX * 0.0625) * InvNSnakes Then ''' Ramble -.75 -0.0625
    '--2024
    '    If Cos((CNT + MyIDX * 50500) * 0.0075) > (NSnakes - MyIDX * 0.0625) * InvNSnakes Then   ''' Ramble -.75 -0.0625
    
    
'    Debug.Print MyIDX, Cos((CNT + MyIDX * 50500) * 0.0075), (NSnakes - MyIDX * 0.0625) * InvNSnakes
'    Stop
    
            DirChange.x = Rnd * 2 - 1
            DirChange.y = Rnd * 2 - 1
            'Commented to make AI sometimes not to Work
            'DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, HEAD.POS, HEAD.Vel))
        Else
            DirChange = PointToNearestFood(Head, MyIDX)
            DirChange = VectorSUM(DirChange, AvoidEnemy(MyIDX, Head.POS, Head.Vel))
        End If



        'DirChange = VectorMUL(DirChange, 0.15)
        DirChange = VectorMUL(DirChange, 3.25 * InvDiam)
        
        Head.Vel = VectorNormalize(VectorSUM(Head.Vel, DirChange))

        If FASTspeed Then
          '  If Rnd < 0.025 Then FASTspeed = False
            If Rnd < 0.02 Then FASTspeed = False
        Else
            'If Rnd < 0.00625 Then FASTspeed = True
            If Rnd < 0.005 Then FASTspeed = True
        End If


    End If


    '    HEAD.POS = VectorSUM(HEAD.POS, VectorMUL(HEAD.Vel, MySIZE))


    If FASTspeed And (fLength > 3) Then
        FASTspeedCOUNT = FASTspeedCOUNT + 1
        Head.POS = VectorSUM(Head.POS, VectorMUL(Head.Vel, SPEED * 1.72))
    Else
        Head.POS = VectorSUM(Head.POS, VectorMUL(Head.Vel, SPEED))
    End If



    If FASTspeedCOUNT > 25 Then '20
        FASTspeedCOUNT = 0
        AddFoodParticle Token((tokIDX + 1) Mod Ntokens).POS, False
        SetSize = GetSize - FoodLengthValue
    End If

    '    If HEAD.Pos.X < wMinX Then HEAD.Pos.X = wMaxX
    '    If HEAD.Pos.Y < wMinY Then HEAD.Pos.Y = wMaxY
    '    If HEAD.Pos.X > wMaxX Then HEAD.Pos.X = wMinX
    '    If HEAD.Pos.Y > wMaxY Then HEAD.Pos.Y = wMinY

    If Head.POS.x < wMinX Then Head.POS.x = wMinX: Head.Vel.x = -Head.Vel.x
    If Head.POS.y < wMinY Then Head.POS.y = wMinY: Head.Vel.y = -Head.Vel.y
    If Head.POS.x > wMaxX Then Head.POS.x = wMaxX: Head.Vel.x = -Head.Vel.x
    If Head.POS.y > wMaxY Then Head.POS.y = wMaxY: Head.Vel.y = -Head.Vel.y



    Token(tokIDX).POS = Head.POS


    For I = tokIDX + Ntokens - 1 To tokIDX + 1 Step -1
        J = I Mod Ntokens
        J2 = J - 1
        If J2 < 0 Then J2 = J2 + Ntokens
        'If J2 <> tokIDX Then Token(J).POS = VectorSUM(VectorMUL(Token(J).POS, 0.18), VectorMUL(Token(J2).POS, 0.82)) '2nd Video
        If J2 <> tokIDX Then Token(J).POS = VectorSUM(VectorMUL(Token(J).POS, 0.17), VectorMUL(Token(J2).POS, 0.83))
    Next
    If fLength \ 1 < Ntokens Then
        Ntokens = Ntokens - 1
        InvNtokens = 1 / Ntokens
InvNtokens2 = 1 / (Ntokens - 1)
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
        InvNtokens = 1 / Ntokens
        InvNtokens2 = 1 / (Ntokens - 1)
    End If


End Sub

Friend Sub Kill()

    If Not (IsDying) Then
        IsDying = -1
        DyingTime = CNT
    End If

End Sub

Private Sub Class_Initialize()
SearchForFOODMODE = 1.2 + Rnd * 0.8


End Sub
