Attribute VB_Name = "modMATH"
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
' Suggestions or new Tools are wellcome!
' Most Function taken from http://paulbourke.net/geometry
'***********************************************************************************

Option Explicit

Public Const Pi As Double = 3.14159265358979    ' Atn (1) * 4
Public Const PI2 As Double = 6.28318530717959    'PI * 2
Public Const PIh As Double = 1.5707963267949    ' PI * 0.5

Public Type geoVector2D
    x       As Double
    y       As Double
    Bool    As Boolean
End Type

Public Type geoLine
    P1      As geoVector2D
    P2      As geoVector2D
    ANG     As Double
    Bool    As Boolean
End Type

Public Type geoCircle
    Center  As geoVector2D
    Radius  As Double
    Bool    As Boolean
End Type

Public Type geoARC
    Circle  As geoCircle
    A1      As Double
    A2      As Double
    X1      As Double
    Y1      As Double
    x2      As Double
    Y2      As Double
End Type




Public Function mkPoint(x As Double, y As Double) As geoVector2D
    mkPoint.x = x
    mkPoint.y = y
End Function

Public Function mkLine(P1 As geoVector2D, P2 As geoVector2D) As geoLine
    Dim dx  As Double
    Dim dy  As Double

    mkLine.P1 = P1
    mkLine.P2 = P2
    dx = P2.x - P1.x
    dy = P2.y - P1.y
    mkLine.ANG = Atan2(dx, dy)
End Function
Public Sub UpdateLineAng(ByRef L As geoLine)
    Dim dx  As Double
    Dim dy  As Double
    dx = L.P2.x - L.P1.x
    dy = L.P2.y - L.P1.y
    L.ANG = Atan2(dx, dy)
    If L.ANG < 0 Then L.ANG = L.ANG + PI2
End Sub
Public Sub UpdateArcPts(ByRef A As geoARC)
'Knowing A1 and A2 of the arc
'calc x1,y1 and x2,y2
    With A
        .X1 = .Circle.Center.x + .Circle.Radius * Cos(.A1)
        .Y1 = .Circle.Center.y + .Circle.Radius * Sin(.A1)
        .x2 = .Circle.Center.x + .Circle.Radius * Cos(.A2)
        .Y2 = .Circle.Center.y + .Circle.Radius * Sin(.A2)
    End With
End Sub

Public Function mkLine2(X1 As Double, Y1 As Double, x2 As Double, Y2 As Double) As geoLine
    Dim dx  As Double
    Dim dy  As Double

    mkLine2.P1.x = X1
    mkLine2.P1.y = Y1
    mkLine2.P2.x = x2
    mkLine2.P2.y = Y2
    dx = x2 - X1
    dy = Y2 - Y1
    mkLine2.ANG = Atan2(dx, dy)
End Function

Public Function mkCircle(C As geoVector2D, R As Double) As geoCircle
    mkCircle.Center = C
    mkCircle.Radius = R
End Function

Public Function mkCircle2(cx As Double, cy As Double, R As Double) As geoCircle
    mkCircle2.Center.x = cx
    mkCircle2.Center.y = cy
    mkCircle2.Radius = R
End Function

Public Function mkCircle3Points(ByRef P1 As geoVector2D, ByRef P2 As geoVector2D, ByRef P3 As geoVector2D) As geoCircle
    mkCircle3Points.Bool = False

    If privIsNotPerpendicular(P1, P2, P3) Then
        mkCircle3Points = privCircle3Points(P1, P2, P3)
    ElseIf privIsNotPerpendicular(P1, P3, P2) Then
        mkCircle3Points = privCircle3Points(P1, P3, P2)
    ElseIf privIsNotPerpendicular(P2, P1, P3) Then
        mkCircle3Points = privCircle3Points(P2, P1, P3)
    ElseIf privIsNotPerpendicular(P2, P3, P1) Then
        mkCircle3Points = privCircle3Points(P2, P3, P1)
    ElseIf privIsNotPerpendicular(P3, P2, P1) Then
        mkCircle3Points = privCircle3Points(P3, P2, P1)
    ElseIf privIsNotPerpendicular(P3, P1, P2) Then
        mkCircle3Points = privCircle3Points(P3, P1, P2)
    Else
        'msgBox "The three pts are perpendicular to axis"
        If (P2.x - P1.x) = 0 Then
            mkCircle3Points.Center.y = (P2.y + P1.y) / 2
            mkCircle3Points.Center.x = (P3.x + P2.x) / 2
            mkCircle3Points.Radius = DistFromPoint(mkCircle3Points.Center, P2)
            mkCircle3Points.Bool = True
        End If
        If (P3.x - P2.x) = 0 Then
            mkCircle3Points.Center.y = (P3.y + P2.y) / 2
            mkCircle3Points.Center.x = (P2.x + P1.x) / 2
            mkCircle3Points.Radius = DistFromPoint(mkCircle3Points.Center, P2)
            mkCircle3Points.Bool = True
        End If
    End If

End Function

Private Function privCircle3Points(ByRef P1 As geoVector2D, ByRef P2 As geoVector2D, ByRef P3 As geoVector2D) As geoCircle
    Dim aSlope As Double
    Dim bSlope As Double
    aSlope = (P2.y - P1.y) / (P2.x - P1.x)
    bSlope = (P3.y - P2.y) / (P3.x - P2.x)
    If (Abs(aSlope - bSlope) <= 0.000001) Then    'checking whether the given points are colinear.
        MsgBox "The three pts are colinear"
        Exit Function
    End If
    privCircle3Points.Center.x = (aSlope * bSlope * (P1.y - P3.y) + bSlope * (P1.x + P2.x) - aSlope * (P2.x + P3.x)) / (2 * (bSlope - aSlope))
    privCircle3Points.Center.y = -1 * (privCircle3Points.Center.x - (P1.x + P2.x) / 2) / aSlope + (P1.y + P2.y) / 2
    privCircle3Points.Radius = DistFromPoint(P1, privCircle3Points.Center)
    privCircle3Points.Bool = True
End Function

Private Function privIsNotPerpendicular(ByRef P1 As geoVector2D, ByRef P2 As geoVector2D, ByRef P3 As geoVector2D) As Boolean
    Dim xDelta_A As Double
    Dim yDelta_A As Double
    Dim xDelta_B As Double
    Dim yDelta_B As Double

    privIsNotPerpendicular = True

    ' Check the given point are perpendicular to x or y axis
    yDelta_A = P2.y - P1.y
    xDelta_A = P2.x - P1.x
    yDelta_B = P3.y - P2.y
    xDelta_B = P3.x - P2.x

    ' checking whether the line of the two pts are vertical
    If (Abs(xDelta_A) <= 0.0000001 And Abs(yDelta_B) <= 0.0000001) Then
        'The points are pependicular and parallel to x-y axis
        privIsNotPerpendicular = False
        Exit Function             'return false;
    ElseIf (Abs(yDelta_A) <= 0.0000001) Then
        'A line of two point are perpendicular to x-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(yDelta_B) <= 0.0000001) Then
        'A line of two point are perpendicular to x-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(xDelta_A) <= 0.0000001) Then
        'A line of two point are perpendicular to y-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    ElseIf (Abs(xDelta_B) <= 0.0000001) Then
        'A line of two point are perpendicular to y-axis
        privIsNotPerpendicular = False
        Exit Function             'return true;
    Else

    End If

End Function

Public Function Atan2(x As Double, y As Double) As Double
    If x Then
        Atan2 = -Pi + Atn(y / x) - (x > 0) * Pi
    Else
        Atan2 = -PIh - (y > 0) * Pi
    End If
End Function
Public Function FowlerAngle(ByRef dx As Double, ByRef dy As Double) As Double
'Faster than Atan2
'http://paulbourke.net/geometry/fowler/

'   This function is due to Rob Fowler.  Given dy and dx between 2 points
'   A and B, we calculate a number in [0.0, 8.0) which is a monotonic
'   function of the direction from A to B.
'
'   (0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0) correspond to
'   (  0,  45,  90, 135, 180, 225, 270, 315, 360) degrees, measured
'   counter-clockwise from the positive x axis.

    Dim Adx As Double    'Absolute Values of Dx and Dy
    Dim Ady As Double
    Dim Code As Long    'Angular Region Classification Code

    Const K = PI2 * 0.125


    Adx = Abs(dx)                 'Compute the absolute values.
    Ady = Abs(dy)

    If Adx < Ady Then Code = 1 Else: Code = 0
    If dx < 0 Then Code = Code + 2
    If dy < 0 Then Code = Code + 4

    Select Case Code
        Case 0
            If dx = 0 Then
                FowlerAngle = 0
            Else
                FowlerAngle = Ady / Adx    ';  /* [  0, 45] */
            End If
        Case 1
            FowlerAngle = 2 - (Adx / Ady)    ';      /* ( 45, 90] */
        Case 3
            FowlerAngle = 2 + (Adx / Ady)    ';      /* ( 90,135) */
        Case 2
            FowlerAngle = 4 - (Ady / Adx)    ';      /* [135,180] */
        Case 6
            FowlerAngle = 4 + (Ady / Adx)    ';      /* (180,225] */
        Case 7
            FowlerAngle = 6 - (Adx / Ady)    ';      /* (225,270) */
        Case 5
            FowlerAngle = 6 + (Adx / Ady)    ';      /* [270,315) */
        Case 4
            FowlerAngle = 8 - (Ady / Adx)    ';      /* [315,360) */
    End Select

    FowlerAngle = FowlerAngle * K

End Function
Public Function Atan2Fast1(ByRef x As Double, ByRef y As Double) As Double
'http://www.gamedev.net/topic/441464-manually-implementing-atan2-or-atan/
'maximum error slightly larger than 4 degrees.
'public double aTan2(double y, double x)
'{   double coeff_1 = Math.PI / 4d;  double coeff_2 = 3d * coeff_1;  double abs_y = Math.abs(y); double angle;   if (x >= 0d) {      double r = (x - abs_y) / (x + abs_y);       angle = coeff_1 - coeff_1 * r;  } else {        double r = (x + abs_y) / (abs_y - x);       angle = coeff_2 - coeff_1 * r;  }   return y < 0d ? -angle : angle;}
    Const C1 As Double = 0.785398163397448    'atn(1)
    Const C2 As Double = 2.35619449019234    'atn(1)*3
    Dim AbsY As Double
    Dim R   As Double

    AbsY = Abs(y)
    If (x >= 0) Then
        R = (x - AbsY) / (x + AbsY)
        Atan2Fast1 = C1 - C1 * R
    Else
        R = (x + AbsY) / (AbsY - x)
        Atan2Fast1 = C2 - C1 * R
    End If

    If y < 0 Then Atan2Fast1 = -Atan2Fast1

End Function
Public Function Atan2Fast2(ByRef x As Double, ByRef y As Double) As Double
'http://lists.apple.com/archives/perfoptimization-dev/2005/Jan/msg00051.html
'|error| < 0.005 radians

    Dim Z   As Double

    If x = 0 Then
        If (y > 0) Then Atan2Fast2 = PIh: Exit Function
        If (y = 0) Then Atan2Fast2 = 0: Exit Function
        Atan2Fast2 = -PIh: Exit Function
    End If

    Z = y / x
    If (Abs(Z) < 1) Then
        Atan2Fast2 = Z / (1 + 0.28 * Z * Z)
        If (x < 0) Then
            If (y < 0) Then Atan2Fast2 = Atan2Fast2 + Pi: Exit Function
            Atan2Fast2 = Atan2Fast2 + Pi: Exit Function
        End If
    Else
        Atan2Fast2 = PIh - Z / (Z * Z + 0.28)
        If (y < 0) Then Atan2Fast2 = Atan2Fast2 + Pi: Exit Function
    End If

    If Atan2Fast2 < 0 Then Atan2Fast2 = Atan2Fast2 + PI2

End Function

Public Function AngleDIFF(A1 As Double, A2 As Double) As Double
'double difference = secondAngle - firstAngle;

    AngleDIFF = A2 - A1
    While AngleDIFF < -Pi
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > Pi
        AngleDIFF = AngleDIFF - PI2
    Wend

End Function



Public Function LineLen(ByRef L As geoLine) As Double
    Dim dx  As Double
    Dim dy  As Double
    dx = L.P2.x - L.P1.x
    dy = L.P2.y - L.P1.y
    LineLen = Sqr(dx * dx + dy * dy)
End Function

Public Function DistFromPoint(ByRef P1 As geoVector2D, ByRef P2 As geoVector2D) As Double
    Dim dx  As Double
    Dim dy  As Double
    dx = P2.x - P1.x
    dy = P2.y - P1.y
    DistFromPoint = Sqr(dx * dx + dy * dy)
End Function

Public Function DistFromPoint2(ByRef P As geoVector2D, x As Double, y As Double) As Double
    Dim dx  As Double
    Dim dy  As Double
    dx = x - P.x
    dy = y - P.y
    DistFromPoint2 = Sqr(dx * dx + dy * dy)
End Function


Public Function DistFromPointSQU(ByRef P1 As geoVector2D, ByRef P2 As geoVector2D) As Double
    Dim dx  As Double
    Dim dy  As Double
    dx = P2.x - P1.x
    dy = P2.y - P1.y
    DistFromPointSQU = (dx * dx + dy * dy)
End Function

Public Function DistFromLine(ByRef P As geoVector2D, ByRef L As geoLine) As Double
'
' Returns distance from the line, or if the intersecting point on the line nearest
' the point tested is outside the endpoints of the line, the distance to the
' nearest endpoint.
'
' Returns 9999 on 0 denominator conditions.
    Dim LineMag As Double, u As Double
    Dim iX As Double, iY As Double    ' intersecting point

    LineMag = LineLen(L)
    If LineMag < 0.000001 Then DistFromLine = 9999: Exit Function

    u = (((P.x - L.P1.x) * (L.P2.x - L.P1.x)) + ((P.y - L.P1.y) * (L.P2.y - L.P1.y)))
    u = u / (LineMag * LineMag)

    'If u < 0.00001 Or u > 1 Then
    '    '// closest point does not fall within the line segment, take the shorter distance
    '    '// to an endpoint
    '    ix = DistFromPoint(P, L.P1)
    '    iy = DistFromPoint(P, L.P2)
    '    If ix > iy Then DistFromLine = iy Else DistFromLine = ix
    'Else
    ' Intersecting point is on the line, use the formula
    iX = L.P1.x + u * (L.P2.x - L.P1.x)
    iY = L.P1.y + u * (L.P2.y - L.P1.y)
    DistFromLine = DistFromPoint2(P, iX, iY)
    'End If

End Function
Public Function NearestFromLine(ByRef P As geoVector2D, ByRef L As geoLine) As geoVector2D
'
' Returns distance from the line, or if the intersecting point on the line nearest
' the point tested is outside the endpoints of the line, the distance to the
' nearest endpoint.
'
' Returns 9999 on 0 denominator conditions.
    Dim LineMag As Double, u As Double
    Dim iX As Double, iY As Double    ' intersecting point
    NearestFromLine.Bool = False
    LineMag = LineLen(L)
    If LineMag < 0.000001 Then Exit Function

    u = (((P.x - L.P1.x) * (L.P2.x - L.P1.x)) + ((P.y - L.P1.y) * (L.P2.y - L.P1.y)))
    u = u / (LineMag * LineMag)

    NearestFromLine.Bool = True
    If u < 0.00001 Or u > 1 Then
        '// closest point does not fall within the line segment, take the shorter distance
        '// to an endpoint
        iX = DistFromPoint(P, L.P1)
        iY = DistFromPoint(P, L.P2)
        If iX < iY Then NearestFromLine = L.P1 Else NearestFromLine = L.P2
    Else
        ' Intersecting point is on the line, use the formula
        NearestFromLine.x = L.P1.x + u * (L.P2.x - L.P1.x)
        NearestFromLine.y = L.P1.y + u * (L.P2.y - L.P1.y)

    End If

End Function
Public Function IntersectOfLines(ByRef L1 As geoLine, ByRef L2 As geoLine) As geoVector2D
    Dim D   As Double
    Dim NA  As Double
    Dim NB  As Double
    Dim DX1 As Double
    Dim DX2 As Double
    Dim DY1 As Double
    Dim DY2 As Double
    Dim uA  As Double
    Dim uB  As Double

    DX1 = L1.P2.x - L1.P1.x
    DY1 = L1.P2.y - L1.P1.y
    DX2 = L2.P2.x - L2.P1.x
    DY2 = L2.P2.y - L2.P1.y

    ' Denominator for ua and ub are the same, so store this calculation
    D = (DY2) * (DX1) - _
        (DX2) * (DY1)

    'NA and NB are calculated as seperate values for readability
    NA = (DX2) * (L1.P1.y - L2.P1.y) - _
         (DY2) * (L1.P1.x - L2.P1.x)

    NB = (DX1) * (L1.P1.y - L2.P1.y) - _
         (DY1) * (L1.P1.x - L2.P1.x)

    ' Make sure there is not a division by zero - this also indicates that
    ' the lines are parallel.
    ' If NA and NB were both equal to zero the lines would be on top of each
    ' other (coincidental).  This check is not done because it is not
    ' necessary for this implementation (the parallel check accounts for this).
    IntersectOfLines.Bool = False

    If D = 0 Then Exit Function

    ' Calculate the intermediate fractional point that the lines potentially intersect.
    uA = NA / D

    ' The fractional point will be between 0 and 1 inclusive if the lines
    ' intersect.  If the fractional calculation is larger than 1 or smaller
    ' than 0 the lines would need to be longer to intersect.
    If uA >= 0 Then
        If uA <= 1 Then
            ' Calculate the intermediate fractional point that the lines potentially intersect.
            uB = NB / D
            If uB >= 0 Then
                If uB <= 1 Then
                    IntersectOfLines.x = L1.P1.x + (uA * (DX1))
                    IntersectOfLines.y = L1.P1.y + (uA * (DY1))
                    IntersectOfLines.Bool = True
                End If
            End If
        End If
    End If

End Function
Public Function IntersectOfLines2(ByRef L1 As geoLine, ByRef L2 As geoLine) As geoVector2D
'********************************************
'*  Intersection of LINES (not segments)    *
'********************************************

    Dim D   As Double
    Dim NA  As Double
    Dim NB  As Double
    Dim DX1 As Double
    Dim DX2 As Double
    Dim DY1 As Double
    Dim DY2 As Double
    Dim uA  As Double
    Dim uB  As Double

    DX1 = L1.P2.x - L1.P1.x
    DY1 = L1.P2.y - L1.P1.y
    DX2 = L2.P2.x - L2.P1.x
    DY2 = L2.P2.y - L2.P1.y

    ' Denominator for ua and ub are the same, so store this calculation
    D = (DY2) * (DX1) - _
        (DX2) * (DY1)

    'NA and NB are calculated as seperate values for readability
    NA = (DX2) * (L1.P1.y - L2.P1.y) - _
         (DY2) * (L1.P1.x - L2.P1.x)

    NB = (DX1) * (L1.P1.y - L2.P1.y) - _
         (DY1) * (L1.P1.x - L2.P1.x)

    ' Make sure there is not a division by zero - this also indicates that
    ' the lines are parallel.
    ' If NA and NB were both equal to zero the lines would be on top of each
    ' other (coincidental).  This check is not done because it is not
    ' necessary for this implementation (the parallel check accounts for this).
    IntersectOfLines2.Bool = False

    If D = 0 Then Exit Function

    ' Calculate the intermediate fractional point that the lines potentially intersect.
    uA = NA / D
    ' Calculate the intermediate fractional point that the lines potentially intersect.
    'uB = NB / D

    IntersectOfLines2.x = L1.P1.x + (uA * (DX1))
    IntersectOfLines2.y = L1.P1.y + (uA * (DY1))
    IntersectOfLines2.Bool = True


End Function
Public Sub IntersectCircleLine(ByRef C As geoCircle, _
                               ByRef L As geoLine, _
                               ByRef Sol1 As geoVector2D, _
                               ByRef Sol2 As geoVector2D)


    Dim dx  As Double
    Dim dy  As Double
    Dim I   As Double
    Dim AA  As Double
    Dim BB  As Double
    Dim CC  As Double
    Dim mu  As Double

    Sol1.Bool = False
    Sol2.Bool = False


    dx = L.P2.x - L.P1.x
    dy = L.P2.y - L.P1.y

    AA = dx * dx + dy * dy        '
    If AA = 0 Then Exit Sub

    BB = 2 * ((dx) * (L.P1.x - C.Center.x) + _
              (dy) * (L.P1.y - C.Center.y))


    CC = (C.Center.x) ^ 2 + (C.Center.y) ^ 2 + _
         (L.P1.x) ^ 2 + _
         (L.P1.y) ^ 2 - _
         2 * (C.Center.x * L.P1.x + C.Center.y * L.P1.y) - (C.Radius) ^ 2

    I = BB * BB - 4 * AA * CC


    Select Case I
        Case Is < 0
            'No intersection
            Exit Sub
        Case 0
            'one intersection
            Sol1.Bool = True
            mu = -BB / (2 * AA)
            Sol1.x = L.P1.x + mu * (dx)
            Sol1.y = L.P1.y + mu * (dy)
        Case Is > 0
            ' two intersections
            ' first intersection
            Sol1.Bool = True
            Sol2.Bool = True
            mu = (-BB + Sqr(BB * BB - 4 * AA * CC)) / (2 * AA)
            Sol1.x = L.P1.x + mu * (dx)
            Sol1.y = L.P1.y + mu * (dy)
            ' second intersection
            mu = (-BB - Sqr(BB * BB - 4 * AA * CC)) / (2 * AA)
            Sol2.x = L.P1.x + mu * (dx)
            Sol2.y = L.P1.y + mu * (dy)

    End Select

    'to make this work for "LINE SEGMENT"
    If NearestFromLine(Sol1, L).Bool = False Then Sol1.Bool = False
    If NearestFromLine(Sol2, L).Bool = False Then Sol2.Bool = False


End Sub

Public Sub IntersectOfCircles(ByRef C1 As geoCircle, _
                              ByRef C2 As geoCircle, _
                              ByRef Sol1 As geoVector2D, _
                              ByRef Sol2 As geoVector2D)

    Dim D   As Double
    Dim c1R As Double
    Dim c2R As Double
    Dim M   As Double
    Dim N   As Double
    Dim A   As Double
    Dim H   As Double
    Dim P   As geoVector2D

    'Calculate distance between centres of circle
    D = DistFromPoint(C1.Center, C2.Center)
    c1R = C1.Radius
    c2R = C2.Radius
    M = c1R + c2R
    N = c1R - c2R
    If (N < 0) Then N = -N

    Sol1.Bool = False
    Sol2.Bool = False

    'No solns
    If (D > M) Then Exit Sub

    'Circle are contained within each other
    If (D < N) Then Exit Sub

    'Circles are the same
    If (D = 0) And (c1R = c2R) Then Exit Sub

    'Solve for a
    A = (c1R * c1R - c2R * c2R + D * D) / (2 * D)

    'Solve for h
    H = Sqr(c1R * c1R - A * A)

    'Calculate point p, where the line through the circle intersection points crosses the line between the circle centers.
    P.x = C1.Center.x + (A / D) * (C2.Center.x - C1.Center.x)
    P.y = C1.Center.y + (A / D) * (C2.Center.y - C1.Center.y)

    '1 soln , circles are touching
    If D = (c1R + c2R) Then Sol1 = P: Sol1.Bool = True: Exit Sub

    '2solns
    H = H / D
    Sol1.x = P.x + (H) * (C2.Center.y - C1.Center.y)
    Sol1.y = P.y - (H) * (C2.Center.x - C1.Center.x)
    Sol2.x = P.x - (H) * (C2.Center.y - C1.Center.y)
    Sol2.y = P.y + (H) * (C2.Center.x - C1.Center.x)
    Sol1.Bool = True
    Sol2.Bool = True

End Sub

Public Function VectorProject(ByRef V As geoVector2D, ByRef Vto As geoVector2D) As geoVector2D
'Poject Vector V to vector Vto
    Dim K   As Double
    Dim D   As Double

    D = Sqr(Vto.x * Vto.x + Vto.y * Vto.y)
    If D = 0 Then Exit Function
    K = (V.x * Vto.x + V.y * Vto.y) / D

    VectorProject.x = (Vto.x / D) * K
    VectorProject.y = (Vto.y / D) * K

End Function

Public Function VectorReflect(ByRef V As geoVector2D, ByRef wall As geoVector2D) As geoVector2D
'Function returning the reflection of one vector around another.
'it's used to calculate the rebound of a Vector on another Vector
'Vector "V" represents current velocity of a point.
'Vector "Wall" represent the angle of a wall where the point Bounces.
'Returns the vector velocity that the point takes after the rebound

    Dim vDot As Double
    Dim D   As Double
    Dim NwX As Double
    Dim NwY As Double

    D = Sqr(wall.x * wall.x + wall.y * wall.y)
    If D = 0 Then Exit Function

    NwX = wall.x / D
    NwY = wall.y / D
    '    'Vect2 = Vect1 - 2 * WallN * (WallN DOT Vect1)
    'vDot = N.DotV(V)
    vDot = V.x * NwX + V.y * NwY

    NwX = NwX * vDot * 2
    NwY = NwY * vDot * 2

    VectorReflect.x = -V.x + NwX
    VectorReflect.y = -V.y + NwY


End Function
Public Function VectorSUM(ByRef V1 As geoVector2D, V2 As geoVector2D) As geoVector2D
    VectorSUM.x = V1.x + V2.x
    VectorSUM.y = V1.y + V2.y
End Function
Public Function VectorSUB(ByRef V1 As geoVector2D, V2 As geoVector2D) As geoVector2D
    VectorSUB.x = V1.x - V2.x
    VectorSUB.y = V1.y - V2.y
End Function
Public Function VectorMUL(ByRef V As geoVector2D, Value As Double) As geoVector2D
'Scalar
    VectorMUL.x = V.x * Value
    VectorMUL.y = V.y * Value
End Function
Public Function VectorDIV(ByRef V As geoVector2D, Value As Double) As geoVector2D
'Scalar
    If Value = 0 Then Exit Function
    VectorDIV.x = V.x / Value
    VectorDIV.y = V.y / Value
End Function
Public Function VectorDOT(ByRef V1 As geoVector2D, V2 As geoVector2D) As Double
'Dot Product
    VectorDOT = V1.x * V2.x + V1.y * V2.y
End Function
Public Function VectorCROSS(ByRef V1 As geoVector2D, V2 As geoVector2D) As Double
'The cross product is a 3D thing, which do not have sense in a 2D world.
    VectorCROSS = V1.x * V2.y - V2.x * V1.y
End Function

Public Function VectorMAG(ByRef V As geoVector2D) As Double
'V magnitude
    VectorMAG = Sqr(V.x * V.x + V.y * V.y)
End Function
Public Function VectorNormalize(ByRef V As geoVector2D) As geoVector2D
'convert vector to UNIT length
    Dim M   As Double
    M = VectorMAG(V)
    If M = 0 Then Exit Function
    VectorNormalize.x = V.x / M
    VectorNormalize.y = V.y / M
    'VectorNormalize = VectorDIV(V, M)
End Function
Public Function VectorNormal(ByRef V As geoVector2D) As geoVector2D
'Normal [Perpendicular]
    VectorNormal.x = -V.y
    VectorNormal.y = V.x
End Function



Public Sub TangentTwoCircles(ByRef C1 As geoCircle, ByRef C2 As geoCircle, _
                             ByRef retL1 As geoLine, ByRef retL2 As geoLine)
'by Roberto Mior (reexre)

    Dim C3  As geoCircle
    Dim R3  As Double
    Dim CM  As geoCircle

    Dim L1P1 As geoVector2D
    Dim L1P2 As geoVector2D
    Dim L2P1 As geoVector2D
    Dim L2P2 As geoVector2D

    Dim A1  As Double
    Dim A2  As Double
    Dim Offset As Double

    CM.Center.x = (C1.Center.x + C2.Center.x) * 0.5
    CM.Center.y = (C1.Center.y + C2.Center.y) * 0.5
    CM.Radius = DistFromPoint(C1.Center, C2.Center) * 0.5

    R3 = C1.Radius - C2.Radius
    If R3 > 0 Then
        C3.Center = C1.Center
        C3.Radius = R3
        L1P2 = C2.Center
        L2P2 = C2.Center
        Offset = C2.Radius
    Else
        C3.Center = C2.Center
        C3.Radius = -R3
        L1P2 = C1.Center
        L2P2 = C1.Center
        Offset = C1.Radius
    End If

    IntersectOfCircles CM, C3, L1P1, L2P1

    If L1P1.Bool Or L1P2.Bool Then

        retL1 = mkLine(L1P1, L1P2)
        retL2 = mkLine(L2P1, L2P2)

        A1 = retL1.ANG + PIh
        A2 = retL2.ANG - PIh

        retL1.P1.x = retL1.P1.x + Cos(A1) * Offset
        retL1.P2.x = retL1.P2.x + Cos(A1) * Offset
        retL1.P1.y = retL1.P1.y + Sin(A1) * Offset
        retL1.P2.y = retL1.P2.y + Sin(A1) * Offset

        retL2.P1.x = retL2.P1.x + Cos(A2) * Offset
        retL2.P2.x = retL2.P2.x + Cos(A2) * Offset
        retL2.P1.y = retL2.P1.y + Sin(A2) * Offset
        retL2.P2.y = retL2.P2.y + Sin(A2) * Offset

    End If

End Sub





Public Function LineOffset(L As geoLine, D As Double, Optional LeftSide As Boolean = False) As geoLine
    Dim iX  As Double
    Dim iY  As Double
    Dim S   As Double

    UpdateLineAng L

    S = IIf(LeftSide, -1, 1)

    iX = S * D * Cos(L.ANG + PIh)
    iY = S * D * Sin(L.ANG + PIh)

    LineOffset.P1.x = L.P1.x + iX
    LineOffset.P1.y = L.P1.y + iY
    LineOffset.P2.x = L.P2.x + iX
    LineOffset.P2.y = L.P2.y + iY


End Function


Public Function Fillet(ByRef L1 As geoLine, ByRef L2 As geoLine, Radius As Double, retArc As geoARC, Optional ModifyLines As Boolean = False)
'by Roberto Mior (reexre)
'Find Arc (of a given radius) tangent to two lines

    Dim tmpL1 As geoLine
    Dim tmpL2 As geoLine
    Dim P   As geoVector2D
    Dim IntesectP As geoVector2D
    Dim ArcCenterP As geoVector2D
    Dim I   As Long
    Dim J   As Long

    Dim L(1 To 4) As geoLine

    Dim A1  As Double
    Dim A2  As Double
    Dim A3  As Double

    Dim arcP1 As geoVector2D
    Dim arcP2 As geoVector2D


    IntesectP = IntersectOfLines2(L1, L2)

    If DistFromPoint(IntesectP, L1.P1) < DistFromPoint(IntesectP, L1.P2) Then
        tmpL1.P1 = IntesectP
        tmpL1.P2 = L1.P2
    Else
        tmpL1.P1 = L1.P1
        tmpL1.P2 = IntesectP
    End If

    If DistFromPoint(IntesectP, L2.P1) < DistFromPoint(IntesectP, L2.P2) Then
        tmpL2.P1 = IntesectP
        tmpL2.P2 = L2.P2
    Else
        tmpL2.P1 = L2.P1
        tmpL2.P2 = IntesectP
    End If


    L(1) = LineOffset(tmpL1, Radius, False)
    L(2) = LineOffset(tmpL1, Radius, True)
    L(3) = LineOffset(tmpL2, Radius, False)
    L(4) = LineOffset(tmpL2, Radius, True)


    ' Find intersection point of 4 offset lines
    ArcCenterP.Bool = False
    For I = 1 To 3
        For J = I + 1 To 4
            Debug.Print I, J
            P = IntersectOfLines(L(I), L(J))
            If P.Bool Then ArcCenterP = P
            If ArcCenterP.Bool Then: Exit For
        Next
        If ArcCenterP.Bool Then: Exit For
    Next
    '------------------------------------------

    retArc.Circle.Center = ArcCenterP
    retArc.Circle.Radius = Radius

    arcP1 = NearestFromLine(ArcCenterP, tmpL1)
    arcP2 = NearestFromLine(ArcCenterP, tmpL2)

    If arcP1.Bool = False Then ArcCenterP.Bool = False
    If arcP2.Bool = False Then ArcCenterP.Bool = False

    'conpute arc "start" and "end" angles
    A1 = Atan2(arcP1.x - ArcCenterP.x, arcP1.y - ArcCenterP.y)
    A2 = Atan2(arcP2.x - ArcCenterP.x, arcP2.y - ArcCenterP.y)

    If AngleDIFF(A1, A2) > 0 Then
        retArc.A1 = A1
        retArc.A2 = A2
    Else
        retArc.A1 = A2
        retArc.A2 = A1
    End If

    If ModifyLines And ArcCenterP.Bool Then
        If DistFromPoint(arcP1, tmpL1.P1) < DistFromPoint(arcP1, tmpL1.P2) Then
            L1.P1 = arcP1
        Else
            L1.P2 = arcP1
        End If
        If DistFromPoint(arcP2, L2.P1) < DistFromPoint(arcP2, L2.P2) Then
            L2.P1 = arcP2
        Else
            L2.P2 = arcP2
        End If
    End If

    UpdateArcPts retArc

End Function



Public Function maX(A As Double, b As Double) As Double
    If A > b Then maX = A Else: maX = b
End Function

Public Function Min(A As Double, b As Double) As Double
    If A < b Then Min = A Else: Min = b
End Function


