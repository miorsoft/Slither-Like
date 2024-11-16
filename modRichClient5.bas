Attribute VB_Name = "modRichClient5"
Option Explicit


' After Draw ---> REFRESH:
'vbDRAW.Srf.DrawToDC PicHDC
'DoEvents



Public Srf As cCairoSurface, CC As cCairoContext    'Srf is similar to a DIB, the derived CC similar to a hDC

Public vbDRAW As cVBDraw
Public vbDrawCC As cCairoContext

Public CONS As cConstructor
Attribute CONS.VB_VarUserMemId = 1610809344

Public PicHDC As Long
Attribute PicHDC.VB_VarUserMemId = 1073741828
Public MaxW As Long
Attribute MaxW.VB_VarUserMemId = 1073741829
Public maxH As Long
Attribute maxH.VB_VarUserMemId = 1073741830

Public CenX As Double
Public CenY As Double

Public wMinX As Double
Public wMinY As Double
Public wMaxX As Double
Public wMaxY As Double


Public IconR() As Double
Public IconG() As Double
Public IconB() As Double

Public Sub InitRC()
   ' Set Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
   ' Set CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above


    MaxW = fMain.PIC.Width
    maxH = fMain.PIC.Height

    CenX = MaxW * 0.5
    CenY = maxH * 0.5


'    wMinX = CenX - MaxW * 2.2    'Must be<0
'    wMinY = CenY - maxH * 2.2
'    wMaxX = CenX + MaxW * 2.2
'    wMaxY = CenY + maxH + 2.2

wMinX = CenX - 840 * 3.5 '2
wMaxX = CenX + 840 * 3.5
wMinY = CenY - 640 * 3.5
wMaxY = CenY + 640 * 3.5




    Set vbDRAW = Cairo.CreateVBDrawingObject
'    Set vbDRAW.Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
    Set vbDRAW.Srf = Cairo.CreateSurface(fMain.PIC.Width, fMain.PIC.Height, ImageSurface)       'size of our rendering-area in Pixels

    Set vbDrawCC = vbDRAW.Srf.CreateContext    'create a Drawing-Context from the PixelSurface above


    'vbDRAW.BindTo fMain.PIC


    With vbDrawCC

        '.AntiAlias = CAIRO_ANTIALIAS_GRAY
        .AntiAlias = CAIRO_ANTIALIAS_FAST
        
        '.CC.SetSourceSurface Srf
        .SetLineCap CAIRO_LINE_CAP_ROUND
        .SetLineJoin CAIRO_LINE_JOIN_ROUND


        
        .SetLineWidth 2, True


        .SelectFont "Courier New", 9, vbWhite


    End With

    PicHDC = fMain.PIC.hDC

    '    fMain.PIC.Cls
    '    fMain.PIC.Height = 640    '480    '360    ' 480
    '    fMain.PIC.Width = Int(fMain.PIC.Height * 4 / 3)


End Sub

Public Sub UnloadRC()
    Set CC = Nothing
    Set Srf = Nothing
    Set vbDRAW = Nothing

    Set CONS = New cConstructor
    CONS.CleanupRichClientDll
End Sub


Public Sub InitFoodIcon(ByVal BySnake As Long, ByVal FoodNum As Long)
    Dim Srf       As cCairoSurface
    Dim CC        As cCairoContext
    Dim b()       As Byte
    Dim x         As Long
    Dim y         As Long

    Dim cR        As Double
    Dim cG        As Double
    Dim cB        As Double

    If BySnake > UBound(IconR) Then
        ReDim Preserve IconR(-1 To BySnake)
        ReDim Preserve IconG(-1 To BySnake)
        ReDim Preserve IconB(-1 To BySnake)
    End If
    '    Set Srf = Cairo.CreateSurface(Cairo.ImageList("FoodIcon").Width, _
         '                                  Cairo.ImageList("FoodIcon").Height, ImageSurface)
    '    Set CC = Srf.CreateContext
    '    CC.RenderSurfaceContent "FoodIcon", 0, 0

    Set Srf = Cairo.ImageList("FoodIcon").CreateSimilar(CAIRO_CONTENT_COLOR_ALPHA, , , True)

    Srf.BindToArray b

    If BySnake >= 0 Then
        cR = Snake(BySnake).ColorR
        cG = Snake(BySnake).ColorG
        cB = Snake(BySnake).ColorB

        If (IconR(BySnake) <> cR) Or _
           (IconG(BySnake) <> cG) Or _
           (IconB(BySnake) <> cB) Then
            For y = 0 To UBound(b, 2)
                For x = 0 To UBound(b, 1) Step 4
                    b(x + 0, y) = cB * b(x + 3, y)
                    b(x + 1, y) = cG * b(x + 3, y)
                    b(x + 2, y) = cR * b(x + 3, y)
                    b(x + 3, y) = 0
                Next
            Next
            Cairo.ImageList.AddImage "FoodIcon" & CStr(BySnake), b()

            IconR(BySnake) = Snake(BySnake).ColorR
            IconG(BySnake) = Snake(BySnake).ColorG
            IconB(BySnake) = Snake(BySnake).ColorB

        End If

    Else
        If Not (Cairo.ImageList.Exists("FoodIcon" & CStr(BySnake))) Then
            Cairo.ImageList.AddImage "FoodIcon" & CStr(BySnake), b()
        End If
    End If



    Srf.ReleaseArray b()

    Set CC = Nothing
    Set Srf = Nothing

End Sub

Public Sub InitResources()
    Dim Srf    As cCairoSurface
    Dim CC     As cCairoContext
    Dim size   As Double
    Dim I      As Long
    Dim x      As Double
    Dim y      As Double
    Dim Gray   As Double
    Dim RR#, GG#, BB#
    
    
    ReDim IconR(-1 To 0)
    ReDim IconG(-1 To 0)
    ReDim IconB(-1 To 0)
    

    Const LowResScale As Double = 0.25 ' 0.33


    '    Cairo.ImageList.AddImage "FoodIcon", App.Path & "\Resources\Orb.png", 16, 16
    Cairo.ImageList.AddImage "FoodIcon", App.Path & "\Resources\greenlight.png", FoodSize * 2, FoodSize * 2
    Cairo.ImageList.AddImage "FoodIconLight", App.Path & "\Resources\whitelight.png", FoodSize * 4, FoodSize * 4




    Gray = 40 / 255  '45    '60

    'Set Srf = New_c.Cairo.CreateSurface(wMaxX - wMinX, wMaxY - wMinY, ImageSurface)
    'Lower Res
    Set Srf = New_c.Cairo.CreateSurface((wMaxX - wMinX) * LowResScale, (wMaxY - wMinY) * LowResScale, ImageSurface)

    Set CC = Srf.CreateContext
    CC.SetSourceRGB Gray * 0.78, Gray * 0.78, Gray * 0.78
    CC.Paint

    CC.RotateDrawings PIh / 12

    size = 200 * LowResScale    '


    I = 0
    CC.SetSourceRGB Gray, Gray, Gray
    For x = 0 To Srf.Width * 1.2 Step size * Cos(Pi / 6)
        I = I + 1
        For y = -Srf.Width * 0.2 To Srf.Height * 1.2 Step size
            CC.DrawRegularPolygon x, y + (I Mod 2) * size * 0.5, size * 0.4, 6, splSmallest    'splNone
            CC.Fill
        Next
    Next
    'Srf.FastBlur
    I = 0
    '  CC.SetSourceColor RGB(74, 74, 74)
    For x = 0 To Srf.Width * 1.2 Step size * Cos(Pi / 6)
        I = I + 1
        For y = -Srf.Width * 0.2 To Srf.Height * 1.2 Step size
        

            CC.SetSourceRGB Gray * 1.3, Gray * 1.3, Gray * 1.3
            CC.DrawRegularPolygon x - size * 0.02, -size * 0.02 + y + (I Mod 2) * size * 0.5, size * 0.3, 6, splSmallest    'splNone
            CC.Fill
            CC.SetSourceRGB Gray * 0.9, Gray * 0.9, Gray * 0.9
            CC.DrawRegularPolygon x + size * 0.02, size * 0.02 + y + (I Mod 2) * size * 0.5, size * 0.3, 6, splSmallest    'splNone
            CC.Fill

RR = (RndM * 2 - 1) * 0.03
GG = (RndM * 2 - 1) * 0.03
BB = (RndM * 2 - 1) * 0.03
            CC.SetSourceRGB Gray * 1.16 + RR, Gray * 1.16 + GG, Gray * 1.16 + BB
            CC.DrawRegularPolygon x, y + (I Mod 2) * size * 0.5, size * 0.3, 6, splNone
            CC.Fill

        Next
    Next
    '-----------------------------------------

    CC.Restore
    
    Dim b()    As Byte
    Srf.BindToArray b()
    Dim XX#, YY#
    Dim dx#, dy#
    Dim D#, X4&
    YY = UBound(b, 2)
    XX = UBound(b, 1) \ 4
    For y = 0 To YY
        dy = Abs(2 * (y - YY * 0.5) / YY)
        For x = 0 To XX
        X4 = x * 4
            dx = Abs(2 * (x - XX * 0.5) / XX)
            D = 0.998 - maX(dx, dy) ^ 8
            If D < 0 Then D = 0
            b(X4 + 0, y) = b(X4 + 0, y) * D
            b(X4 + 1, y) = b(X4 + 1, y) * D
            b(X4 + 2, y) = b(X4 + 2, y) * D
        Next
    Next


    Cairo.ImageList.AddImage "BK", b()
    Srf.ReleaseArray b()
    '-----------------------------------------

    '    Srf.WriteContentToJpgFile App.Path & "\Resources\BK.jpg"
    '    ''' Cairo.ImageList.AddSurface "BK", Srf
    '    Cairo.ImageList.AddImage "BK", App.Path & "\Resources\BK.jpg"
    '-----------------------------------------

    Set CC = Nothing
    Set Srf = Nothing

End Sub
