Attribute VB_Name = "modRichClient5"
Option Explicit


' After Draw ---> REFRESH:
'vbDRAW.Srf.DrawToDC PicHDC
'DoEvents



Public Srf As cCairoSurface, CC As cCairoContext    'Srf is similar to a DIB, the derived CC similar to a hDC
Attribute CC.VB_VarUserMemId = 1073741824
Public vbDRAW  As cVBDraw
Attribute vbDRAW.VB_VarUserMemId = 1073741826
Public CONS    As cConstructor
Attribute CONS.VB_VarUserMemId = 1610809344

Public PicHDC  As Long
Attribute PicHDC.VB_VarUserMemId = 1073741828
Public MaxW    As Long
Attribute MaxW.VB_VarUserMemId = 1073741829
Public maxH    As Long
Attribute maxH.VB_VarUserMemId = 1073741830

Public CenX    As Double
Public CenY    As Double

Public wMinX   As Double
Public wMinY   As Double
Public wMaxX   As Double
Public wMaxY   As Double



Public Sub InitRC()
    Set Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
    Set CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above


    MaxW = fMain.PIC.Width
    maxH = fMain.PIC.Height

    CenX = MaxW * 0.5
    CenY = maxH * 0.5


    wMinX = CenX - MaxW * 2.2    'Must be<0
    wMinY = CenY - maxH * 2.2
    wMaxX = CenX + MaxW * 2.2
    wMaxY = CenY + maxH + 2.2




    Set vbDRAW = Cairo.CreateVBDrawingObject
    Set vbDRAW.Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels

    Set vbDRAW.CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above


    vbDRAW.BindTo fMain.PIC

    With vbDRAW

        .CC.AntiAlias = CAIRO_ANTIALIAS_GRAY

        .CC.SetSourceSurface Srf
        .CC.SetLineCap CAIRO_LINE_CAP_ROUND
        .CC.SetLineJoin CAIRO_LINE_JOIN_ROUND


        .CC.SetLineWidth 1, True


        .CC.SelectFont "Courier New", 9, vbWhite


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




Public Sub InitResources()
    Dim Srf    As cCairoSurface
    Dim CC     As cCairoContext
    Dim size   As Double
    Dim I      As Long
    Dim x      As Double
    Dim y      As Double
    Dim Gray   As Double

    '    Cairo.ImageList.AddImage "FoodIcon", App.Path & "\Resources\Orb.png", 16, 16
    Cairo.ImageList.AddImage "FoodIcon", App.Path & "\Resources\greenlight.png", FoodSize * 2, FoodSize * 2




    Gray = 45    '60

    Set Srf = New_c.Cairo.CreateSurface(wMaxX - wMinX, wMaxY - wMinY, ImageSurface)
    Set CC = Srf.CreateContext
    CC.SetSourceColor RGB(Gray * 0.8, Gray * 0.8, Gray * 0.8)
    CC.Paint

    CC.RotateDrawings PIh / 12

    size = 200    ' 140
    I = 0
    CC.SetSourceColor RGB(Gray, Gray, Gray)
    For x = 0 To Srf.Width * 1.2 Step size * Cos(Pi / 6)
        I = I + 1
        For y = -Srf.Width * 0.2 To Srf.Height * 1.2 Step size
            CC.DrawRegularPolygon x, y + (I Mod 2) * size * 0.5, size * 0.4, 6, splSmallest 'splNone
            CC.Fill
        Next
    Next
    'Srf.FastBlur
    I = 0
    '  CC.SetSourceColor RGB(74, 74, 74)
    For x = 0 To Srf.Width * 1.2 Step size * Cos(Pi / 6)
        I = I + 1
        For y = -Srf.Width * 0.2 To Srf.Height * 1.2 Step size

            CC.SetSourceColor RGB(Gray * 1.3, Gray * 1.3, Gray * 1.3)
            CC.DrawRegularPolygon x - size * 0.02, -size * 0.02 + y + (I Mod 2) * size * 0.5, size * 0.3, 6, splSmallest 'splNone
            CC.Fill
            CC.SetSourceColor RGB(Gray * 0.9, Gray * 0.9, Gray * 0.9)
            CC.DrawRegularPolygon x + size * 0.02, size * 0.02 + y + (I Mod 2) * size * 0.5, size * 0.3, 6, splSmallest 'splNone
            CC.Fill

            CC.SetSourceColor RGB(Gray * 1.16, Gray * 1.16, Gray * 1.16)
            CC.DrawRegularPolygon x, y + (I Mod 2) * size * 0.5, size * 0.3, 6, splNone
            CC.Fill



        Next
    Next

    CC.Restore

    Srf.WriteContentToJpgFile App.Path & "\Resources\BK.jpg"
    ' Cairo.ImageList.AddSurface "BK", Srf
    Cairo.ImageList.AddImage "BK", App.Path & "\Resources\BK.jpg"

    Set CC = Nothing
    Set Srf = Nothing

End Sub
