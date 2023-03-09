Attribute VB_Name = "mTiles"
Option Explicit
'Print_X & Print_Y are the Physical screen Printing Locations.
'Little x & y are the Map Coords.
''----------------------------------------------
'X_Tile_Width = 50
'Y_Tile_Height = 25
'X_Increment = 31 'reverse increment to the left per each row
'Y_Increment = 11 'the drop per each increment in x
''Calculate Print_X and Y
'Print_X = X_Tile_Width * X - Y * X_Increment
'Print_Y = Y_Tile_Height* Y + X * Y_Increment


'     D
'   /  \
' /      A
'C      /
'  \  /
'   B


Public Const tileW As Long = 50
Public Const tileH As Long = 25
Public Const XIncrement As Long = 31    '31
Public Const YIncrement As Long = 11    '11

Private NtilesImg As Long
Public Type tTile
    H             As Double

    ImgIdx        As Long
    '    ImgKeySha     As String 'Shadow

    scrX          As Double
    scrY          As Double
End Type

Public Type tTileImg
    tSrf          As cCairoSurface
    offX          As Double
    offY          As Double
End Type

Public TilesMAP() As tTile

Public TILE()     As tTileImg


Public TW         As Long
Public TH         As Long

Private BYTESBackgr() As Byte
Private BYTESScreen() As Byte

Private ZETA      As Long
Private Z()       As Long


Public Sub TileXYtoScreen(X#, Y#, scrX#, scrY#)
'    scrX = (X + Y) * (tileW * 0.5)
'    scrY = (Y - X) * (tileH * 0.5)

    scrX = tileW * X - Y * XIncrement
    scrY = tileH * Y + X * YIncrement

    '    scrX = tileW * X + Y * XIncrement
    '    scrY = -tileH * Y + X * YIncrement

End Sub

Public Sub InitTiles()
    Dim X&, Y&

    Dim ceX       As Double
    Dim ceY       As Double
    Dim Ax#, Bx#, cX#, DX#
    Dim Ay#, By#, cY#, DY#

    Dim cR#, cG#, cB#
    Dim H#

    Dim tmpSrf    As cCairoSurface
    Dim tmpCC     As cCairoContext


    TW = 20                  '''''' TILEMAP Size -.<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    TH = 20

    Set tmpSrf = Cairo.CreateSurface(81, 36 * 4, ImageSurface)
    Set tmpCC = tmpSrf.CreateContext

    NtilesImg = 15
    ReDim TILE(NtilesImg)

    TileXYtoScreen -0.5, -0.5, DX, DY
    TileXYtoScreen -0.5, 0.5, Ax, Ay
    TileXYtoScreen 0.5, 0.5, Bx, By
    TileXYtoScreen 0.5, -0.5, cX, cY

    ceX = 40.5
    ceY = tmpSrf.Height - 18

    Ax = Ax + ceX: Bx = Bx + ceX: cX = cX + ceX: DX = DX + ceX
    Ay = Ay + ceY: By = By + ceY: cY = cY + ceY: DY = DY + ceY


    For X = 0 To NtilesImg
        Set tmpSrf = Cairo.CreateSurface(81, 36 * 4, ImageSurface)
        Set tmpCC = tmpSrf.CreateContext


        H = (X / NtilesImg) * 36 * 3

        cR = Rnd: cG = Rnd: cB = Rnd
        tmpCC.MoveTo Ax, Ay - H
        tmpCC.LineTo Bx, By - H
        tmpCC.LineTo cX, cY - H
        tmpCC.LineTo DX, DY - H
        tmpCC.SetSourceRGB cR, cG, cB
        tmpCC.Fill

        If H Then
            tmpCC.MoveTo Bx, By - H
            tmpCC.LineTo Ax, Ay - H
            tmpCC.LineTo Ax, Ay
            tmpCC.LineTo Bx, By
            tmpCC.SetSourceRGB cR * 0.8, cG * 0.8, cB * 0.8
            tmpCC.Fill

            tmpCC.MoveTo cX, cY - H
            tmpCC.LineTo cX, cY
            tmpCC.LineTo Bx, By
            tmpCC.LineTo Bx, By - H
            tmpCC.SetSourceRGB cR * 0.4, cG * 0.4, cB * 0.4
            tmpCC.Fill
        End If

        '        Set TILE(X).tmpSrf = tmpSrf.CreateSimilar(CAIRO_CONTENT_COLOR_ALPHA, , , True)'<<< do not work....

        Set TILE(X).tSrf = Cairo.CreateSurface(81, 36 * 4, ImageSurface)
        TILE(X).tSrf.CreateContext.RenderSurfaceContent tmpSrf, 0, 0

        TILE(X).offX = -cX
        TILE(X).offY = -cY


        CC.RenderSurfaceContent TILE(X).tSrf, 0, 0    'ceX, ceY
        Srf.DrawToDC fMain.hDC


    Next


    '''        For X = 0 To 100
    '''            CC.RenderSurfaceContent TILE(Rnd * NtilesImg).tmpSrf, 800 * Rnd, 600 * Rnd
    '''        Next
    '''        Srf.DrawToDC fMain.hDC


    LoadTile App.Path & "\PNG\tree_PNG212.png"
    LoadTile App.Path & "\PNG\tree_PNG3470.png"
    LoadTile App.Path & "\PNG\tree_PNG3477.png"


    CC.RenderSurfaceContent TILE(UBound(TILE)).tSrf, 0, 0    'ceX, ceY
    Srf.DrawToDC fMain.hDC


    ReDim TilesMAP(TW, TH)
    For X = 0 To TW
        For Y = 0 To TH
            With TilesMAP(X, Y)
                .ImgIdx = Int(Rnd * (NtilesImg + 1))
                If Rnd < 0.76 Then .ImgIdx = 0
                TileXYtoScreen X * 1, Y * 1, .scrX, .scrY
            End With
        Next
    Next

    CamPosX = TW * 0.5
    CamPosY = TH * 0.5
    '

    TileXYtoScreen -0.5, -0.5, DX, DY
    TileXYtoScreen 0.5 + TW, -0.5, Ax, Ay
    TileXYtoScreen 0.5 + TW, 0.5 + TH, Bx, By
    TileXYtoScreen -0.5, 0.5 + TH, cX, cY


    Ax = Ax + 100: cX = cX - 100
    By = By + 100: DY = DY - 100
    Set srfbkg = Cairo.CreateSurface(Ax - cX, By - DY, ImageSurface)
    Set srf2Screen = Cairo.CreateSurface(srfbkg.Width, srfbkg.Height, ImageSurface)
    ReDim Z(srfbkg.Width, srfbkg.Height)

    SetupBACKGROUND



    srf2Screen.CreateContext.RenderSurfaceContent srfbkg, 0, 0




    bgCX = srfbkg.Width * 0.5
    bgCY = srfbkg.Height * 0.5


End Sub






Public Sub SetupBACKGROUND()

    Dim X         As Long
    Dim Y         As Long
    Dim TX        As Double
    Dim TY        As Double
    Dim Idx       As Long


    Dim tmpCC     As cCairoContext

    Set tmpCC = srfbkg.CreateContext

    tmpCC.Save

    srfbkg.BindToArray BYTESBackgr
    srf2Screen.BindToArray BYTESScreen


    TileXYtoScreen CamPosX, CamPosY, TX, TY

    tmpCC.TranslateDrawings -TX, -TY

    tmpCC.SetSourceColor 0: tmpCC.Paint
    tmpCC.SelectFont "Courier New", 8, vbGreen
    tmpCC.SetSourceColor 255

    ' FLOOR

    For X = 0 To TW
        For Y = 0 To TH
            Idx = 0
            TX = TilesMAP(X, Y).scrX + TILE(Idx).offX + srfbkg.Width * 0.5
            TY = TilesMAP(X, Y).scrY + TILE(Idx).offY + srfbkg.Height * 0.5
            tmpCC.RenderSurfaceContent TILE(Idx).tSrf, TX, TY
        Next
    Next

    ' other TilesMAP
    ZETA = 0
    For X = 0 To TW
        For Y = 0 To TH


            Idx = TilesMAP(X, Y).ImgIdx
            If Idx Then
                TX = TilesMAP(X, Y).scrX + TILE(Idx).offX + srfbkg.Width * 0.5
                TY = TilesMAP(X, Y).scrY + TILE(Idx).offY + srfbkg.Height * 0.5
                tmpCC.RenderSurfaceContent TILE(Idx).tSrf, TX, TY
                UPDATEZ      '********************************************************************************
                srf2Screen.CreateContext.RenderSurfaceContent srfbkg, 0, 0
            End If

            ZETA = X + Y

        Next
    Next



    tmpCC.Restore
    srfbkg.DrawToDC fMainhDC

    DoEvents
End Sub
Private Sub UPDATEZ()

    Dim X&, Y&


    For X = 0 To UBound(BYTESBackgr, 1) Step 4
        For Y = 0 To UBound(BYTESBackgr, 2) Step 1
            If BYTESBackgr(X + 0, Y) <> BYTESScreen(X + 0, Y) Then Z(X \ 4, Y) = ZETA
            If BYTESBackgr(X + 1, Y) <> BYTESScreen(X + 1, Y) Then Z(X \ 4, Y) = ZETA
            If BYTESBackgr(X + 2, Y) <> BYTESScreen(X + 2, Y) Then Z(X \ 4, Y) = ZETA
        Next
    Next
    DoEvents

End Sub






Private Sub LoadTile(FN As String)
    Dim W#, H#
    Dim TS        As cCairoSurface

    NtilesImg = NtilesImg + 1
    ReDim Preserve TILE(NtilesImg)

    With TILE(NtilesImg)

        Set TS = Cairo.ImageList.AddImage("tmp", FN)
        W = TS.Width
        H = TS.Height

        If H > W Then
            H = 200
            W = H * TS.Width / TS.Height
        Else
            W = 200
            H = W * TS.Height / TS.Width
        End If

        Set .tSrf = Cairo.ImageList.AddImage("tmp2", FN, W, H)
        .offX = -W * 0.5 - 40.5
        .offY = -H + 18 - 9
    End With

    Cairo.ImageList.RemoveAll


End Sub


Public Sub DRAWBackGround()
    Dim TX#, TY#

    TileXYtoScreen CamPosX - TW * 0.5, CamPosY - TH * 0.5, TX, TY
    '    TX = 0
    '    TY = 0

    CC.RenderSurfaceContent srf2Screen, ScreenW * 0.5 - bgCX - TX, ScreenH * 0.5 - bgCY - TY

    CC.Arc ScreenW * 0.5, ScreenH * 0.5, 11: CC.Fill

    Srf.DrawToDC fMainhDC
    DoEvents

End Sub

Private Function CutFrom(SrcSrf As cCairoSurface, _
                         X As Double, Y As Double, DX As Long, DY As Long) As cCairoSurface
    Set CutFrom = Cairo.CreateSurface(DX, DY)
    CutFrom.CreateContext.RenderSurfaceContent SrcSrf, -X, -Y
End Function

Public Sub SetTile(Tidx&, pX#, pY#)

    Dim fx#, fy#
    Dim tfx#, tfy#

    Dim tmpCC     As cCairoContext

    Dim TX#, TY#
    Dim TrX#, TrY#

    Dim iX&, iy&
    Dim X&, Y&
    Dim Idx&
    Dim Cut       As cCairoSurface
    Dim XX&, YY&

    iX = Int(pX)
    iy = Int(pY)

    Set tmpCC = srf2Screen.CreateContext

    tmpCC.Save


    TileXYtoScreen TW * 0.5, TH * 0.5, TrX, TrY
    tmpCC.TranslateDrawings -TrX, -TrY

    fx = pX - iX
    fy = pY - iy

    TileXYtoScreen fx, fy, tfx, tfy


    '    TX = TilesMAP(iX, iY).scrX + TILE(Tidx).offX + srf2Screen.Width * 0.5
    '    TY = TilesMAP(iX, iY).scrY + TILE(Tidx).offY + srf2Screen.Height * 0.5
    '    Set Cut = CutFrom(srfBKG, TX - 50 + TrX, TY - 50 + TrY, 150, 150)
    '    tmpCC.RenderSurfaceContent Cut, TX - 50, TY - 50

    ''    floor
    '    For Y = TH To 0 Step -1
    '        For X = 0 To TW Step 1
    For Y = iy + 3 To iy - 3 Step -1
        For X = iX - 3 To iX + 3

            Idx = 0
            TX = TilesMAP(X, Y).scrX + TILE(Idx).offX + srf2Screen.Width * 0.5
            TY = TilesMAP(X, Y).scrY + TILE(Idx).offY + srf2Screen.Height * 0.5
            tmpCC.RenderSurfaceContent TILE(Idx).tSrf, TX, TY
        Next
    Next

    ''    ' other TilesMAP
    '    For Y = TH To 0 Step -1
    '        For X = 0 To TW Step 1
    For Y = iy + 3 To iy - 3 Step -1
        For X = iX - 3 To iX + 3
            Idx = TilesMAP(X, Y).ImgIdx

            If X = iX And Y = iy Then
                Idx = Tidx
                TX = TilesMAP(X, Y).scrX + TILE(Idx).offX + srf2Screen.Width * 0.5
                TY = TilesMAP(X, Y).scrY + TILE(Idx).offY + srf2Screen.Height * 0.5
                tmpCC.RenderSurfaceContent TILE(Idx).tSrf, TX + tfx, TY + tfy
            Else

                If Idx Then
                    TX = TilesMAP(X, Y).scrX + TILE(Idx).offX + srf2Screen.Width * 0.5
                    TY = TilesMAP(X, Y).scrY + TILE(Idx).offY + srf2Screen.Height * 0.5
                    tmpCC.RenderSurfaceContent TILE(Idx).tSrf, TX, TY
                End If
            End If
        Next
    Next



    tmpCC.Restore



End Sub


Public Sub SetTile2(Tidx&, pX#, pY#)
    Dim fx#, fy#
    Dim tfx#, tfy#

    Dim tmpCC     As cCairoContext

    Dim TX#, TY#
    Dim TrX#, TrY#

    Dim iX&, iy&
    Dim X&, Y&
    Dim X4&, TX4&
    Dim BYTESTile() As Byte

    Set tmpCC = srf2Screen.CreateContext

    TileXYtoScreen TW * 0.5, TH * 0.5, TrX, TrY

    iX = Int(pX)
    iy = Int(pY)
    fx = pX - iX
    fy = pY - iy

    TileXYtoScreen fx, fy, tfx, tfy



    TX = TilesMAP(iX, iy).scrX + TILE(Tidx).offX + srf2Screen.Width * 0.5 + tfx - TrX
    TY = TilesMAP(iX, iy).scrY + TILE(Tidx).offY + srf2Screen.Height * 0.5 + tfy - TrY


    ZETA = Z(TX, TY)


    TX = TilesMAP(iX, iy).scrX + TILE(Tidx).offX + srf2Screen.Width * 0.5 + tfx - TrX
    TY = TilesMAP(iX, iy).scrY + TILE(Tidx).offY + srf2Screen.Height * 0.5 + tfy - TrY

    '        tmpCC.RenderSurfaceContent TILE(Tidx).tSrf, TX, TY


    TX4 = Int(TX) * 4

    For X = -12 To TILE(Tidx).tSrf.Width + 12
        X4 = X * 4
        For Y = -12 To TILE(Tidx).tSrf.Height + 12
            BYTESScreen(X4 + 0 + TX4, TY + Y) = BYTESBackgr(X4 + 0 + TX4, TY + Y)
            BYTESScreen(X4 + 1 + TX4, TY + Y) = BYTESBackgr(X4 + 1 + TX4, TY + Y)
            BYTESScreen(X4 + 2 + TX4, TY + Y) = BYTESBackgr(X4 + 2 + TX4, TY + Y)
            BYTESScreen(X4 + 3 + TX4, TY + Y) = 255    'BYTESbackgr(X4 + 3 + TX4, TY + Y)
        Next
    Next


    '--------------------------------------------
    TILE(Tidx).tSrf.BindToArray BYTESTile

    For X = 0 To TILE(Tidx).tSrf.Width - 1
        X4 = X * 4
        For Y = 0 To TILE(Tidx).tSrf.Height - 1
            If Z(X + TX, Y + TY) <= ZETA Then
                If BYTESTile(X4 + 3, Y) Then
                    BYTESScreen(X4 + 0 + TX4, TY + Y) = BYTESTile(X4 + 0, Y)
                    BYTESScreen(X4 + 1 + TX4, TY + Y) = BYTESTile(X4 + 1, Y)
                    BYTESScreen(X4 + 2 + TX4, TY + Y) = BYTESTile(X4 + 2, Y)
                End If
            End If
        Next
    Next

    TILE(Tidx).tSrf.ReleaseArray BYTESTile
    '--------------------------------------------

    tmpCC.Arc TX, TY, 3
    tmpCC.Fill


End Sub



Public Sub SETOverlay(V As Double)
    Const Inv255  As Double = 1 / 255
    Const PI2     As Double = 6.28318530717959


    Dim X         As Long
    Dim Y         As Long

    Dim iR#, iG#, iB#
    Dim bR#, bG#, bB#
    Dim rR#, rG#, rB#


    bR = 0.5 + 0.5 * Cos(V * 2 * PI2)
    bG = 0.5 + 0.25 * Cos(V * 1 * PI2)
    bB = 0.5 + 0.5 * Cos(V * PI2)


    Dim Bytes()   As Byte


    srf2Screen.CreateContext.RenderSurfaceContent srfbkg, 0, 0

    srf2Screen.BindToArray Bytes

    For X = 0 To UBound(Bytes, 1) - 3 Step 4
        For Y = 0 To UBound(Bytes, 2)
            iB = Bytes(X + 0, Y) * Inv255
            iG = Bytes(X + 1, Y) * Inv255
            iR = Bytes(X + 2, Y) * Inv255

            rR = BlendOverlay(iR, bR)
            rG = BlendOverlay(iG, bB)
            rB = BlendOverlay(iB, bB)

            Bytes(X + 0, Y) = rB * 255
            Bytes(X + 1, Y) = rG * 255
            Bytes(X + 2, Y) = rR * 255


        Next
    Next


    srf2Screen.ReleaseArray Bytes



End Sub

Public Function BlendOverlay(ByVal base As Double, ByVal blend As Double) As Double
    If base < 0.5 Then
        BlendOverlay = 2# * base * blend
    Else
        BlendOverlay = 1# - (2# * (1# - base) * (1# - blend))
    End If
End Function
