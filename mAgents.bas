Attribute VB_Name = "mAgents"
Option Explicit

Public Type tAgent
    X             As Double
    Y             As Double
    dirX          As Double
    dirY          As Double

    TileIdx       As Long
    DrawOrder     As Long
    XY            As Double
    Speed As Double
    
End Type

Public Agent()    As tAgent
Public NA         As Long

Private tmpCCagents As cCairoContext

Private Sub AgentRandomDir(I&)
    With Agent(I)
        .dirX = 0
        .dirY = 0

        If Rnd > 0.5 Then
            Do: .dirX = Int(Rnd * 3 - 1): Loop While .dirX = 0
        Else
            Do: .dirY = Int(Rnd * 3 - 1): Loop While .dirY = 0
        End If
    End With

End Sub
Public Sub SetUpAgents()
    Dim I&
    Dim X#, Y#
    NA = 0

    For I = 1 To 10
        Do
            X = 1 + Int(Rnd * (TW - 2)): Y = 1 + Int(Rnd * (TH - 2))
        Loop While TilesMAP(X, Y).ImgIdx <> 0
        AddAgent X, Y, 1 + Rnd * 15
        AgentRandomDir I
    Next

End Sub

Public Sub AddAgent(PosX#, PosY#, TileIdx As Long)
    NA = NA + 1
    ReDim Preserve Agent(NA)
    With Agent(NA)
        .X = PosX
        .Y = PosY
        .TileIdx = TileIdx
        .DrawOrder = NA
        If NA = 1 Then
            .Speed = 0.08
        Else
            .Speed = 0.01 + Rnd * 0.15
        End If

    End With
End Sub


Public Sub RenderAgentShadow(AgentIdx&)
'Exit Sub


    Dim fX#, fY#
    Dim tfx#, tfy#

    Dim TX#, TY#
    Dim TrX#, TrY#

    Dim iX&, iy&

    Dim Tidx      As Long
    Dim PosX#, PosY#

    Tidx = Agent(AgentIdx).TileIdx
    PosX = Agent(AgentIdx).X
    PosY = Agent(AgentIdx).Y

    TileXYtoScreen TW * 0.5, TH * 0.5, TrX, TrY

    iX = Int(PosX)
    iy = Int(PosY)
    fX = PosX - iX
    fY = PosY - iy

    TileXYtoScreen fX, fY, tfx, tfy

    TX = TilesMAP(iX, iy).scrX + TILEShad(Tidx).offX + srf2Screen.Width * 0.5 + tfx - TrX
    TY = TilesMAP(iX, iy).scrY + TILEShad(Tidx).offY + srf2Screen.Height * 0.5 + tfy - TrY

    tmpCCagents.SetSourceSurface TILEShad(Tidx).tSrf, TX, TY
    tmpCCagents.MaskSurface ShadowsMaskSrf, -TrX, -TrY
    

End Sub


Public Sub RenderAgent(AgentIdx&)

    Dim fX#, fY#
    Dim tfx#, tfy#

    Dim TX#, TY#
    Dim TrX#, TrY#

    Dim iX&, iy&

    Dim Tidx      As Long
    Dim PosX#, PosY#

    Dim Diag      As Long


    Tidx = Agent(AgentIdx).TileIdx
    PosX = Agent(AgentIdx).X
    PosY = Agent(AgentIdx).Y

    TileXYtoScreen TW * 0.5, TH * 0.5, TrX, TrY

    iX = Int(PosX)
    iy = Int(PosY)
    fX = PosX - iX
    fY = PosY - iy

    TileXYtoScreen fX, fY, tfx, tfy

    TX = TilesMAP(iX, iy).scrX + TILE(Tidx).offX + srf2Screen.Width * 0.5 + tfx - TrX
    TY = TilesMAP(iX, iy).scrY + TILE(Tidx).offY + srf2Screen.Height * 0.5 + tfy - TrY

    Diag = PosX + PosY + 0.5

    tmpCCagents.SetSourceSurface TILE(Tidx).tSrf, TX, TY
    tmpCCagents.MaskSurface MASKSRF(Diag), -TrX + MaskSrfOffX(Diag), -TrY + MaskSrfOffY(Diag)

    ' TEST with MASKSRF with Standard Alpha
    '     tmpCCagents.RenderSurfaceContent TILE(Tidx).tSrf, TX, TY
    '     tmpCCagents.RenderSurfaceContent MASKSRF(iX + iy + 2), -TrX, -TrY
End Sub


Public Sub RENDERallAgents()
    Dim I         As Long

    Set tmpCCagents = srf2Screen.CreateContext
    tmpCCagents.RenderSurfaceContent srfbkg, 0, 0    '<<<<----------!!!!!!!!!!!

    For I = 1 To NA
        Agent(I).XY = -Agent(I).X - Agent(I).Y
    If DynObjShad Then RenderAgentShadow I
    Next
    QuickSortAgent Agent(), 1, NA

    For I = 1 To NA
        RenderAgent Agent(I).DrawOrder
    Next
End Sub



Public Sub MOVEAGENTS()
    Dim I         As Long
    Dim toX&, toY&
    For I = 1 To NA
        With Agent(I)

            toX = .X + .dirX
            toY = .Y + .dirY

            If toX < 0 Or toX > TW Then .dirX = -.dirX
            If toY < 0 Or toY > TH Then .dirY = -.dirY

            toX = .X + .dirX * 0.5
            toY = .Y + .dirY * 0.5

            If TilesMAP(toX, toY).ImgIdx Then
                .X = .X - .dirX * .Speed
                .Y = .Y - .dirY * .Speed
                AgentRandomDir I
            End If

            .X = .X + .dirX * .Speed
            .Y = .Y + .dirY * .Speed


        End With
    Next
End Sub


Private Sub QuickSortAgent(List() As tAgent, ByVal min As Long, ByVal max As Long)
' FROM HI to LOW  'https://www.vbforums.com/showthread.php?11192-quicksort
    Dim Low As Long, high As Long, temp As tAgent, TestElement As Double
    Dim DD&
    Low = min: high = max
    '    TestElement = List((min + max) / 2).XY
    TestElement = List(List((min + max) / 2).DrawOrder).XY
    Do
        Do While List(List(Low).DrawOrder).XY > TestElement: Low = Low + 1&: Loop
        Do While List(List(high).DrawOrder).XY < TestElement: high = high - 1&: Loop
        If (Low <= high) Then
            'temp = List(Low): List(Low) = List(high): List(high) = temp
            DD = List(Low).DrawOrder
            List(Low).DrawOrder = List(high).DrawOrder
            List(high).DrawOrder = DD
            Low = Low + 1&: high = high - 1&
        End If
    Loop While (Low <= high)
    If (min < high) Then QuickSortAgent List, min, high
    If (Low < max) Then QuickSortAgent List, Low, max
End Sub
