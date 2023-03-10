Attribute VB_Name = "mMain"
Option Explicit

Public ScreenH#
Public ScreenW#

Public Srf        As cCairoSurface
Public CC         As cCairoContext

Public fMainhDC   As Long

Public CamPosX    As Double
Public CamPosY    As Double

Public Tick       As clsTick

Public DoLOOP     As Boolean

Private tDRAW     As Long
Private t1Sec     As Long

Public CNT        As Long
Public oCnt       As Long


Public srfbkg     As cCairoSurface
Public srf2Screen As cCairoSurface

Public bgCX       As Double
Public bgCY       As Double





Public Sub MAINLOOP()
    Dim overlayValue As Double

    Set Tick = New clsTick

    tDRAW = Tick.Add(120)    '<---- set desired draw FPS
    t1Sec = Tick.Add(1)

    DoLOOP = True

    fMainhDC = fMain.hDC


    Do
        Select Case Tick.WaitForNext
        Case tDRAW

            SetTile2 14, CamPosX, CamPosY


            DRAWBackGround


            If (CNT \ 1300) Mod 2 = 0 Then
                CamPosX = TW * 0.5 + Cos(Timer * 0.4) * TW * 0.35
                CamPosY = TH * 0.5    ' + Cos(Timer * 0.5) * 5
            Else
                CamPosX = TW * 0.5    '+ Cos(Timer * 0.5) * 5
                CamPosY = TH * 0.5 + Cos(Timer * 0.4) * TH * 0.35
            End If

            If (CNT And 1023) = 0 Then
                SETOverlay overlayValue
                overlayValue = overlayValue + 0.01: If overlayValue > 1 Then overlayValue = overlayValue - 1
            End If

            CNT = CNT + 1

        Case t1Sec
            fMain.Caption = "Draw FPS: " & CNT - oCnt & "      (ESC and then window 'X' to quit)"
            oCnt = CNT

            DoEvents

        End Select

    Loop While DoLOOP
    Tick.RemoveByID tDRAW
    Tick.RemoveByID t1Sec

End Sub
