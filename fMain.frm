VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "-"
   ClientHeight    =   7425
   ClientLeft      =   4155
   ClientTop       =   2115
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub Form_Click()

    SETUPALL
    '


    Dim X#, Y#
    TileXYtoScreen -0.5, -0.5, X, Y: Debug.Print X, Y    'D
    TileXYtoScreen 0.5, -0.5, X, Y: Debug.Print X, Y    'A
    TileXYtoScreen 0.5, 0.5, X, Y: Debug.Print X, Y    'B
    TileXYtoScreen -0.5, 0.5, X, Y: Debug.Print X, Y    'C
    '     D
    '   /  \
    ' /      A
    'C      /
    '  \  /
    '   B

    'AC = 40.5 * 2 = 81
    'BD = 18 * 2 = 36

End Sub

Private Sub Form_Load()
'    DynObjShad = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DoLOOP = False
End Sub


Private Sub Form_Resize()

    ScreenW = Me.ScaleWidth
    ScreenH = Me.ScaleHeight

    Set Srf = Cairo.CreateSurface(ScreenW, ScreenH)
    Set CC = Srf.CreateContext

    Randomize Timer

    If Not DoLOOP Then
        SETUPALL
        MAINLOOP

    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then DoLOOP = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
