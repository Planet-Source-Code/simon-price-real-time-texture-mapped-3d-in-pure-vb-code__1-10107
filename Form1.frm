VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Real Time 3D in VB! - by Simon Price"
   ClientHeight    =   3588
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7092
   DrawWidth       =   10
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FPStimer 
      Interval        =   1000
      Left            =   2280
      Top             =   0
   End
   Begin VB.Label Label5 
      Caption         =   "SPACEBAR = Change Rendering Mode"
      Height          =   612
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   3732
   End
   Begin VB.Label Label4 
      Caption         =   "S = Stop Rotation"
      Height          =   252
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   3732
   End
   Begin VB.Label Label3 
      Caption         =   "Arrow Keys / PageUp / PageDown = Rotate Cube"
      Height          =   492
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   3612
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Controls :"
      Height          =   240
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   708
   End
   Begin VB.Label RML 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rendering Mode = Affine Multiple Texture Mapped Polygons"
      Height          =   492
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   3732
   End
   Begin VB.Label FPSlabel 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FPS"
      Height          =   240
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Frames Per Second (FPS) ="
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2052
   End
   Begin VB.Image TextureImage 
      Height          =   2400
      Left            =   360
      Top             =   720
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Image ViewImage 
      Height          =   3000
      Left            =   120
      Top             =   480
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type t3Dvector
  x As Single
  y As Single
  z As Single
  x2 As Integer
  y2 As Integer
End Type

Dim Corner(1 To 8) As t3Dvector

Private Type tPolygon
  P(1 To 4) As Byte
  Color As Byte
End Type

Dim Poly(1 To 6) As tPolygon

Dim FPS As Integer

Dim RotX, RotY, RotZ As Single

Const XX = 6
Const YY = 25

Const SIZE = 250
Const HALFSIZE = SIZE \ 2

Private Type tEdgeBuffer
  x As Byte
  u As Byte
  v As Byte
End Type

Dim LBuff(SIZE) As tEdgeBuffer
Dim RBuff(SIZE) As tEdgeBuffer

Dim RenderMode As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const FORCE = 0.03
Select Case KeyCode

  Case vbKeyUp
    RotX = RotX + FORCE
  Case vbKeyDown
    RotX = RotX - FORCE
    
  Case vbKeyRight
    RotY = RotY + FORCE
  Case vbKeyLeft
    RotY = RotY - FORCE
    
  Case vbKeyPageUp
    RotZ = RotZ + FORCE
  Case vbKeyPageDown
    RotZ = RotZ - FORCE
    
  Case vbKeyS
    RotX = 0
    RotY = 0
    RotZ = 0
    
  Case vbKeySpace
        If RenderMode = 5 Then
            RenderMode = 0
        Else
            RenderMode = RenderMode + 1
        End If
        RML = "Rendering Mode = "
        Select Case RenderMode
            Case 0
                RML = RML & "Corners Only"
                LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
            Case 1
                RML = RML & "Wire Frame"
                LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
            Case 2
                RML = RML & "Outline"
                LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
            Case 3
                RML = RML & "Flat Colour Filled Polygons"
                LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
            Case 4
                RML = RML & "Affine Single Texture Mapped Polygons"
                LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
            Case 5
                RML = RML & "Affine Multiple Texture Mapped Polygons"
                LoadPicArray2D TextureImage, App.Path & "\car1.bmp", TextureSA, TextureBMP, TextureData()
         End Select
         
  Case vbKeyEscape
    Unload Me
    
End Select
End Sub

Private Sub Form_Load()
Show
RenderMode = 5
LoadPicArray2D ViewImage, App.Path & "\blank.bmp", ViewSA, ViewBMP, ViewData()
LoadPicArray2D TextureImage, App.Path & "\texture.bmp", TextureSA, TextureBMP, TextureData()
CreateCube
For y = 0 To SIZE
    RBuff(y).x = 0
    LBuff(y).x = SIZE
Next
Dim x As Byte
For x = 0 To 5
  Poly(x + 1).Color = TextureData(x, 199)
Next
MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
PicArrayKill ViewData()
PicArrayKill TextureData()
MsgBox "This 3D program was made by Simon Price using Visual Basic 6 only! NO DirectX, NO OpenGL, NO non-VB controls, NO DLL's, JUST PURE VB CODE! It doesn't even use any API's in the main loop (API's are only used during loading!). Now if you think that's a cool acheivement, please vote for it!"
End
End Sub

Sub CreateCube()
Const S = 2500

Corner(1).x = S
Corner(1).y = S
Corner(1).z = S
Corner(2).x = -S
Corner(2).y = S
Corner(2).z = S
Corner(3).x = -S
Corner(3).y = -S
Corner(3).z = S
Corner(4).x = S
Corner(4).y = -S
Corner(4).z = S
Corner(5).x = S
Corner(5).y = S
Corner(5).z = -S
Corner(6).x = -S
Corner(6).y = S
Corner(6).z = -S
Corner(7).x = -S
Corner(7).y = -S
Corner(7).z = -S
Corner(8).x = S
Corner(8).y = -S
Corner(8).z = -S

Poly(1).P(1) = 1
Poly(1).P(2) = 2
Poly(1).P(3) = 3
Poly(1).P(4) = 4
Poly(2).P(1) = 6
Poly(2).P(2) = 5
Poly(2).P(3) = 8
Poly(2).P(4) = 7
Poly(3).P(1) = 2
Poly(3).P(2) = 6
Poly(3).P(3) = 7
Poly(3).P(4) = 3
Poly(4).P(1) = 5
Poly(4).P(2) = 1
Poly(4).P(3) = 4
Poly(4).P(4) = 8
Poly(5).P(1) = 1
Poly(5).P(2) = 5
Poly(5).P(3) = 6
Poly(5).P(4) = 2
Poly(6).P(1) = 8
Poly(6).P(2) = 4
Poly(6).P(3) = 3
Poly(6).P(4) = 7

End Sub

Sub MainLoop()
On Error Resume Next
Do
DoEvents
   RotateCube
   ProjectCube
   ClearView
   Select Case RenderMode
   Case 5
     DrawAffineMultipleTextureCube
   Case 4
     DrawAffineSingleTextureCube
   Case 3
     DrawFlatPolyCube
   Case 2
     DrawOutlineCube
   Case 1
     DrawWireFrameCube
   Case 0
     DrawDotFrameCube
   End Select
   ViewImage.Refresh
   FPS = FPS + 1
Loop
End Sub

Sub ProjectCube()
Dim i As Byte
Dim LensDivDist As Single
For i = 1 To 8
    LensDivDist = 256 / (10000 - Corner(i).z)
    Corner(i).x2 = (Corner(i).x * LensDivDist) + HALFSIZE
    Corner(i).y2 = HALFSIZE - (Corner(i).y * LensDivDist)
Next
End Sub

Sub ClearView()
On Error Resume Next
Dim x As Byte
Dim y As Byte

For x = 0 To SIZE
For y = 0 To SIZE
  ViewData(x, y) = 0
Next
Next
End Sub

Sub RotateCube()
Dim i As Byte
Dim t As Single

For i = 1 To 8
  'rotate on x-axis
  t = Corner(i).y * Cos(RotX) - Corner(i).z * Sin(RotX)
  Corner(i).z = Corner(i).z * Cos(RotX) + Corner(i).y * Sin(RotX)
  Corner(i).y = t
  'rotate on y-axis
  t = Corner(i).z * Cos(RotY) - Corner(i).x * Sin(RotY)
  Corner(i).x = Corner(i).x * Cos(RotY) + Corner(i).z * Sin(RotY)
  Corner(i).z = t
  'rotate on z-axis
  t = Corner(i).x * Cos(RotZ) - Corner(i).y * Sin(RotZ)
  Corner(i).y = Corner(i).y * Cos(RotZ) + Corner(i).x * Sin(RotZ)
  Corner(i).x = t
Next
Const FRICTION = 0.96
RotX = RotX * FRICTION
RotY = RotY * FRICTION
RotZ = RotZ * FRICTION
End Sub

Sub DrawDotFrameCube()
On Error Resume Next
Dim i As Byte
For i = 1 To 8
    ViewData(Corner(i).x2, Corner(i).y2) = TextureData(XX, YY)
    ViewData(Corner(i).x2 + 1, Corner(i).y2 + 1) = TextureData(XX, YY)
    ViewData(Corner(i).x2 + 1, Corner(i).y2 - 1) = TextureData(XX, YY)
    ViewData(Corner(i).x2 - 1, Corner(i).y2 + 1) = TextureData(XX, YY)
    ViewData(Corner(i).x2 - 1, Corner(i).y2 - 1) = TextureData(XX, YY)
Next
End Sub

Private Sub FPStimer_Timer()
FPSlabel = FPS
FPS = 0
End Sub

Sub DrawWireFrameCube()
DrawLine Corner(1).x2, Corner(1).y2, Corner(2).x2, Corner(2).y2, TextureData(XX, YY)
DrawLine Corner(2).x2, Corner(2).y2, Corner(3).x2, Corner(3).y2, TextureData(XX, YY)
DrawLine Corner(3).x2, Corner(3).y2, Corner(4).x2, Corner(4).y2, TextureData(XX, YY)
DrawLine Corner(4).x2, Corner(4).y2, Corner(1).x2, Corner(1).y2, TextureData(XX, YY)
DrawLine Corner(5).x2, Corner(5).y2, Corner(6).x2, Corner(6).y2, TextureData(XX, YY)
DrawLine Corner(6).x2, Corner(6).y2, Corner(7).x2, Corner(7).y2, TextureData(XX, YY)
DrawLine Corner(7).x2, Corner(7).y2, Corner(8).x2, Corner(8).y2, TextureData(XX, YY)
DrawLine Corner(8).x2, Corner(8).y2, Corner(5).x2, Corner(5).y2, TextureData(XX, YY)
DrawLine Corner(1).x2, Corner(1).y2, Corner(5).x2, Corner(5).y2, TextureData(XX, YY)
DrawLine Corner(2).x2, Corner(2).y2, Corner(6).x2, Corner(6).y2, TextureData(XX, YY)
DrawLine Corner(3).x2, Corner(3).y2, Corner(7).x2, Corner(7).y2, TextureData(XX, YY)
DrawLine Corner(4).x2, Corner(4).y2, Corner(8).x2, Corner(8).y2, TextureData(XX, YY)
End Sub

Sub DrawOutlineCube()
Dim i As Byte
For i = 1 To 6
    If Corner(Poly(i).P(1)).z + Corner(Poly(i).P(2)).z + Corner(Poly(i).P(3)).z + Corner(Poly(i).P(4)).z > 2500 Then
        DrawLine Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, TextureData(XX, YY)
        DrawLine Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, TextureData(XX, YY)
        DrawLine Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, TextureData(XX, YY)
        DrawLine Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, TextureData(XX, YY)
    End If
Next
End Sub

Sub DrawLine(ByVal xx1 As Byte, ByVal yy1 As Byte, ByVal xx2 As Byte, ByVal yy2 As Byte, Color As Byte)
Dim StepX, x, StepY, y As Single
Dim dx, dy As Integer
Dim x1, y1, x2, y2 As Byte
x1 = xx1
y1 = yy1
x2 = xx2
y2 = yy2
dx = x2 - x1
dy = y2 - y1
x = x1
y = y1
If Abs(dx) < Abs(dy) Then
        StepX = dx / dy
        If dy > 0 Then
            Do
               ViewData(x, y) = Color
               x = x + StepX
               y = y + 1
               If y = y2 Then Exit Do
            Loop
        Else
            Do
               ViewData(x, y) = Color
               x = x - StepX
               y = y - 1
               If y = y2 Then Exit Do
            Loop
        End If
Else
        StepY = dy / dx
        If dx > 0 Then
            Do
               ViewData(x, y) = Color
               y = y + StepY
               x = x + 1
               If x = x2 Then Exit Do
            Loop
        Else
            Do
               ViewData(x, y) = Color
               y = y - StepY
               x = x - 1
               If x = x2 Then Exit Do
            Loop
        End If
End If
End Sub

Sub DrawFlatPolyCube()
On Error Resume Next
Dim x, y, t, i As Byte

For i = 1 To 6
    If Corner(Poly(i).P(1)).z + Corner(Poly(i).P(2)).z + Corner(Poly(i).P(3)).z + Corner(Poly(i).P(4)).z > 2500 Then
        FastScanPolyEdge Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2
        FastScanPolyEdge Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2
        FastScanPolyEdge Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2
        FastScanPolyEdge Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2
        For y = 0 To SIZE
            If RBuff(y).x Then
                x = LBuff(y).x
                Do
                    ViewData(x, y) = Poly(i).Color
                    x = x + 1
                    If x > RBuff(y).x Then Exit Do
                Loop
                RBuff(y).x = 0
                LBuff(y).x = SIZE
            End If
        Next
    End If
Next
End Sub

Sub FastScanPolyEdge(ByVal xx1 As Byte, ByVal yy1 As Byte, ByVal xx2 As Byte, ByVal yy2 As Byte)
Dim StepX, x, StepY, y As Single
Dim dx, dy As Integer
Dim x1, y1, x2, y2, LastY As Byte
x1 = xx1
y1 = yy1
x2 = xx2
y2 = yy2
dx = x2 - x1
dy = y2 - y1
x = x1
y = y1
If Abs(dx) < Abs(dy) Then
        StepX = dx / dy
        If dy > 0 Then
            Do
               If x < LBuff(y).x Then LBuff(y).x = x
               If x > RBuff(y).x Then RBuff(y).x = x
               x = x + StepX
               y = y + 1
               If y = y2 Then Exit Do
            Loop
        Else
            Do
               If x < LBuff(y).x Then LBuff(y).x = x
               If x > RBuff(y).x Then RBuff(y).x = x
               x = x - StepX
               y = y - 1
               If y = y2 Then Exit Do
            Loop
        End If
Else
        StepY = dy / dx
        If dx > 0 Then
            Do
               If x < LBuff(y).x Then LBuff(y).x = x
               If x > RBuff(y).x Then RBuff(y).x = x
               y = y + StepY
               x = x + 1
               If x = x2 Then Exit Do
            Loop
        Else
            Do
               If x < LBuff(y).x Then LBuff(y).x = x
               If x > RBuff(y).x Then RBuff(y).x = x
               y = y - StepY
               x = x - 1
               If x = x2 Then Exit Do
            Loop
        End If
End If
End Sub

Sub DrawAffineSingleTextureCube()
On Error Resume Next
Dim x, y, t, i, dx, n As Byte
Dim u, StepU, su, v, StepV, sv As Single
Dim du, dv As Integer
n = 0
For i = 1 To 6
    If Corner(Poly(i).P(1)).z + Corner(Poly(i).P(2)).z + Corner(Poly(i).P(3)).z + Corner(Poly(i).P(4)).z > 2500 Then
        AffineScanPolyEdge Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, 0, 0, 200, 0
        AffineScanPolyEdge Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, 200, 0, 0, 200
        AffineScanPolyEdge Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, 0, 200, 200, 0
        AffineScanPolyEdge Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, 0, 0, 0, 200
        For y = 0 To SIZE
            If RBuff(y).x Then
                With LBuff(y)
                    x = .x
                    u = .u
                    v = .v
                End With
                With RBuff(y)
                    dx = .x - x
                    If .u > 200 Then .u = .u - 200
                    If .v > 200 Then .v = .v - 200
                    du = .u - u
                    dv = .v - v
                End With
                su = du / dx
                sv = dv / dx
                Do
                    ViewData(x, y) = TextureData(u, v)
                    x = x + 1
                    u = u + su
                    v = v + sv
                    If x > RBuff(y).x Then Exit Do
                Loop
                RBuff(y).x = 0
                LBuff(y).x = SIZE
            End If
        Next
    End If
Next
End Sub

Sub DrawAffineMultipleTextureCube()
On Error Resume Next
Dim x, y, t, i, dx, n As Byte
Dim u, StepU, su, v, StepV, sv As Single
Dim du, dv As Integer
n = 0
For i = 1 To 6
LoadPicArray2D TextureImage, App.Path & "\car" & i & ".bmp", TextureSA, TextureBMP, TextureData()
    If Corner(Poly(i).P(1)).z + Corner(Poly(i).P(2)).z + Corner(Poly(i).P(3)).z + Corner(Poly(i).P(4)).z > 2500 Then
        AffineScanPolyEdge Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, 0, 0, 200, 0
        AffineScanPolyEdge Corner(Poly(i).P(2)).x2, Corner(Poly(i).P(2)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, 200, 0, 0, 200
        AffineScanPolyEdge Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, Corner(Poly(i).P(3)).x2, Corner(Poly(i).P(3)).y2, 0, 200, 200, 0
        AffineScanPolyEdge Corner(Poly(i).P(1)).x2, Corner(Poly(i).P(1)).y2, Corner(Poly(i).P(4)).x2, Corner(Poly(i).P(4)).y2, 0, 0, 0, 200
        For y = 0 To SIZE
            If RBuff(y).x Then
                With LBuff(y)
                    x = .x
                    u = .u
                    v = .v
                End With
                With RBuff(y)
                    dx = .x - x
                    If .u > 200 Then .u = .u - 200
                    If .v > 200 Then .v = .v - 200
                    du = .u - u
                    dv = .v - v
                End With
                su = du / dx
                sv = dv / dx
                Do
                    ViewData(x, y) = TextureData(u, v)
                    x = x + 1
                    u = u + su
                    v = v + sv
                    If x > RBuff(y).x Then Exit Do
                Loop
                RBuff(y).x = 0
                LBuff(y).x = SIZE
            End If
        Next
    End If
Next
End Sub

Sub AffineScanPolyEdge(ByVal xx1 As Byte, ByVal yy1 As Byte, ByVal xx2 As Byte, ByVal yy2 As Byte, InitU As Byte, InitV As Byte, du As Byte, dv As Byte)
On Error Resume Next
Dim StepX, x, StepY, y, StepU, u, StepV, v As Single
Dim dx, dy As Integer
Dim x1, y1, x2, y2, LastY As Byte
x1 = xx1
y1 = yy1
x2 = xx2
y2 = yy2
dx = x2 - x1
dy = y2 - y1
x = x1
y = y1
u = InitU
v = InitV
If Abs(dx) < Abs(dy) Then
        StepX = dx / dy
        StepU = Abs(du) / Abs(dy)
        StepV = Abs(dv) / Abs(dy)
        If dy > 0 Then
            Do
               With LBuff(y)
                    If x < .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               With RBuff(y)
                    If x > .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               u = u + StepU
               v = v + StepV
               x = x + StepX
               y = y + 1
               If y = y2 Then Exit Do
            Loop
        Else
            Do
               With LBuff(y)
                    If x < .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               With RBuff(y)
                    If x > .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               u = u + StepU
               v = v + StepV
               x = x - StepX
               y = y - 1
               If y = y2 Then Exit Do
            Loop
        End If
Else
        StepY = dy / dx
        StepU = Abs(du) / Abs(dx)
        StepV = Abs(dv) / Abs(dx)
        If dx > 0 Then
            Do
               With LBuff(y)
                    If x < .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               With RBuff(y)
                    If x > .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               u = u + StepU
               v = v + StepV
               y = y + StepY
               x = x + 1
               If x = x2 Then Exit Do
            Loop
        Else
            Do
               With LBuff(y)
                    If x < .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               With RBuff(y)
                    If x > .x Then
                       .x = x
                       .u = u
                       .v = v
                    End If
               End With
               u = u + StepU
               v = v + StepV
               y = y - StepY
               x = x - 1
               If x = x2 Then Exit Do
            Loop
        End If
End If
End Sub

Private Sub RM_Click(Index As Integer)
RenderMode = Index
End Sub
