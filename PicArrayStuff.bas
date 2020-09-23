Attribute VB_Name = "PicArrayStuff"
Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public ViewData() As Byte 'View bitmap array
Public TextureData() As Byte 'Textures bitmap array

Public ViewSA As SAFEARRAY2D
Public TextureSA As SAFEARRAY2D

Public ViewBMP As BITMAP
Public TextureBMP As BITMAP

Sub LoadPicArray2D(xPicture As Image, filepath As String, sa As SAFEARRAY2D, bmp As BITMAP, data() As Byte)
' load picture into image box
If filepath <> "" Then
  xPicture.Picture = LoadPicture(filepath)
End If
' get bitmap info from image box
GetObjectAPI xPicture.Picture, Len(bmp), bmp 'dest
' exit if not 8-bit bitmap
If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
  MsgBox " 8-Bit Bitmaps Only!", vbCritical
  Exit Sub
End If
' make the local matrix point to bitmap pixels
With sa
  .cbElements = 1
  .cDims = 2
  .Bounds(0).lLbound = 0
  .Bounds(0).cElements = bmp.bmHeight
  .Bounds(1).lLbound = 0
  .Bounds(1).cElements = bmp.bmWidthBytes
  .pvData = bmp.bmBits
End With
' copy bitmap data into byte array
CopyMemory ByVal VarPtrArray(data), VarPtr(sa), 4
End Sub

Sub PicArrayKill(data() As Byte)
' MUST be called to free up memory
CopyMemory ByVal VarPtrArray(data), 0&, 4
End Sub


