Attribute VB_Name = "Module1"
'Picture DIBs & Display

Option Explicit

Private Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private Type BITMAPINFO
   bmi As BITMAPINFOHEADER
End Type

Private Const DIB_PAL_COLORS = 1 '  system colors

Private Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long
'--------------------------
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal x As Long, ByVal y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'StretchDIBits PICD.hDC, 0&, 0&, W4, H4, 0&, 0&, W, H, b8(1, 1), BS, DIB_RGB_COLORS, vbSrcCopy
'--------------------------

Public ARR() As Long


Public Sub MovePICtoARR(p As PictureBox, W As Long, H As Long)
Dim BS As BITMAPINFO
Dim PIMAGE As Long
Dim NewDC As Long
Dim OldH As Long

   ReDim ARR(1 To W, 1 To H)
   PIMAGE = p.Image
   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PIMAGE)
   With BS.bmi
      .biSize = 40
      .biwidth = W
      .biheight = H
      .biPlanes = 1
      .biBitCount = 32     ' 32-bit colors
   End With
   
   If GetDIBits(NewDC, PIMAGE, 0, H, ARR(1, 1), BS, DIB_PAL_COLORS) = 0 Then
      MsgBox "DIB Error in MovePICtoARR 32bpp", vbCritical, " "
      End
   End If
   ' Clear up
   SelectObject NewDC, OldH
   DeleteDC NewDC
End Sub

Public Sub DisplayARR(p As PictureBox, W As Long, H As Long)
Dim BS As BITMAPINFO
   With BS.bmi
      .biSize = 40
      .biwidth = W
      .biheight = H
      .biPlanes = 1
      .biBitCount = 32    ' Sets up 32-bit colors
   End With
   
   p.Picture = LoadPicture
      
   SetStretchBltMode p.hdc, HALFTONE  ' NB Of Dest picbox
   If StretchDIBits(p.hdc, 0, 0, W, H, 0, 0, _
      W, H, ARR(1, 1), BS, DIB_PAL_COLORS, vbSrcCopy) = 0 Then
         MsgBox "StretchDIBits Error", vbCritical, " "
      End
   End If
   p.Refresh
End Sub


