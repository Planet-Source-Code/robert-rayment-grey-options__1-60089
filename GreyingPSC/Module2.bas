Attribute VB_Name = "Module2"
' Greying & HSL Code

Option Explicit

Public Sub GreyARR(W As Long, H As Long, Index As Integer)
' Picture in Public ARR(W,H)
Dim ix As Long, iy As Long
Dim Cul As Long
Dim R As Byte, G As Byte, B As Byte
' For Hue
Dim zR As Single, zG As Single, zB As Single
Dim zH As Single, zS As Single, zL As Single
   
   For iy = 1 To H
   For ix = 1 To W
      Cul = ARR(ix, iy)
      B = (Cul And &HFF&)
      G = (Cul And &HFF00&) / &H100&
      R = (Cul And &HFF0000) / &H10000
      Select Case Index
      Case 1
         G = 0.3 * R + 0.6 * G + 0.1 * B
      Case 2
         G = (1& * R + G + B) \ 3
      Case 3
         ' G = G
      Case 4, 5
         Cul = (Sqr(1& * R * R + 1& * G * G + 1& * B * B))
         If Index = 5 Then Cul = Cul \ 2
         If Cul > 255 Then Cul = 255
         G = Cul
      Case 6, 7, 8   ' H,S,L
         zR = R
         zG = G
         zB = B
         RGB2HSL zR, zG, zB, zH, zS, zL  ' In: zRGB  Out: zHSL
         Select Case Index
         Case 6: G = zH
         Case 7: G = zS
         Case 8: G = zL
         End Select
      End Select
      
      ARR(ix, iy) = RGB(G, G, G)
   
   Next ix
   Next iy
End Sub
   
'#######################################################

Public Sub RGB2HSL(ByVal zR As Single, ByVal zG As Single, ByVal zB As Single, _
   zH As Single, zS As Single, zL As Single)
' In: zRGB  Out: zHSL
Dim ColMax As Long, ColMin As Long
Dim MmM As Long, MpM As Long
Dim zMul As Single
Dim zRD As Single, zGD As Single, zBD As Single
   ColMax = zR
   If zG > zR Then ColMax = zG
   If zB > ColMax Then ColMax = zB
   ColMin = zR
   If zG < zR Then ColMin = zG
   If zB < ColMin Then ColMin = zB
   MmM = ColMax - ColMin
   MpM = ColMax + ColMin
   zL = MpM / 2
   
   If ColMax = ColMin Then
      zS = 0
      zH = 170
   Else
      If zL <= 127.5 Then
         zS = MmM * 255 / MpM
      Else
         zS = MmM * 255 / (510 - MpM)
      End If
      zMul = 255 / (MmM * 6)
      zRD = (ColMax - zR) * zMul
      zGD = (ColMax - zG) * zMul
      zBD = (ColMax - zB) * zMul
      Select Case ColMax
      Case zR: zH = zBD - zGD
      Case zG: zH = 85 + zRD - zBD
      Case zB: zH = 170 + zGD - zRD
      End Select
      If zH < 0 Then zH = zH + 255
   End If
End Sub

Public Sub HSL2RGB(ByVal zH As Single, ByVal zS As Single, ByVal zL As Single, _
   zR As Single, zG As Single, zB As Single)
' In: zHSL   Out: zRGB
Dim zFactA As Single, zFactB As Single

   If zH > 255 Then zH = 255
   If zS > 255 Then zS = 255
   If zL > 255 Then zS = 255
   If zH < 0 Then zH = 0
   If zS < 0 Then zS = 0
   If zL < 0 Then zS = 0
   If zS = 0 Then
      zR = zL
      zG = zR
      zB = zR
   Else
      If zL <= 127.5 Then
         zFactA = zL * (255 + zS) / 255
      Else
         zFactA = zL + zS - zL * zS / 255
      End If
      zFactB = zL + zL - zFactA
            
      zR = (Hue2RGB(zFactA, zFactB, zH + 85)) And 255
      zG = (Hue2RGB(zFactA, zFactB, zH)) And 255
      zB = (Hue2RGB(zFactA, zFactB, zH - 85)) And 255
   End If
End Sub

Public Function Hue2RGB(zFA As Single, zFB As Single, ByVal zH As Single) As Long
' Called by HSL2RGB
   Select Case zH
   Case Is < 0: zH = zH + 255
   Case Is > 255: zH = zH - 255
   End Select
       
   Select Case zH
   Case Is < 42.5
      Hue2RGB = zFB + 6 * (zFA - zFB) * zH / 255
   Case Is < 127.5
      Hue2RGB = zFA
   Case Is < 170
      Hue2RGB = zFB + 6 * (zFA - zFB) * (170 - zH) / 255
   Case Else
      Hue2RGB = zFB
   End Select
End Function



