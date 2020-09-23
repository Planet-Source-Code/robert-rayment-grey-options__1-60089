VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Greying"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFile 
      Caption         =   "Luminance"
      Height          =   285
      Index           =   8
      Left            =   3465
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Saturation"
      Height          =   285
      Index           =   7
      Left            =   2310
      TabIndex        =   9
      Top             =   720
      Width           =   1035
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Hue"
      Height          =   285
      Index           =   6
      Left            =   1335
      TabIndex        =   8
      Top             =   720
      Width           =   900
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "(Sqr(R^2 + G^2 + B^2)) \ 2"
      Height          =   285
      Index           =   5
      Left            =   3450
      TabIndex        =   7
      Top             =   375
      Width           =   2385
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Sqr(R^2 + G^2 + B^2)"
      Height          =   285
      Index           =   4
      Left            =   1335
      TabIndex        =   6
      Top             =   375
      Width           =   2010
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Green"
      Height          =   285
      Index           =   3
      Left            =   4980
      TabIndex        =   5
      Top             =   45
      Width           =   840
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "( R + G + B) \3"
      Height          =   285
      Index           =   2
      Left            =   3450
      TabIndex        =   4
      Top             =   45
      Width           =   1470
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Save picture"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   3
      Top             =   375
      Width           =   1140
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "0.3*R + 0.6*G + 0.1*B"
      Height          =   285
      Index           =   1
      Left            =   1335
      TabIndex        =   2
      Top             =   45
      Width           =   2010
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Load picture"
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   1155
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   75
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   0
      Top             =   1170
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Grey Options  by Robert Rayment  (April 2005)

Option Explicit

Private aPicLoaded As Boolean
Private aPicGreyed As Boolean
' PIC width & height
Private W As Long
Private H As Long
Private Pathspec$, FileSpec$, OpenPath$, SavePath$

Dim CommonDialog1 As OSDialog

Private Sub cmdFile_Click(Index As Integer)
   Select Case Index
   Case 0
      LoadPic
   Case 1 To 8
      If Len(FileSpec$) > 0 Then
         Screen.MousePointer = vbHourglass
         PIC.Picture = LoadPicture(FileSpec$)
         Refresh  ' Flash reloaded picture
         W = PIC.Width
         H = PIC.Height
         MovePICtoARR PIC, W, H
         Refresh  ' Flash reloaded picture
         GreyPic Index
      End If
      'Index
      ' 1  0.3 * R + 0.6 * G + 0.1 * B  !If not much green too dark but works well for most images!
      ' 2  (R + G + B) \ 3              !If not much green bit dark but better than 1!
      ' 3  Just Green                   !Of course no good if not much green - all black if no green!
      ' 4  Intensity  Sqr( R^2 + G^2 + B^2)          !Too bright!
      ' 5  Intensity\2  (Sqr( R^2 + G^2 + B^2)) \ 2  !Flat!
      ' 6  Hue         !Grey effect. Can be all grey!
      ' 7  Saturation  !Different grey effect. Can be all black!
      ' 8  Luminance   !Similar to Intensity\2!
   Case 9   'Save as bmp
      SaveAsPic
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub GreyPic(Index As Integer)
   If aPicLoaded Then
      GreyARR W, H, Index
      DisplayARR PIC, W, H
      aPicGreyed = True
   End If
End Sub

'#############################################################

Private Sub LoadPic()
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Filt$ = "BMP, JPG, GIF|*.bmp;*.jpg;*.gif"
   'Filt$ = "BMP(*.bmp)|*.bmp|JPEG(*.jpg)|*.jpg|GIF(*.gif)|*.gif"
   FileSpec$ = ""
   Title$ = "Load PIC"
   InDir$ = OpenPath$ 'Pathspec$
   Set CommonDialog1 = New OSDialog
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
   Set CommonDialog1 = Nothing
   If Len(FileSpec$) > 0 Then
      OpenPath$ = FileSpec$
      'FileSpec$ = Pathspec$ & "Black_Hole.jpg"
      PIC.Picture = LoadPicture(FileSpec$)
      W = PIC.Width
      H = PIC.Height
      MovePICtoARR PIC, W, H
      Refresh
      aPicLoaded = True
      aPicGreyed = False
   End If
End Sub

Private Sub SaveAsPic()
Dim SaveSpec$
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   If aPicGreyed Then
      Filt$ = "BMP|*.bmp"
      SaveSpec$ = ""
      Title$ = "Save As BMP"
      InDir$ = SavePath$ 'Pathspec$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(SaveSpec$) > 0 Then
         FixExtension SaveSpec$, ".bmp"
         SavePath$ = SaveSpec$
         'SavePicture PIC.Image, Pathspec$ & "Black_Hole.jpg"
         SavePicture PIC.Image, SaveSpec$
      End If
   End If
End Sub

Public Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".") ' NB Clears any 2nd dot + chars
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub

Private Sub Form_Load()
   Pathspec$ = App.Path
   If Right$(Pathspec$, 1) <> "\" Then Pathspec$ = Pathspec$ & "\"
   OpenPath$ = Pathspec$
   SavePath$ = Pathspec$
   FileSpec$ = ""
   
   With PIC
      .AutoRedraw = True
      .AutoSize = True
      .ScaleMode = vbPixels
   End With
   aPicLoaded = False
   aPicGreyed = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub
