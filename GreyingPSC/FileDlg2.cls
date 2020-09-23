VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OSDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OSDialog  FileDlg2.cls  Open Save Dialog

' (Modified from vbAccelerator.com)

'  Use eg
'  On Form1:=

'' For using OSDialog(FileDlg2.cls)
'  Dim CommonDialog1 As OSDialog

' Examples:-
'Dim Title$, Filt$, InDir$
'Dim FIndex As Long
'  LOAD
'   Set CommonDialog1 = New OSDialog
'   Title$ = "Load a picture file"
'   Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
'   Filt$ = "Open vbp (*.vbp)|*.vbp|All files (*.*)|*.*"

'   InDir$ = CurrPath$ 'Pathspec$
'   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
'   Set CommonDialog1 = Nothing
'   FIndex = 1 bmp
'   FIndex = 2 jpg
'   etc

'  SAVE
'   Set CommonDialog1 = New OSDialog
'   Title$ = "Save Mask as 2-color bmp"
'   Filt$ = "Save bmp|*.bmp"
'   InDir$ = CurrPath$ 'Pathspec$
'   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
'   Set CommonDialog1 = Nothing
'
'   Len(FileSpec$)=0 for cancel

Option Explicit
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Private Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" _
    (ByVal lpString As String) As Long

Private Const MAX_PATH = 2048 ' To accomodate multi-select string
Private Const MAX_FILE = 2048
Private Const MULTIFILEOPENORD = 1537

Private Type OPENFILENAME
    lStructSize As Long          ' UDT length
    hwndOwner As Long            ' Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Filter
    lpstrCustomFilter As String  ' Ignored
    nMaxCustFilter As Long       ' Ignored
    nFilterIndex As Long         ' FilterIndex
    lpstrFile As String          ' FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' InitDir
    lpstrTitle As String         ' Dialog Title
    Flags As Long                ' Flags
    nFileOffset As Integer       ' Ignored
    nFileExtension As Integer    ' Ignored
    lpstrDefExt As String        ' DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored
    lpTemplateName As Long       ' Ignored
End Type

Public Enum OpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONE2RKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" _
    (file As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" _
    (file As OPENFILENAME) As Long

Dim m_lExtendedError As Long

' If parameter MultiSelect is True, dialog will be new style

Function ShowOpen(Optional FileName As String, _
                  Optional DlgTitle As String, _
                  Optional Filter As String = "All (*.*)| *.*", _
                  Optional InitDir As String, _
                  Optional DefaultExt As String = "", _
                  Optional owner As Long = -1, _
                  Optional FilterIndex As Long = 1, _
                  Optional MultiSelect As Boolean = False, _
                  Optional lpTemplateName As Long = False, _
                  Optional FileTitle As String, _
                  Optional FileMustExist As Boolean = True, _
                  Optional ReadOnly As Boolean = False, _
                  Optional HideReadOnly As Boolean = False, _
                  Optional Flags As Long = 0) As String

Dim typOpenFile As OPENFILENAME
Dim S As String
Dim CHS As String
Dim i As Integer
Dim mResult As Long

Dim p As Long

m_lExtendedError = 0

With typOpenFile
    .lStructSize = Len(typOpenFile)

     ' Add in specific flags and STRIP out non-VB flags
    .Flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
         (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
         (-ReadOnly * OFN_READONLY) Or _
         (-HideReadOnly * OFN_HIDEREADONLY) Or _
         (.Flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    If owner <> -1 Then .hwndOwner = owner
    .Flags = .Flags Or OFN_EXPLORER
    .lpstrInitialDir = InitDir
    .lpstrDefExt = DefaultExt
    .lpstrTitle = DlgTitle
    .lpTemplateName = MULTIFILEOPENORD

    ' To make Windows-style filter, replace | and : with nulls
    For i = 1 To Len(Filter)
        CHS = Mid$(Filter, i, 1)
        If CHS = "|" Or CHS = ":" Then
             S = S & vbNullChar
        Else
             S = S & CHS
        End If
    Next

    ' Put double null at end
    S = S & vbNullChar & vbNullChar
    .lpstrFilter = S
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    S = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = S
    .nMaxFile = MAX_PATH
    S = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = S
    .nMaxFileTitle = MAX_FILE

    mResult = GetOpenFileName(typOpenFile)
    
    If mResult = 1 Then
         ' Find terminating string of at least double vbNullChars ||
         mResult = InStr(1, .lpstrFile, vbNullChar & vbNullChar)
         If mResult = 0 Then
            FileName$ = .lpstrFile
         Else
            '' Original
            '' Remove excess vbNullChars
            ''FileName$ = Left$(.lpstrFile, mResult - 1)
            
            ' Find 1st vbNullChar
            p = InStr(1, .lpstrFile, vbNullChar)
            If p = 0 Then  ' No vbNullChar - ERROR
               FileName$ = vbNullString
               If mResult <> 0 Then    ' 0 is Cancel, else extended error
                    m_lExtendedError = CommDlgExtendedError()
               End If
            Else
               FileName$ = Left$(.lpstrFile, p - 1)
               FilterIndex = .nFilterIndex
            End If
         End If
    Else
         FileName$ = vbNullString
         If mResult <> 0 Then    ' 0 is Cancel, else extended error
              m_lExtendedError = CommDlgExtendedError()
         End If
    End If
End With
ShowOpen = FileName
End Function


Private Function StrZToStr(S As String) As String
    StrZToStr = Left$(S, lstrlen(S))
End Function

Function ShowSave(Optional FileName As String, _
                  Optional DlgTitle As String, _
                  Optional Filter As String = "All (*.*)| *.*", _
                  Optional InitDir As String, _
                  Optional DefaultExt As String, _
                  Optional owner As Long = -1, _
                  Optional FilterIndex As Long = 1, _
                  Optional FileTitle As String, _
                  Optional OverWritePrompt As Boolean = True, _
                  Optional Flags As Long) As String
            
Dim typOpenFile As OPENFILENAME
Dim S As String
Dim CHS As String
Dim i As Integer
Dim mResult As Long

m_lExtendedError = 0

With typOpenFile
    .lStructSize = Len(typOpenFile)

    ' Add in specific flags and STRIP out non-VB flags
    .Flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
     OFN_HIDEREADONLY Or _
     (Flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    If owner <> -1 Then .hwndOwner = owner
    .lpstrInitialDir = InitDir
    .lpstrDefExt = DefaultExt
    .lpstrTitle = DlgTitle

    ' Make new filter with bars (|) replacing nulls
    ' and double null at end
    For i = 1 To Len(Filter)
         CHS = Mid$(Filter, i, 1)
         If CHS = "|" Or CHS = ":" Then
              S = S & vbNullChar
         Else
              S = S & CHS
         End If
    Next
    ' Put double null at end
    S = S & vbNullChar & vbNullChar
    .lpstrFilter = S
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    S = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = S
    .nMaxFile = MAX_PATH
    S = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = S
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero

    mResult = GetSaveFileName(typOpenFile)
    
    If mResult = 1 Then
         FileName = StrZToStr(.lpstrFile)
         '  If you initiate the variables,
         '  you can return the value(s)
         'FileTitle = StrZToStr(.lpstrFileTitle)
           ' Return the filter index '' here 1 bmp, 2 gif
         FilterIndex = .nFilterIndex
    Else
         FileName = vbNullString
         If mResult <> 0 Then   ' 0 is Cancel, else extended error
              m_lExtendedError = CommDlgExtendedError()
         End If
    End If
End With
ShowSave = FileName
End Function

