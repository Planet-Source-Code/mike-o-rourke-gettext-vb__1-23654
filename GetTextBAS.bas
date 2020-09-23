Attribute VB_Name = "GetTextBAS"
Option Explicit
Public Pass As Integer
Public Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function DialogFile(hWnd As Long, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String, szDestDir As String) As String
Dim X As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String
' I Modified this, see name below
OFN.lStructSize = Len(OFN)
OFN.hWnd = hWnd
OFN.lpstrTitle = szDialogTitle
OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
OFN.nMaxFile = 255
OFN.lpstrFileTitle = String$(255, 0)
OFN.nMaxFileTitle = 255
OFN.lpstrFilter = szFilter
OFN.nFilterIndex = 1
OFN.lpstrInitialDir = szDefDir
OFN.lpstrDefExt = szDefExt
OFN.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST

X = GetOpenFileName(OFN)
    
If X <> 0 Then
   If InStr(OFN.lpstrFile, Chr(0)) > 0 Then
     szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr(0)) - 1)
   End If
   DialogFile = szFile
   szDestDir = Left(szFile, OFN.nFileOffset)
Else
   DialogFile = ""
End If

'Added
If InStr(OFN.lpstrFileTitle, Chr(0)) > 0 Then ' if you want just the name. I'd rename it though. I'm using the extention here (because I'm not using it otherwise).
   szDefExt = Left$(OFN.lpstrFileTitle, InStr(OFN.lpstrFileTitle, Chr(0)) - 1)
End If

End Function
