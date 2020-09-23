VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmGetText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GetText v1.0"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "frmGetText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmds 
      Caption         =   "C"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   7020
      TabIndex        =   29
      ToolTipText     =   "Clear total count"
      Top             =   4590
      Width           =   235
   End
   Begin VB.ComboBox CboObj 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGetText.frx":030A
      Left            =   6120
      List            =   "frmGetText.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   25
      Text            =   "(Object)"
      Top             =   4140
      Width           =   2175
   End
   Begin VB.ComboBox CboRef 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGetText.frx":030E
      Left            =   6120
      List            =   "frmGetText.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   24
      Text            =   "(Reference)"
      Top             =   3690
      Width           =   2175
   End
   Begin VB.ListBox LstTypes 
      Enabled         =   0   'False
      Height          =   1230
      ItemData        =   "frmGetText.frx":0312
      Left            =   6120
      List            =   "frmGetText.frx":0314
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2340
      Width           =   2145
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "C"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   5670
      TabIndex        =   8
      ToolTipText     =   "Clear text"
      Top             =   4590
      Width           =   235
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Search"
      Top             =   4590
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3780
      TabIndex        =   6
      Top             =   4680
      Width           =   1185
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1530
      TabIndex        =   4
      ToolTipText     =   "Reload last file "
      Top             =   4590
      Width           =   235
   End
   Begin RichTextLib.RichTextBox Text1 
      CausesValidation=   0   'False
      Height          =   4065
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7170
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmGetText.frx":0316
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "&Browse"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Browse for file"
      Top             =   4590
      Width           =   1365
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "&Exit"
      Height          =   375
      Index           =   6
      Left            =   7380
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   4590
      Width           =   915
   End
   Begin VB.CommandButton Cmds 
      Caption         =   "&Get Text"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   1890
      TabIndex        =   0
      ToolTipText     =   "Find and display text"
      Top             =   4590
      Width           =   915
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      Caption         =   "Total Lines:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   8
      Left            =   6120
      TabIndex        =   28
      Top             =   4500
      Width           =   825
   End
   Begin VB.Label LblTotalCount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   6120
      TabIndex        =   27
      Top             =   4680
      Width           =   870
   End
   Begin VB.Label LblVBP 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   6210
      TabIndex        =   26
      Top             =   0
      Width           =   1920
   End
   Begin VB.Label LblCounts 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   7470
      TabIndex        =   23
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label LblCounts 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   7470
      TabIndex        =   22
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label LblCounts 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   7470
      TabIndex        =   21
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label LblCounts 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   7470
      TabIndex        =   20
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      Caption         =   "Click to Load"
      Enabled         =   0   'False
      Height          =   195
      Index           =   7
      Left            =   6120
      TabIndex        =   19
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label Lbls 
      Alignment       =   1  'Right Justify
      Caption         =   "Objects"
      Enabled         =   0   'False
      Height          =   195
      Index           =   4
      Left            =   6300
      TabIndex        =   18
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Lbls 
      Alignment       =   1  'Right Justify
      Caption         =   "References"
      Enabled         =   0   'False
      Height          =   195
      Index           =   3
      Left            =   6300
      TabIndex        =   17
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label LblCounts 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   7470
      TabIndex        =   15
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Lbls 
      Alignment       =   1  'Right Justify
      Caption         =   "Classes"
      Enabled         =   0   'False
      Height          =   195
      Index           =   2
      Left            =   6300
      TabIndex        =   14
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Lbls 
      Alignment       =   1  'Right Justify
      Caption         =   "Modules"
      Enabled         =   0   'False
      Height          =   195
      Index           =   1
      Left            =   6300
      TabIndex        =   13
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Lbls 
      Alignment       =   1  'Right Justify
      Caption         =   "Forms"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   6300
      TabIndex        =   12
      Top             =   360
      Width           =   960
   End
   Begin VB.Label LblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   4690
      Width           =   780
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Enabled         =   0   'False
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   10
      Top             =   4500
      Width           =   360
   End
   Begin VB.Label Lbls 
      AutoSize        =   -1  'True
      Caption         =   "Search:"
      Enabled         =   0   'False
      Height          =   195
      Index           =   6
      Left            =   3780
      TabIndex        =   9
      Top             =   4500
      Width           =   555
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "frmGetText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------
'  Mike O'Rourke - 5-31-01
'  Extract text from forms and bas files to
'  insert into a spell checker.
'  You can add more filters, extentions,
'  second text window???
'  I have no idea if this is already done to death or not.
'  But I needed a quick routine today.
'  Not meant to learn from, but actually use.
'  You must of course make corrections by hand.
'  Insert built-in spell checker and correction insertion,
'  and you've got it made <g>. I have a nice spell checker,
'  so I just wanted the text from my forms.
'  I'm sure this could be made a lot tighter and faster.
'  Used RichTextBox for large files.
'  Works with forms in same directory.
'  My Res is at 1152x864, so some forms are closer then they appear. <g>
'-----------------------------------------------------

Option Explicit
Dim TotalCount As Long
Dim FilePath As String
Dim Temp As String
Dim Types As Integer
Dim FName As String
Dim FoundHold As Integer
Dim WordFound As Integer

Private Sub Cmds_Click(Index As Integer)

Select Case Index
Case 0 ' Browse
  GoBrowse
Case 1 ' Reload
  Reload
Case 2 ' Display Text
  FindOnlyText
Case 3 ' Search for text
  SearchText
Case 4 ' Clear Search
  ClearText
Case 5 ' Clear total count
  TotalCount = 0
  LblTotalCount.Caption = "0"
Case 6 ' Exit
  Unload Me
End Select

End Sub

Private Sub GoBrowse()
' Get file
On Error GoTo OpenErr

Dim FExt As String
Dim FFilter As String
Dim FDir As String
Dim FTitle As String

EnableVBP False
TotalCount = 0

FFilter = "Select VB File (*.frm;*.bas;*.vbp)" & Chr$(0) & "*.frm;*.bas;*.vbp" & Chr$(0)
FDir = ""
FTitle = "Select VB File"
FExt = "frm"

Temp = DialogFile(Me.hWnd, FTitle, FDir, FFilter, FDir, FExt, FDir) ' modified
If Temp = "" Then Exit Sub

LblVBP.Caption = FExt
FilePath = FDir

Dim ih As Integer
ih = Len(Temp)
If Right$(Temp, 1) = Chr$(0) Then
 Temp = Left$(Temp, ih - 1)
End If

FName = Temp
Reload
Enable
WordFound = 0
Text1.SetFocus
Exit Sub

OpenErr:
MsgBox "Error opening file.", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "Error!"

End Sub

Private Sub FindOnlyText()
' I'm sure there is an easier way to do this.
' But this is the instant version

Dim a As String, g As Long, e As Long, i As Long, i2 As Long, h As Long, Hold As String, t As Long
On Error GoTo TextErr
FoundHold = 0
WordFound = 0
Lbls(5).Enabled = True
LblTotal.Enabled = True

Mouse1
a = Text1
Temp = ""
Text1.Text = ""

For i = 1 To Len(a)
  g = g + 1
  If Mid$(a, g, 1) = Chr$(34) Then
  e = g + 1
    For i2 = e To Len(a)
      If Mid$(a, i2, 1) = Chr$(34) Then
        h = i2 - e
        Temp = Trim$(Mid$(a, e, h))
         If Temp = "" Then Exit For ' Filters
         If InStr(Temp, ".frx") <> 0 Then Exit For
         If InStr(Temp, "MS Sans Serif") <> 0 Then Exit For
         If InStr(Temp, "MS Serif") <> 0 Then Exit For
         If InStr(Temp, "\par") <> 0 Then Exit For
         If InStr(Temp, "\\") <> 0 Then Exit For
         If Left$(Temp, 1) = "&" Then Temp = Mid$(Temp, 2, Len(Temp) - 1)
         If Len(Temp) < 2 Then Exit For ' How can you mess up on a single letter <g> - set to skill level (do you know all 4 letter words? - set to 5)
         Hold = Hold & Temp & vbNewLine
         g = i2 + 1
         t = t + 1
         Exit For
      End If
      Temp = ""
    Next
    Temp = ""
  End If
  DoEvents
Next

LblTotal.Caption = Trim$(Str$(t))
Text1.Text = Hold
Text1.SetFocus
Mouse2
Exit Sub

TextErr:
MsgBox "Error finding text.", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "Error!"
Mouse2

End Sub
Private Sub EnableVBP(Switch As Boolean)
Dim i As Integer

LstTypes.Enabled = Switch
CboRef.Enabled = Switch
CboObj.Enabled = Switch
For i = 0 To 4
  Lbls(i).Enabled = Switch
  LblCounts(i).Enabled = Switch
Next
Lbls(7).Enabled = Switch

End Sub
Private Sub Enable()
Dim i As Integer

Text2.Enabled = True
Lbls(6).Enabled = True
Lbls(8).Enabled = True
LblTotalCount.Enabled = True
For i = 1 To 5
  Cmds(i).Enabled = True
Next

End Sub
Private Sub Reload()
Dim Forms As Boolean
On Error GoTo OpenErr

Mouse1
Text1.Text = ""
Text1.LoadFile FName, rtfText
LblName.Caption = FName
Lbls(5).Enabled = False
LblTotal.Enabled = False

If LCase$(Right$(FName, 4)) = ".vbp" Then
  LstTypes.Clear
  CboRef.Clear
  CboObj.Clear
  Calculate
End If

If LCase$(Right$(FName, 4)) = ".frm" Then Forms = True Else Forms = False
If LCase$(Right$(FName, 4)) = ".vbp" Then Else CountCodeLines Forms

Text1.SetFocus
Mouse2
Exit Sub

OpenErr:
MsgBox "Error opening file.", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "Error!"
Mouse2

End Sub

Private Sub SearchText()
On Error Resume Next
Dim FoundPos As Integer
Dim FoundLine As Integer

ClearText

FoundPos = Text1.Find(Text2.Text, FoundHold, , 1)

 If FoundPos <> -1 Then
   FoundLine = Text1.GetLineFromChar(FoundPos)
   Text1.SelBold = True
   Text1.SelColor = &HF00F50
   FoundHold = FoundPos + 1
   Text1.SelStart = FoundHold
   WordFound = WordFound + 1
Else
   MsgBox "Word not found.  " & "Found " & Trim$(Str$(WordFound)), vbInformation + vbMsgBoxSetForeground + vbOKOnly, "Info"
   FoundHold = 0
   WordFound = 0
   Me.ZOrder
End If

Text1.SetFocus

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload CopyPaste
End Sub

Private Sub LstTypes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To LstTypes.ListCount - 1
  If LstTypes.Selected(i) = True Then
    FName = FilePath & LstTypes.List(i)
    Reload
    Exit For
  End If
Next
End Sub

Private Sub LstTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To LstTypes.ListCount - 1
  If LstTypes.Selected(i) = True Then
    LstTypes.ToolTipText = LstTypes.List(i) ' just so you can see the whole thing. Could add simple API for 2 bars
    Exit For
  End If
Next
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
  PopupMenu CopyPaste.mnuTools
  DoKeys
End If

End Sub

Public Sub DoKeys()
Text1.SetFocus

Select Case Pass
Case 0
  SendKeys "^c", False
Case 1
  SendKeys "^v", False
Case 2
  SendKeys "^x", False
Case 3
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
End Select

End Sub
Private Sub ClearText()

Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Text1.SelFontSize = 8
Text1.SelBold = False
Text1.SelColor = &H0
Text1.SelStart = 0
Lbls(5).Enabled = False
LblTotal.Enabled = False

End Sub
Private Sub Calculate()
On Error GoTo vbpErr
Dim Cnt(4) As Integer, i As Integer, f As Integer, a As String

f = FreeFile
Open FName For Input Shared As #f
   Do While Not EOF(f)
      Line Input #f, a
      If (Left$(a, 5) = "Form=") Then
          Cnt(0) = Cnt(0) + 1
          a = Mid$(a, 6, Len(a) - 5)
          JustName a
          LstTypes.AddItem a
      ElseIf (Left$(a, 7) = "Module=") Then
          Cnt(1) = Cnt(1) + 1
          JustName a
          LstTypes.AddItem a
      ElseIf (Left$(a, 6) = "Class=") Then
          Cnt(2) = Cnt(2) + 1
          JustName a
          LstTypes.AddItem a
      ElseIf (Left$(a, 10) = "Reference=") Then
          Cnt(3) = Cnt(3) + 1
          JustName a
          CboRef.AddItem a
      ElseIf (Left$(a, 7) = "Object=") Then
          Cnt(4) = Cnt(4) + 1
          JustName a
          CboObj.AddItem a
      End If
   Loop
   
EnableVBP True
  
For i = 0 To 4
  LblCounts(i).Caption = Cnt(i)
Next

Finish:
Close #f
Mouse2
Exit Sub

vbpErr:
MsgBox "Error opening file.", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "Error!"
Resume Finish
         
End Sub

Private Sub JustName(a As String)
Dim i As Integer

For i = 1 To Len(a)
  If Mid$(a, i, 1) = ";" Or Mid$(a, i, 1) = ":" Then
    a = Trim$(Mid$(a, i + 1, Len(a) - i))
    Exit For
  End If
Next

End Sub

Private Sub CountCodeLines(Forms As Boolean)
' I was going to run through all listed files. But I
' like this way better. Just click on individual files to get
' a total or single count. (this way you can get the total
' lines on just say .bas or .frm files)
Dim f As Integer, Count As Long, a As String, CountBegin As Boolean
On Error GoTo OpenErr

If Not Forms Then CountBegin = True

f = FreeFile
Open FName For Input Shared As #f
Do While Not EOF(f)
Line Input #f, a

'----  ' not fair counting these
If Left$(a, 20) = "Attribute VB_Exposed" Then CountBegin = True
If InStr(a, " Sub ") <> 0 Then
ElseIf InStr(a, " Function ") <> 0 Then
ElseIf Trim$(a) = "" Then
Else
  If CountBegin Then Count = Count + 1
End If

Loop

TotalCount = TotalCount + Count
LblTotalCount.Caption = TotalCount

Close #f
Exit Sub

OpenErr:
MsgBox "Error opening file.", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "Error!"
Close #f

End Sub
Private Sub Mouse1()
Screen.MousePointer = 11
End Sub
Private Sub Mouse2()
Screen.MousePointer = 0
End Sub
