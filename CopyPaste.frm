VERSION 5.00
Begin VB.Form CopyPaste 
   BorderStyle     =   0  'None
   ClientHeight    =   720
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuAction 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Paste"
         Index           =   1
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Select All"
         Index           =   3
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Cancel"
         Index           =   4
      End
   End
End
Attribute VB_Name = "CopyPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAction_Click(Index As Integer)
Pass = Index
Unload Me
End Sub
