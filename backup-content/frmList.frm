VERSION 5.00
Begin VB.Form frmList 
   Caption         =   "frmList"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox inList 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
inList.Width = Me.Width - 100
inList.Height = Me.Height - 350
End Sub

Private Sub Form_Resize()
inList.Width = Me.Width - 100
inList.Height = Me.Height - 350
End Sub
