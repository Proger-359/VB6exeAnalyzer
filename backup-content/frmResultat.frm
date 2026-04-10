VERSION 5.00
Begin VB.Form frmResultat 
   Caption         =   "Résultat d'analyse"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox inText 
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmResultat.frx":0000
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmResultat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
inText.Width = Me.Width - 100
inText.Height = Me.Height - 375
End Sub

Private Sub Form_Resize()
inText.Width = Me.Width - 100
inText.Height = Me.Height - 375
End Sub
