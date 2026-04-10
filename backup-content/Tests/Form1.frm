VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "T4C Ultra trainer"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Module d'accès aux commandes de Summon"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valider le password"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   2040
      ScaleHeight     =   2835
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "lespoirfaitvivr!" Then
Text2.Text = "Bravo petit cracker"
Else
Text2.Text = "Mauvais pass lamers!"
End If
End Sub

Private Sub Command2_Click()
MsgBox "Golden axe faire 7848465467"
MsgBox "SUPer truc de nunche chefeef test abcde"
End Sub
