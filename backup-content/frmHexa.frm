VERSION 5.00
Begin VB.Form frmHexa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aperçu hexadécimal"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3975
      LargeChange     =   10
      Left            =   8880
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmHexa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private AffFile As String
Private cOffset As Long


Sub SetParams(LeFichier As String, CurOffset As Long)
Dim l, s
AffFile = LeFichier
cOffset = CurOffset

l = FileLen(LeFichier)
s = l \ 16
If s > 32767 Then s = 32767

VScroll1.Min = 0
VScroll1.Max = s

End Sub


Private Sub VScroll1_Change()
PrintExe AffFile, VScroll1.Value, 24, Picture1
cOffset = CLng(VScroll1.Value) * 16
Me.Caption = "Hex : offset " & Hex(cOffset) & "  (" & cOffset & ")"
End Sub

Private Sub VScroll1_Scroll()
cOffset = CLng(VScroll1.Value) * 16
Me.Caption = "Hex : offset " & Hex(cOffset) & "  (" & cOffset & ")"
End Sub
