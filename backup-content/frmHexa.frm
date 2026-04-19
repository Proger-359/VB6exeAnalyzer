VERSION 5.00
Begin VB.Form frmHexa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aperçu hexadécimal"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Text            =   "0"
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "00"
      Top             =   0
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      LargeChange     =   10
      Left            =   7800
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Goto offset"
      Height          =   300
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Offset hexa :"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmHexa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private AffFile As String
Private cOffset As Long
Private InhibR As Boolean

Sub SetParams(LeFichier As String, CurOffset As Long)
Dim l, s
AffFile = LeFichier
cOffset = CurOffset

l = FileLen(LeFichier)
s = l \ 16
If s > 32767 Then s = 32767 'limitation hscrollbar !

VScroll1.min = 0
VScroll1.max = s
InhibR = True
    VScroll1.Value = cOffset \ 16
InhibR = False

End Sub


Private Sub Command1_Click()
Dim tval As Double
tval = Val("&H" & Text1.Text)
If tval > 0 And tval < (CLng(VScroll1.max) * 16) Then
    cOffset = CLng((tval \ 16) * 16)
    PrintExe AffFile, cOffset / 16, 24, Picture1
    Me.Caption = "Hex : offset " & Hex(cOffset) & "  (" & cOffset & ")"
    VScroll1.Value = cOffset / 16
End If
End Sub

Private Sub Text2_Change()
'hex to int
    Label2.Caption = Val("&h" & Text2.Text)
End Sub

Private Sub Text2_Click()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
'max text len = 8
If Len(Text2.Text) > 8 Then Text2.Text = Right$(Text2.Text, 8)
End Sub

Private Sub VScroll1_Change()
If Not InhibR Then
    PrintExe AffFile, VScroll1.Value, 20, Picture1
    cOffset = CLng(VScroll1.Value) * 16
    Me.Caption = "Hex : offset " & Hex(cOffset) & "  (" & cOffset & ")"
Else
    Exit Sub
End If
End Sub

Private Sub VScroll1_Scroll()
PrintExe AffFile, VScroll1.Value, 20, Picture1
cOffset = CLng(VScroll1.Value) * 16
Me.Caption = "Hex : offset " & Hex(cOffset) & "  (" & cOffset & ")"
End Sub
