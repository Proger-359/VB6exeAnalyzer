VERSION 5.00
Begin VB.Form frmUnAsmDebug 
   Caption         =   "unAsmDebug"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "00400000"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "00"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code hexa"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RVA"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frmUnAsmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Init_unASM
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ByteTbl(1 To 10) As Byte
Dim emp As Long
Dim oc2 As Long
'On Local Error Resume Next

If KeyAscii = 13 Then
    Text1.Text = Left$(Text1.Text & String$(14, "0"), 14)
    ByteTbl(1) = Val("&h" & Mid$(Text1.Text, 1, 2))
    ByteTbl(2) = Val("&h" & Mid$(Text1.Text, 3, 2))
    ByteTbl(3) = Val("&h" & Mid$(Text1.Text, 5, 2))
    ByteTbl(4) = Val("&h" & Mid$(Text1.Text, 7, 2))
    ByteTbl(5) = Val("&h" & Mid$(Text1.Text, 9, 2))
    ByteTbl(6) = Val("&h" & Mid$(Text1.Text, 11, 2))
    ByteTbl(7) = Val("&h" & Mid$(Text1.Text, 13, 2))
    oc2 = &HFFFF And (CLng(ByteTbl(1)) Or CLng(ByteTbl(2)) * 256)
    
    If oc2 > 32767 Then oc2 = -(32768 - oc2 + 32768)
    Label1.Caption = CodeToStr(ByteTbl(), _
    GetVASM(TblPtrASM(ByteTbl(1)), oc2), _
    Val("&H" & Text2.Text), emp)
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Text2.Text = Left$(Text2.Text & "00000000", 8)
    Call Text1_KeyPress(KeyAscii)
End If
End Sub
