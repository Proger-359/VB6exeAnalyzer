VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Hex Dump"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8640
   LinkTopic       =   "Form2"
   ScaleHeight     =   7635
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dump!"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   6840
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DUMP"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Dumper.frx":0000
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Taille"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pointeur"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Deb As Long
Dim LenG As Long
Dim BufRec() As Byte
Dim Ostr As String
Dim LBf As String
LBf = ""
Deb = Val(Text2.Text)
LenG = Val(Text3.Text)
If Deb = 0 Then Exit Sub
If LenG < 1 Then Exit Sub
ReDim BufRec(1 To LenG) As Byte
RtlMoveMemory BufRec(1), ByVal Deb, LenG
Ostr = "000001"
For i = 1 To LenG
Ostr = Ostr & " " & Right("00" & Hex(BufRec(i)), 2)
LBf = Chr$(CLng(BufRec(i)))
If i Mod 16 = 0 Then
    Ostr = Ostr & vbCrLf & Right("000000" & Val(i), 6)
    LBf = ""
End If

Next i
Text1.Text = Ostr
Text4.Text = LBf
End Sub

Private Sub Command2_Click()
Dim TeT As New MemWork
Dim StTbl() As Long
Dim MNS() As String
Dim StSz() As Long
Dim OutMode() As Byte
Command2.Caption = "Dump...": DoEvents
'Form2.Caption = TeT.GetPIDName(Val(Text5.Text))
t = TeT.GetPIDModule(Val(Text5.Text), MNS(), StTbl(), StSz())
t = TeT.DumpModulePID(Val(Text5.Text), OutMode(), StTbl(1), StSz(1))
'Ostr = "000001"
'LenG = UBound(StTbl())
'For i = 1 To LenG
'Ostr = Ostr & " " & Right("00" & Hex(StTbl(i)), 2)
'LBf = Chr$(CLng(StTbl(i)))
'If i Mod 16 = 0 Then
'    Ostr = Ostr & vbCrLf & Right("000000" & Val(i), 6)
'    LBf = ""
'End If''

'Next i
'Text1.Text = Ostr
'Text4.Text = LBf
Command2.Caption = "Dump!"
t = TeT.SaveDumpToFile(OutMode(), "text.hex")
End Sub

Private Sub Command3_Click()
Dim r As New MemWork
Dim tpd() As Long
t = r.GetActiveProcess(tpd())
t = r.KillPID(Val(Text5.Text))
End Sub

Private Sub Form_Load()
Dim h As Long
Text2.Text = VarPtr(h)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text4.Text = HexToStr(Text1.SelText)
End Sub
