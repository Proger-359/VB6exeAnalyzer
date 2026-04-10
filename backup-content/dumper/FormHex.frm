VERSION 5.00
Begin VB.Form FormHex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualisateur hexa"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8055
   Icon            =   "FormHex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      LargeChange     =   20
      Left            =   7800
      Min             =   -32768
      TabIndex        =   1
      Top             =   0
      Value           =   -32768
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MenuGoto 
         Caption         =   "Aller &A ..."
      End
      Begin VB.Menu MenuT0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSeek 
         Caption         =   "&Rechercher..."
         Index           =   0
      End
      Begin VB.Menu MenuSeek 
         Caption         =   "Rechercher &suivant"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenuT1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSave 
         Caption         =   "Enregister sous..."
      End
      Begin VB.Menu MenuT2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuClose 
         Caption         =   "&Fermer"
      End
   End
End
Attribute VB_Name = "FormHex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================
' afficheur hexa
'****************
'basé sur la fonction GetHexViewTbl de la classe MemWork
'de Proger
'

Option Explicit

Private ByteTblToView() As Byte
Private HexW As New MemWork

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal lpDest As Any, ByVal lpSource As Any, ByVal cbCopy As Long)

Private SeekStr As String
Private SeekLastPos As String

Public Sub StartMe(ByRef TblBytes() As Byte)
Dim TlLen As Long, TmpStr As String, r As Long

TlLen = UBound(TblBytes())
ReDim ByteTblToView(1 To TlLen) As Byte
RtlMoveMemory VarPtr(ByteTblToView(1)), VarPtr(TblBytes(1)), TlLen
r = HexW.GetHexViewTbl(ByteTblToView, 1, 320, TmpStr)
If TlLen / 16 < 65535 Then
    VScroll1.Max = Int(TlLen / 16 - 20) - 32768
End If
Me.Show
DoEvents
Text1.Text = TmpStr
TmpStr = "" 'purge mémoire
End Sub

Private Sub Form_Load()
DoEvents
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Menu, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase ByteTblToView()
Set HexW = Nothing
Text1.Text = ""
Me.Hide
End Sub

Private Sub MenuClose_Click()
Unload Me
End Sub

Private Sub MenuGoto_Click()
Dim t As String
Dim NewPos As Long
t$ = InputBox("Offset (en décimal ou en hexadécimal. Mettre un 'h' à la fin de l'offset si c'est en hexa.)", "Aller à...")
If t$ = "" Or t$ = "2" Then Exit Sub
If LCase(Right(t$, 1)) = "h" Then
    NewPos = Val("&h" & Left(t$, Len(t$) - 1) & "&")
Else
    NewPos = Val(t$)
End If

NewPos = Int(NewPos / 16) * 16
If (NewPos + 320) > UBound(ByteTblToView()) Then
    NewPos = Int((UBound(ByteTblToView) - 320) / 16) * 16
ElseIf NewPos <= 0 Then
    NewPos = 0
Else

End If
VScroll1.Value = CInt(NewPos / 16 - 32768)
End Sub

Private Sub MenuSeek_Click(Index As Integer)
Dim i As Long, bs As Byte
Dim t As String

Select Case Index
Case 0 'premiere recherche
    t$ = InputBox("Entrer la chaine recherché." & vbCrLf & _
        "Mettez un 'a' au bout pour indiquer une chaine ANSI" & vbCrLf & _
        "un 'd' au bout pour indiquer une valeur décimale.", "Rechercher")
    If t$ = "" Or t$ = "2" Then Exit Sub
    Select Case Right$(t$, 1)
        Case "a"
            bs = Asc(Left$(t$, 1))
        Case "d"
            bs = Asc(Left$(t$, 1))
            'i = Val(Left$(t$, Len(t$) - 1))
        Case Else
            MsgBox "erreur"
            Exit Sub
    End Select
    For i = 1 To UBound(ByteTblToView)
        If bs = ByteTblToView(i) Then
            'g pas fini de codéééééééé!
            
            Exit For
        End If
    Next i
Case 1 'suivant

End Select
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Menu, 1
End Sub

Private Sub VScroll1_Change()
Dim TmpStr As String, t As Long, nv As Long, r As Long

If Text1.Text = "" Then Exit Sub
t = CLng(VScroll1.Value) + 32768
nv = t * 16
r = HexW.GetHexViewTbl(ByteTblToView, 1 + nv, 320 + nv, TmpStr)
Text1.Text = TmpStr
TmpStr = ""
End Sub
