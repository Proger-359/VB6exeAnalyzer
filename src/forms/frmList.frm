VERSION 5.00
Begin VB.Form frmList 
   Caption         =   "frmList"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox inList 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Menu menu_root 
      Caption         =   "Root"
      Visible         =   0   'False
      Begin VB.Menu menu_seek 
         Caption         =   "Rechercher..."
      End
      Begin VB.Menu menu_copy 
         Caption         =   "Tout copier dans le presse-papier"
         Enabled         =   0   'False
      End
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

Private Sub inList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu menu_root
End If
End Sub

Private Sub menu_copy_Click()
Dim i As Long


End Sub

Private Sub menu_seek_Click()
Dim sxt As String, ctx As String
Dim i As Long

    sxt = InputBox("Taper le texte à recherché, masque autorisé." & vbCrLf & "La recherche commencera en-dessous de la ligne active.", "Rechercher dans la liste...")
    If sxt = "" Or sxt = "2" Then
        Exit Sub
    Else
        With inList
        sxt = "*" & sxt & "*"
        For i = 0 To .ListCount - 1
            ctx = .List(i)
            If ctx Like sxt Then
                .ListIndex = i 'focus
                Exit Sub
            End If
        Next i
        End With
        MsgBox "Aucunes occurences trouvées.", vbInformation + vbOKOnly, "Rechercher..."
    End If

End Sub
