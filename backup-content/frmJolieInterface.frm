VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJolieInterface 
   Caption         =   "Structure de l'exe VB6"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "frmJolieInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList IconeArbre 
      Left            =   720
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":09D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":0ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":0BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":0DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":0EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1002
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":110A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1212
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":131A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1422
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":152A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1632
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1732
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1832
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1926
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJolieInterface.frx":1B26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Arbre 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IconeArbre"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmJolieInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'la jolie interface!

DefLng A-Z

Private Sub Arbre_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i, fp As Long
Dim Pmr As frmList
Dim NodeHexStr As String

If Node.Parent Is Nothing Then Exit Sub
'si on doubleclique sur une entrée de sub, on ouvre l'éditeur hexa
If Node.Parent = "Sub Main" Then
    'bon, pour debug, juste le sub main s'il est présent.
    fp = FreeFile
    Open PEexe.exeFILENAMElong For Binary Access Read As #fp
        'temporaire...
        unASM.FileDeAsm PEexe.exeVB_CODEMAIN, fp, LOF(fp), PEexe.exeVB_CODEMAIN + &H400000, True
    Close #fp
    
    Set Pmr = New frmList
    Pmr.Caption = "Désassemblage du Sub Main()"
    Pmr.inList.FontName = "Courier New": Pmr.inList.FontSize = 8
    For i = 1 To UBound(unASM.StrDEASM())
        Pmr.inList.AddItem unASM.StrDEASM(i)
    Next i
    Pmr.Show

ElseIf Left$(Node.Parent, 17) = "Point d'entrée : " Then
    'désassemble un des subs
    
    Set Pmr = New frmList
    Pmr.Caption = "Désassemblage de " & Node.Text
    Pmr.inList.FontName = "Courier New": Pmr.inList.FontSize = 8
    i = InStrRev(Node.Text, " ") + 1
    'ici i = point d'entrée
    NodeHexStr = "&H" & Mid$(Node.Text, i, Len(Node.Text) - i)
    If Right$(NodeHexStr, 1) = "h" Then NodeHexStr = Left$(NodeHexStr, Len(NodeHexStr) - 1)
    i = CLng(NodeHexStr)

    fp = FreeFile
    Open PEexe.exeFILENAMElong For Binary Access Read As #fp
        'deasmage
        unASM.FileDeAsm i, fp, LOF(fp), i + &H400000, True
    Close #fp
    'affichage
    For i = 1 To UBound(unASM.StrDEASM())
        Pmr.inList.AddItem unASM.StrDEASM(i)
    Next i
    Pmr.Show
    
End If
End Sub

Private Sub Form_Load()
Arbre.Width = Me.Width - 100
Arbre.Height = Me.Height - 390
End Sub

Private Sub Form_Resize()
Arbre.Width = Me.Width - 100
Arbre.Height = Me.Height - 390
End Sub
