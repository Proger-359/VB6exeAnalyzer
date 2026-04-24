VERSION 5.00
Begin VB.Form FormOMat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CheatOmatic"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   Icon            =   "FormOMat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Définir"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rechercher"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Nouvelle valeur :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Valeur a rechercher (chiffre de 0 à 2 milliards) :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "FormOMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================
' emul'cheatomatic
'******************
'basé sur les fonctions de MemWork
'par Proger
'
'Sources :
'Aucune. Basé sur le principe de Cheat-O-Matic
'
'Attention! ce prog ne recherche que des valeurs type byte, integer ou long
'Ce prog est incapable de rechercher des strings

Option Explicit

Private CheatEngine As New MemWork

Private TableOffsets() As Long
Private UserPID As Long
Private IsNewSearch As Boolean


Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal lpDest As Any, ByVal lpSource As Any, ByVal cbCopy As Long)

Public Sub IniCheating(urpid As Long)
UserPID = urpid
IsNewSearch = True
Label4.Caption = "Indiquez une valeur, contenu dans le programme, puis cliquez sur Rechercher."
FormOMat.Show
Text1.SetFocus
End Sub

Private Sub Command1_Click()
Dim Lv As Long, r As Long
Dim GoSk As Long
Dim LvTb() As Byte
Lv = Val(Text1.Text)
Call CheatEngine.CreateTblFromString(Text1.Text, LvTb())
If IsNewSearch = True Then
    If Lv = 0 Or Lv = 255 Then
        r = MsgBox("Attention, " & Lv & " est une valeur très commune. La recherche sera lente. Continuer ?", vbQuestion + vbYesNoCancel)
        If r = vbNo Or r = vbCancel Then Exit Sub
    End If
    Label4.Caption = "Recherche en cours...": Command1.Enabled = False: DoEvents
    GoSk = CheatEngine.CheatOmatic_Cherche(UserPID, LvTb(), TableOffsets())
    Command1.Enabled = True
    If GoSk = -1 Then
        MsgBox "Une erreur s'est produit à la recherche de la valeur. Le programme a peut-être quitté.", vbExclamation
        Exit Sub
    End If

    If UBound(TableOffsets()) = 1 And TableOffsets(1) = 0 Then
        Label4.Caption = "Valeur introuvable."
    Else
        IsNewSearch = False
        Label4.Caption = "Changez la valeur dans le jeu et dans cette fenêtre et cliquez sur Rechercher."
        Label3.Caption = UBound(TableOffsets()) & " emplacement(s) trouvé(s))"
    End If
    Command2.Enabled = False: Text2.Enabled = False
    Text1.SetFocus
Else
    'une recherche a déjà été effectué
    Dim ReceivedOffset() As Long, oldlen As Long, newlen As Long
    oldlen = UBound(TableOffsets())
    Label4.Caption = "Recherche en cours...": Command1.Enabled = False: DoEvents
    GoSk = CheatEngine.CheatOmatic_ChercheIci(UserPID, LvTb(), TableOffsets(), ReceivedOffset())
    Command1.Enabled = True
    If GoSk = -1 Then
        MsgBox "Une erreur s'est produit à la recherche de la valeur. Le programme a peut-être quitté.", vbExclamation
        IsNewSearch = True
        Command2.Enabled = False: Text2.Enabled = False
        Command1.Enabled = True: Text1.SetFocus
        Exit Sub
    End If
    Command1.Enabled = True
    newlen = UBound(ReceivedOffset())
    If newlen = oldlen And ReceivedOffset(1) <> 0 Then
        Label4.Caption = "Position(s) trouvée(s)! Vous pouvez changer les valeurs du jeu!"
        Command2.Enabled = True: Text2.Enabled = True: Text2.SetFocus
        RtlMoveMemory VarPtr(TableOffsets(1)), VarPtr(ReceivedOffset(1)), newlen
        ReDim Preserve TableOffsets(1 To newlen) As Long
        IsNewSearch = True
        Label3.Caption = newlen & " emplacement(s) commun(s) trouvé(s))": Text2.SetFocus
    ElseIf newlen = 1 And ReceivedOffset(1) <> 0 Then
        Label4.Caption = "Position(s) trouvée(s)! Vous pouvez changer les valeurs du jeu!"
        Command2.Enabled = True: Text2.Enabled = True: Text2.SetFocus
        TableOffsets(1) = ReceivedOffset(1)
        ReDim Preserve TableOffsets(1 To 1) As Long
        IsNewSearch = True
        Label3.Caption = newlen & " emplacement(s) commun(s) trouvé(s))": Text2.SetFocus
    ElseIf newlen <> oldlen And ReceivedOffset(1) <> 0 Then
        Label4.Caption = "Changez encore la valeur dans le jeu et dans cette fenêtre et cliquez sur Rechercher."
        RtlMoveMemory VarPtr(TableOffsets(1)), VarPtr(ReceivedOffset(1)), newlen
        ReDim Preserve TableOffsets(1 To newlen) As Long
        Label3.Caption = newlen & " emplacement(s) commun(s) trouvé(s))"
        Text1.SetFocus
    Else
        Label4.Caption = "Impossible de retrouver les changements. Recommencez à zéro la recherche."
        IsNewSearch = True
        Label3.Caption = ""
        Text1.SetFocus
    End If
    Erase ReceivedOffset()
End If
End Sub

Private Sub Command2_Click()
Dim GoMod As Long
Dim LnV() As Byte
CheatEngine.CreateTblFromString Text2.Text, LnV()
GoMod = CheatEngine.CheatOmatic_EcritValeur(UserPID, LnV(), TableOffsets())
If GoMod = -1 Then
    MsgBox "Une erreur s'est produit lors du changement. Le programme a peut-être quitté.", vbExclamation
    IsNewSearch = True
    Command2.Enabled = False: Text2.Enabled = False
    Command1.Enabled = True: Text1.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Erase TableOffsets()
Set CheatEngine = Nothing
FormOMat.Hide
End Sub
