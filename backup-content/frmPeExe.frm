VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPeExe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyseur de programme VB6 compilé - Proger 2003"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmPeExe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "UnPack PE"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ouvrir"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vue hexadécimale"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reconstituer vbp"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   5040
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Exécutable (*.exe;*.ocx;*.dll)|*.exe;*.ocx;*.dll|Tout les fichiers (*.*)|*.*|"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "F:\Projet1.exe"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyser"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Résultat de l'analyse"
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4095
      Begin VB.CommandButton Command13 
         Caption         =   "Jolie interface"
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "API déclarés"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Rapport"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Etude API VB"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Contrôles VB"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ressources PE"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "API importé PE"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "en-tête PE"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmPeExe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================
' FRMPEEXE
'**********
'
'Superbe interface graphique pour le module PEexe.bas


DefLng A-Z
Option Explicit

#Const DEVMODE = "DEBUG"

Private Sub Command1_Click()
If Dir$(Text1.Text, vbNormal) = "" Then
    MsgBox "Le fichier n'existe pas (plus). Veuillez ouvrir un autre fichier.", vbInformation
    Exit Sub
End If

OpenEXE Text1.Text


If exeISVB Then
    Label4.Caption = "Cet exécutable est reconnu de VB !"
    Command2.Enabled = True
Else
    Label4.Caption = "Cet exécutable ne semble pas être de VB"
    Command2.Enabled = False
End If

If exeISPACKED Then
    Command12.Enabled = True
Else
    Command12.Enabled = False
End If

Command3.Enabled = True

End Sub

Private Sub Command10_Click()
Dim fTx As frmResultat
Set fTx = New frmResultat

    PrintSummary fTx.inText
    fTx.Visible = True
    
End Sub

Private Sub Command11_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintDeclares lFrm.inList
    lFrm.Visible = True


End Sub

Private Sub Command12_Click()
'depack PE
'tente de décompresser manuellement un fichier packé (méthode du Process Dumping, imparable :op )
'puis tente d'analyser le contenu dépacké, si d'un exe VB6 il s'agit.
'Faites gaffe, si l'exe est packé, surement que sa license vous interdit de le dépacké...
Dim r, lpid, g
Dim bArray() As Byte
Dim sAr() As String, lSz() As Long, lAdr() As Long, lID() As Long
Dim TempWork As Object
Set TempWork = New MemWork '<== une classe que j'ai dev ya bien longtemps...
'en fait, si je ne vous passe pas MemWork.cls, vous pourrez pas dépacké... mmh, interessant :p

    r = MsgBox("Attention! vous avez choisi de décompressé un PE auto-compressé! Si l'auteur du programmea choisi de protéger son application via cette méthode, vous risquez des poursuites judiciaires si vous divulgez le code décompressé. Continuer ?", vbCritical + vbYesNoCancel, "UnPacking")
    If r = vbYes Then
        MsgBox "L'exe à analysé va être lançé. Fermez-le lorsqu'il aura fini de se charger.", vbInformation + vbOKOnly, "UnPacking"
        
        'merci de ne pas voler cette méthode de unpacking :) mouais, pas très convaincant. Essayons autre chose :
        'DISCLAMER : for educational purpose only / ce code est présent pour culutre générale seulement.
        'en aucun cas son auteur (Proger) ne sera tenu responsable si quiquonque abuse de cette routine.
        lpid = Shell(Text1.Text, vbHide)
        g = TempWork.GetPIDModule(lpid, sAr(), lID(), lSz(), lAdr())
        If g = -1 Then MsgBox "Echec de l'ouverture du processus": Exit Sub
        g = TempWork.DumpModulePID(lpid, bArray(), lAdr(1), lSz(1))
        If g = -1 Then MsgBox "Echec de dump du processus": Exit Sub
        g = TempWork.SaveDumpToFile(bArray(), "dumpexeanalyse.exe")
        If g = -1 Then MsgBox "Echec de sauvegarde du processsus": Exit Sub
        
        r = MsgBox("Le programme semble être dépacké. L'analyseur va tenter de récupérer les informations VB6 (souvent eparses dans un fichier dépacké). Cliquez sur 'Oui' seulement si vous êtes SURE qu'il s'agit d'un programme VB6.", vbInformation + vbYesNoCancel, "UnPacking")
        If r = vbYes Then
            Call OpenEXE_PK("dumpexeanalyse.exe")
            DoEvents
            If exeISVB Then
                Label4.Caption = "Cet exécutable dépacké est reconnu de VB !"
            Else
                Label4.Caption = "Cet exécutable dépacké ne semble pas être de VB"
            End If

            MsgBox "Analyse terminé. Le fichier dépacké a été supprimé.", vbInformation + vbOKOnly, "UnPacking"
            Close
        End If
        Kill "dumpexeanalyse.exe"
    End If

End Sub

Private Sub Command13_Click()
Dim frmJI As frmJolieInterface
Set frmJI = New frmJolieInterface

    PrintJolieInterface frmJI.Arbre
    frmJI.Visible = True

End Sub

Private Sub Command2_Click()
'reconstruit les feuilles au format .frm :)
Dim i

If exeISVB Then
    For i = 1 To UBound(exeVB_FORMS())
        Call VBCTRL_FrmRebuild(exeVB_FORMS(i), App.Path & "\rebuild\")
    Next i
End If

Call VBCTRL_VbpRebuild(exeVB_FORMS(), App.Path & "\rebuild\")

MsgBox "projet exporté vers " & App.Path & "\rebuild\"

End Sub

Private Sub Command3_Click()
frmHexa.Show
frmHexa.SetParams Text1.Text, 0
PrintExe Text1.Text, 0, 16, frmHexa.Picture1
frmHexa.SetFocus
End Sub

Private Sub Command4_Click()
Dim tRs As String
Dim eHead As Integer

    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    tRs = CommonDialog1.FileName
    If tRs = "" Or tRs = "2" Then Exit Sub
    Text1.Text = tRs
    
    Label1.Caption = Int(FileLen(tRs) / 1024) & " Ko"
    'Label2.Caption = GetAttr(tRs)
    Open Text1.Text For Binary Access Read As #10
        Get #10, 1, eHead
    Close #10
    
    If eHead = 23117 Then
        Label2.Caption = "exe validé par header."
        Command1.Enabled = True
    Else
        Label2.Caption = "exe non validé par header."
        Command1.Enabled = False
    End If
        
End Sub

Private Sub Command5_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintOutPE lFrm.inList
    lFrm.Visible = True
    
End Sub

Private Sub Command6_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintImport lFrm.inList
    lFrm.Visible = True

End Sub

Private Sub Command7_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintRessources lFrm.inList
    lFrm.Visible = True

End Sub

Private Sub Command8_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintControls lFrm.inList
    lFrm.Visible = True

End Sub

Private Sub Command9_Click()
Dim lFrm As frmList
Set lFrm = New frmList
    
    PrintVBAPI lFrm.inList
    lFrm.Visible = True

End Sub

Private Sub Form_Load()

    Me.Show
    DoEvents
    
    #If DEVMODE <> "DEBUG" Then
        Text1.Text = ""
        Label1.Caption = ""
        Label2.Caption = ""
    #End If
    
    'init du corrélateur (ça existe ce mot?)
    PEexe.VBfunc_Description_Init App.Path & "\VB60_APIDEF.txt"
    'init du désassembleur
    unASM.Init_unASM
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
