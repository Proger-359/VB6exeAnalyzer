VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPeExe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analyseur de programme VB6 compilé - Proger 2005"
   ClientHeight    =   4395
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9405
   Icon            =   "frmPeExe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "UnPack PE"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vue hexadécimale"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reconstituer vbp"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Exécutable (*.exe;*.ocx;*.dll)|*.exe;*.ocx;*.dll|Tout les fichiers (*.*)|*.*|"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Résultat de l'analyse"
      Height          =   4095
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command13 
         Caption         =   "Jolie interface"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "API déclarés"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Rapport"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Etude API VB"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Contrôles VB"
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ressources PE"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "API importé PE"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "en-tête PE"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Informations :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fichier en cours d'analyse :"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu menu_root 
      Caption         =   "&Main"
      Index           =   0
      Begin VB.Menu menu 
         Caption         =   "&Ouvrir un exécutable"
         Index           =   1
      End
      Begin VB.Menu menu 
         Caption         =   "&Analyser maintenant!"
         Index           =   2
      End
      Begin VB.Menu menu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu menu 
         Caption         =   "&Quitter"
         Index           =   4
      End
   End
   Begin VB.Menu menu_root 
      Caption         =   "&Analyse"
      Index           =   1
      Begin VB.Menu menu_an 
         Caption         =   "Informations PE"
         Index           =   1
      End
      Begin VB.Menu menu_an 
         Caption         =   "Informations VB"
         Index           =   2
      End
      Begin VB.Menu menu_an 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu menu_an 
         Caption         =   "Rapport h&iérarchisé"
         Index           =   4
      End
      Begin VB.Menu menu_an 
         Caption         =   "Listing &désassemblé"
         Index           =   5
      End
      Begin VB.Menu menu_an 
         Caption         =   "Aperçu en hé&xadécimal"
         Index           =   6
      End
   End
   Begin VB.Menu menu_root 
      Caption         =   "E&xportation des résultats"
      Index           =   2
      Begin VB.Menu menu_ex 
         Caption         =   "Rapport détaillé de l'analyse"
         Index           =   1
      End
      Begin VB.Menu menu_ex 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu menu_ex 
         Caption         =   "Informations pour décompilation"
         Index           =   3
      End
      Begin VB.Menu menu_ex 
         Caption         =   "Désassemblage du code compilé"
         Index           =   4
      End
      Begin VB.Menu menu_ex 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu menu_ex 
         Caption         =   "Binaire des feuilles/objets/modules"
         Index           =   6
      End
   End
   Begin VB.Menu menu_root 
      Caption         =   "O&utils"
      Index           =   3
      Begin VB.Menu menu_util 
         Caption         =   "Sniffeur"
         Index           =   1
      End
      Begin VB.Menu menu_util 
         Caption         =   "BackTrack"
         Index           =   2
      End
      Begin VB.Menu menu_util 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu menu_util 
         Caption         =   "Desassembleur interne"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu menu_util 
         Caption         =   "Désassembleur PERDR"
         Index           =   5
      End
      Begin VB.Menu menu_util 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu menu_util 
         Caption         =   "UnPack PE"
         Index           =   7
      End
   End
   Begin VB.Menu menu_root 
      Caption         =   "&?"
      Index           =   4
      Begin VB.Menu menu_aide 
         Caption         =   "Mode d'emploi"
         Index           =   0
      End
      Begin VB.Menu menu_aide 
         Caption         =   "A propos de..."
         Index           =   1
      End
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
'Copyright : Proger
'Création : mai 2003
'Révision : octobre 2005

DefLng A-Z
Option Explicit

#Const DEVMODE = "DEBUG"

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
                AddInfo "Cet exécutable dépacké est reconnu de VB !"
            Else
                AddInfo "Cet exécutable dépacké ne semble pas être de VB"
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

'==========================
'ADDENDUM 2005 : CE PROGRAMME N'AS PAS POUR OBJECTIF DE DECOMPILER LES FORMs
' SI VOUS SOUHAITEZ RECUPERER LES OBJETS D'UN PROGRAMME VB COMPILE
' MERCI D'UTILISER VB-REFORMER DE WARNING : SA RECUPERATION DES PROPRIETES D'OBJETS
'  EST LARGEMENT PLUS COMPLETE ET DEBUGEE.

MsgBox "Fonction désactivé. L'analyseur d'exe VB6 n'est pas un décompilateur.", vbCritical + vbOKOnly
Exit Sub
'=============================

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
'SUB MAIN

    Me.Width = 4715
    List1.Clear
    menu_root(1).Enabled = False
    menu_root(2).Enabled = False
    menu_root(3).Enabled = False
    menu(2).Enabled = False
    Me.Show
    DoEvents
    
    #If DEVMODE <> "DEBUG" Then
        AddInfo "test"
    #End If
    
    ChDir App.Path
    
    'init du corrélateur (ça existe ce mot?)
    PEexe.VBfunc_Description_Init App.Path & "\VB60_APIDEF.txt"
    'init du désassembleur
    unASM.Init_unASM
    PEexe.VBCTRL_AsmProperty
    
    'cherche si y'a un exe dans le dossier en cours pour le mettre dans text1
    Text1.Text = Dir("*.exe", vbNormal)
    If Text1.Text <> "" Then
        AddInfo "Taille du fichier : " & Int(FileLen(Text1.Text) / 1024) & " Ko"
        If IsExe(Text1.Text) Then
            AddInfo "exe validé par header."
            menu(2).Enabled = True
        Else
            AddInfo "exe non validé par header."
            'menu(2).Enabled = False
        End If
    End If
    
    'init du désassembleur via PERDR
    If Not PeReaderExt.CheckPERDR Then
        Me.menu_util(4).Checked = True
        Me.menu_util(5).Checked = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub List1_DblClick()
MsgBox List1.List(List1.ListIndex), vbOKOnly + vbInformation
End Sub

Private Sub menu_aide_Click(Index As Integer)
'menu "?"
'ajouté en novembre 2008

    Select Case Index
    Case 0
        'mode d'emploi
        MsgBox "1) Menu Main > Ouvrir un executable" & vbCrLf & _
               "2) Menu Main > Analyser maintenant !" & vbCrLf & _
               "3) Menu Analyse > Rapport hiérarchisé" & vbCrLf & _
               "4) Menu Analyse > Listing désassemblé" & vbCrLf & _
               "Un jour je ferai un vrai howto... un jour!", vbInformation + vbOKOnly, "Utilisation basique"

    Case 1
        'a propos de (en anglais pour que tous le monde puisse comprendre!)
        MsgBox "VB6analyse : a VB6 native code compilation partial decompiler. " & vbCrLf & "Base code (c) Proger 2003-2005, proger@cbsky.net" & vbCrLf & _
               "Modifications, if any : -your credit here-", vbInformation + vbOKOnly, "VB6analyse about..."
        
    End Select
    
End Sub

Private Sub menu_an_Click(Index As Integer)
'menu "Analyse"
Dim i As Long
Dim fTx As frmResultat
Dim lFrm As frmList
Dim frmJI As frmJolieInterface

    Select Case Index
    Case 1
        Set lFrm = New frmList
        PrintOutPE lFrm.inList
        lFrm.Caption = "Informations PE"
        lFrm.Visible = True
        
    Case 2
        Set fTx = New frmResultat
        PrintSummary fTx.inText
        fTx.Visible = True
    
    Case 4
        Set frmJI = New frmJolieInterface
        PrintJolieInterface frmJI.Arbre
        frmJI.Visible = True
        
    Case 5
        i = MsgBox("Voulez-vous désassembler le code ? L'opération peut être longue.", vbInformation + vbYesNoCancel, "VB6 Analyse")
        If i = vbYes Then
            frmUnAsm.Show
        End If
    
    Case 6
        If PEexe.exeFILENAMElong = "" Then PEexe.exeFILENAMElong = Me.Text1.Text
        frmHexa.Show
        frmHexa.SetParams PEexe.exeFILENAMElong, 0
        PrintExe PEexe.exeFILENAMElong, 0, 20, frmHexa.Picture1
        frmHexa.SetFocus

    End Select

End Sub

Private Sub menu_Click(Index As Integer)
Dim tRs As String
Dim eHead As Integer


    Select Case Index
    Case 1
    
        CommonDialog1.FileName = ""
        CommonDialog1.ShowOpen
        tRs = CommonDialog1.FileName
        If tRs = "" Or tRs = "2" Then Exit Sub
        Text1.Text = tRs
        
        
        AddInfo "Taille du fichier : " & Int(FileLen(tRs) / 1024) & " Ko"
        If IsExe(tRs) Then
            AddInfo "exe validé par header."
            menu(2).Enabled = True
        Else
            AddInfo "exe non validé par header."
            menu(2).Enabled = False
            menu_root(1).Enabled = True
        End If
    
    Case 2

        If Dir$(Text1.Text, vbNormal) = "" Then
            MsgBox "Le fichier n'existe pas (plus). Veuillez ouvrir un autre fichier.", vbInformation
            Exit Sub
        End If
        
        OpenEXE Text1.Text
        AddInfo "Analyse terminée."
        
        
        If exeISVB Then
            AddInfo "Cet exécutable est reconnu de VB !"
            menu_root(1).Enabled = True
            menu_root(2).Enabled = True
            menu_root(3).Enabled = True
        Else
            AddInfo "Cet exécutable ne semble pas être de VB"
            menu_root(1).Enabled = False
            menu_root(2).Enabled = False
            menu_root(3).Enabled = True
        End If
        
        If exeISPACKED Then
            Command12.Enabled = True
        Else
            Command12.Enabled = False
        End If

    Case 4
        Call Form_QueryUnload(0, 0)
        
    End Select
    
End Sub

Private Sub menu_ex_Click(Index As Integer)
'EXPORTATIONS DES RESULTATS DE L'ANALYSE
'l'objectif est de permettre à des programmes tiers d'approfondir ce qui a été lu par ce programme
Dim i As Long, j As Long, k As Long, n As Long
Dim fp As Integer, fp2 As Integer, buf As String
Dim bAry() As Byte
fp = FreeFile

    Select Case Index
    Case 1 'Rapport
    ' ==> donne les informations générique sur le fichier exe VB
        MsgBox "n/a"

    Case 3 'Informations pour décompilation / nécessite DEASM
    ' ==> sort tous les rvas utilisé dans le fichier exe VB pour qu'un programme tiers désassemble
        AddInfo "Création du listing désassemblé..."
        DoEvents
        Open PEexe.exeFILENAMElong For Binary Access Read As #fp
            unASM.VBCODE_DeAsm PEexe.exeVB_CODEENTRY + 1, fp, PEexe.exeVB_CODELEN, PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase
        Close #fp
        
        buf = Mid$(PEexe.exeFILENAMElong, 1, Len(PEexe.exeFILENAMElong) - 4) & "_rvajmpvb.txt"
        AddInfo "Exportation rvajmp vers : " & buf
        DoEvents

        Open buf For Output As #fp
            Print #fp, "VB6analyse de : ", PEexe.exeFILENAMElong
            Print #fp, "Offset code compilé : ", PEexe.exeVB_CODEENTRY
            Print #fp, "Longueur code compilé : ", PEexe.exeVB_CODELEN
            n = UBound(unASM.RVAT_LIST)
            Print #fp, "RVA préidentifiés : ", n
            Print #fp, "RVA", "Struct", "Data"
            For i = 1 To n
                With unASM.RVAT_LIST(i)
                Print #fp, Hex$(.rva), .StrS, .StrD
                End With
            Next i
            Print #fp, "00000000", "EOL", "EOL"
            
            n = UBound(unASM.ASM_SubIdx)
            Print #fp, "Subs post-identifiés : ", n
            Print #fp, "RVA", "Struct", "Data"
            For i = 1 To n
                With unASM.ASM_LIST(unASM.ASM_SubIdx(i))
                Print #fp, Hex$(.rvaCode), .sStruct, .sData
                End With
            Next i
            Print #fp, "00000000", "EOL", "EOL"
            
            n = UBound(unASM.ASM_StrIdx)
            Print #fp, "Strings post-identifiés : ", n
            Print #fp, "RVA", "String"
            For i = 1 To n
                With unASM.ASM_LIST(unASM.ASM_StrIdx(i))
                Print #fp, Hex$(.rvaJump), .sData
                End With
            Next i
            Print #fp, "00000000", "EOL"
            
        Close #fp
        AddInfo "Exportation rvajmp terminée."
    Case 4 'Listing déasm / nécessite DEASM
    ' ==> copie le résultat du désassemblage dans un fichier texte
        AddInfo "Création du listing désassemblé..."
        DoEvents
        Open PEexe.exeFILENAMElong For Binary Access Read As #fp
            unASM.VBCODE_DeAsm PEexe.exeVB_CODEENTRY + 1, fp, PEexe.exeVB_CODELEN, PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase
        Close #fp
        
        buf = Mid$(PEexe.exeFILENAMElong, 1, Len(PEexe.exeFILENAMElong) - 4) & "_unasmvb.txt"
        AddInfo "Exportation unasm vers : " & buf
        DoEvents
        Open buf For Output As #fp
            Print #fp, "VB6analyse de : ", PEexe.exeFILENAMElong
            Print #fp, "Offset code compilé : ", PEexe.exeVB_CODEENTRY
            Print #fp, "Longueur code compilé : ", PEexe.exeVB_CODELEN
            n = UBound(unASM.ASM_LIST)
            Print #fp, "Total ligne du listing : ", n
            Print #fp, "RVA", "Hex Dump", "UnAsm", "Struct", "Data"
            For i = 1 To n
                With unASM.ASM_LIST(i)
                Print #fp, "00" & Hex$(.rvaCode), .sHexDump, .sUnAsm, .sStruct, .sData
                End With
            Next i
            Print #fp, "00000000", "00", "EOF", "EOF", "EOF"
        Close #fp
        AddInfo "Exportation unasm terminée."
        
    Case 6 'Code objets
    ' ==> sort des hex dump des zones du fichier exe VB définissant les modules/objets/form
        buf = Mid$(PEexe.exeFILENAMElong, 1, Len(PEexe.exeFILENAMElong) - 4) & "_structvb.txt"
        AddInfo "Exportation struct vers : " & buf
        DoEvents
        Open buf For Output As #fp
            Print #fp, "VB6analyse de : ", PEexe.exeFILENAMElong
            Print #fp, "Offset code compilé : ", PEexe.exeVB_CODEENTRY
            Print #fp, "Longueur code compilé : ", PEexe.exeVB_CODELEN
            Print #fp, "[general]"
            Print #fp, "Nom du projet : ", PEexe.exeVB_PROJECTNAME
            
            n = UBound(PEexe.exeVB_MODULES)
            Print #fp, "Nombre de forms : ", n
            Print #fp, "[pages]"
            Print #fp, 98305, "Module"
            Print #fp, 98435, "Form"
            Print #fp, 1146883, "Class Module"
            Print #fp, 1941507, "User Control"
            Print #fp, "Index", "Type", "Nom", "Subs", "VA", "Len"
            For i = 1 To n
                With PEexe.exeVB_MODULES(i)
                    Print #fp, i, .lType, .sName, .numsub, .RvaOffset, .FullLen
                End With
            Next i
            Print #fp, "EOL", 0, "EOL", 0, 0, 0
            
            fp2 = FreeFile
            Open PEexe.exeFILENAMElong For Binary Access Read As #fp2
            n = UBound(PEexe.exeVB_CONTROL)
            Print #fp, "[vb objets]"
            Print #fp, "Nombre d'objets : ", n
            Print #fp, "PageFrom", "Name", "Type", "VA", "Len", "orderID"
            n = -1
            For i = 1 To UBound(PEexe.exeVB_FORMS)
            For j = PEexe.exeVB_FORMS(i).DefPtr To PEexe.exeVB_FORMS(i).DefPtr + PEexe.exeVB_FORMS(i).DefLen - 1
                With PEexe.exeVB_CONTROL(j)
                    Print #fp, PEexe.exeVB_FORMS(i).sName, .sName, .sType, .offset, .LenTr, Hex$(.frmID)
                    ReDim bAry(1 To .LenTr)
                    Get #fp2, .offset, bAry()
                    Print #fp, PEexe.ByteToStr(bAry())
                    n = n + 1
                End With
            Next j
            Next i
            Print #fp, "EOL", "EOL", "EOL", 0, 0
            Close #fp2
        
        Close #fp
        
        AddInfo "Exportation struct terminée."

    End Select

End Sub
Sub AddInfo(StrS As String, Optional ReplaceLast As Boolean = False)
If ReplaceLast Then
    'ajout en nov.2008 pour éviter le flood.
    List1.List(List1.ListCount - 1) = StrS
    List1.ListIndex = List1.ListCount - 1
Else
    List1.AddItem StrS
    List1.ListIndex = List1.ListCount - 1
End If
DoEvents
End Sub

Private Sub menu_util_Click(Index As Integer)
Dim frmH As frmJolieInterface
Dim ib As String
Dim va As Long

Select Case Index
    Case 1 'sniffeur
        ib = InputBox(PEexe.exeFILENAMElong & vbCrLf & "Offset hexa de début du snif :", "Sniffeur")
        If ib = "" Or ib = "2" Then Exit Sub
        va = CLng("&h" & ib)
        If va < 1024 Or va > PEexe.exeFILENAMEsize Then Exit Sub
        Set frmH = New frmJolieInterface
        Call PEexe.Util_SnifStart(frmH.Arbre, va)
        frmH.Caption = "Sniffeur @ " & ib
        frmH.Show
    Case 2 'backtrack (sniffeur inversé pour retrouvé la zone appelant)
        ib = InputBox(PEexe.exeFILENAMElong & vbCrLf & "Offset hexa de début de backtrack :", "BackTrack")
        If ib = "" Or ib = "2" Then Exit Sub
        va = CLng("&h" & ib)
        If va < 1024 Or va > PEexe.exeFILENAMEsize Then Exit Sub
        Set frmH = New frmJolieInterface
        Call PEexe.Util_BackStart(frmH.Arbre, va)
        frmH.Caption = "BackTrack @ " & ib
        frmH.Show
    Case 4 'désassembleur interne
        Me.menu_util(4).Checked = True
        Me.menu_util(5).Checked = False
    Case 5 'désassembleur PERDR
        If PeReaderExt.CheckPERDR Then
            Me.menu_util(4).Checked = False
            Me.menu_util(5).Checked = True
        Else
            Me.menu_util(4).Checked = True
            Me.menu_util(5).Checked = False
            AddInfo "PERDR introuvable."
        End If
    Case 7
        Command12_Click
        menu_root(1).Enabled = True
        menu_root(2).Enabled = True
        
End Select
End Sub

