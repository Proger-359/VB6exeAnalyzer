VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUnAsm 
   Caption         =   "Vue désassemblée"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6376
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "RVA"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hexa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Instruct"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Struct"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Reference"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VBcode"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Trouve texte"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Saute au rva"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lit rva dest."
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Obj ref"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Str ref"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Goto Sub :"
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmUnAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================
'frmUnAsm VBcode
'****************
'
'Pour désassemblé du code VB
'par Proger
'création octobre 2005

Option Explicit
DefLng A-Z


'table des fonctions VB par adresse relative <== DESUET
Private Type LONG_STRING
    Lngv As Long
    Strg As String 'baratin pour reference
    Stru As String 'baratin pour struct
End Type
Private Rva_Descriptor() As LONG_STRING '<== DESUET

Private RVAtoGo As String
'mise en "cache" des infos sur le fichier affiché
Private lImageBase As Long
Private sFileName As String
Private lFileLen As Long

Private Function GotoRVA(ByVal va As Long, Optional iStart As Long = 1) As Long
'recherche et surligne l'adresse à trouvé,
'renvoi l'index de la ligne de code désassemblé où s'est arrêté le scan
Dim i, j

    If va < lImageBase Then va = va + lImageBase
    For i = iStart To UBound(unASM.ASM_LIST())
        If va = unASM.ASM_LIST(i).rvaCode Then
            ListView1.SetFocus
            ListView1.ListItems.Item(i).EnsureVisible
            ListView1.ListItems.Item(i).Selected = True
            GotoRVA = i
            Exit Function
        End If
    Next i

End Function

Private Sub Combo1_Click()
'saut dans la listview à la ligne du sub
Dim v As Long
    'valide si l'adresse offset est sur 5 caractères
    v = Val("&h" & Trim$(Right$(Combo1.Text, 5)) & "&")
    Call GotoRVA(v)
    
End Sub

Private Sub Command1_Click()
Dim TfrmText As frmResultat
Set TfrmText = New frmResultat
Dim bufs As String, cmps As String, p As Long, Offs As Long, bidon As Long, v As Long
Dim fp As Integer
Dim gByte As Byte, gLong As Long, gANSI As String, gUNIC As String, gDump(1 To 14) As Byte, gAsm As String
TfrmText.inText = ""

    'cherche un adresse RVA (004xyyzz) a suivre
    bufs = ListView1.SelectedItem.SubItems(2)
    cmps = "004" '00 40 00 00 : rva des exes VB -il existe des rva étendu mébon
    p = InStr(4, bufs, cmps, vbBinaryCompare)
    If p > 0 And Len(bufs) >= (p + 7) Then
        cmps = Mid$(bufs, p, 8)
        Offs = Val("&H" & cmps)
        If Offs > lImageBase And (Offs - lImageBase) < lFileLen Then
            fp = FreeFile
            Offs = Offs - lImageBase
            Open sFileName For Binary Access Read As #fp
                'byte ?
                Get #fp, Offs + 1, gByte
                'long ?
                Get #fp, Offs + 1, gLong
                'string ANSI ?
                gANSI = PEexe.ScanString(fp, Offs + 1)
                'string UNICODE ?
                gUNIC = PEexe.ScanUnicode(fp, Offs + 1, 4096)
                'dump
                Get #fp, Offs + 1, gDump()
                'ASM decode
                p = &HFFFF And (CLng(gDump(1)) Or CLng(gDump(2)) * 256)
                If p > 32767 Then p = -(32768 - p + 32768)
                gAsm = unASM.CodeToStr(gDump(), _
                    GetVASM(TblPtrASM(gDump(1)), p), _
                    Offs + PEexe.exeOPHEAD.ImageBase, bidon, v)
            Close #fp
            
            TfrmText.Show
                
            TfrmText.inText = _
            "Suivi de l'instruction " & bufs & " :" & vbCrLf & _
            "(pointe vers " & Right$("0000" & Hex$(Offs + lImageBase), 8) & "h / " & Hex$(Offs) & "h physique)" & vbCrLf & _
            "=========" & vbCrLf & _
            "byte : " & Right$("0" & Hex$(gByte), 2) & "h (" & gByte & ")" & vbCrLf & _
            "long : " & Right$("0000000" & Hex$(gLong), 8) & "h (" & gLong & ")" & vbCrLf & _
            "ANSI : " & gANSI & vbCrLf & _
            "UNICODE : " & gUNIC & vbCrLf & _
            "dASM : " & gAsm & vbCrLf
            
            bufs = Space$(36)
            For p = 1 To 12
                Mid$(bufs, (p - 1) * 3 + 1, 2) = Right$("0" & Hex$(gDump(p)), 2) & " "
            Next p
            TfrmText.inText = TfrmText.inText & "HEX : " & bufs
            
            Exit Sub
        End If
    End If
    'inutilisation / unSet
    Unload TfrmText

End Sub

Private Sub Command2_Click()
'va au rva indiqué dans un jump ou call ou conditional jump
Dim i As Long
    For i = 1 To ListView1.ListItems.Count
        If InStr(1, ListView1.ListItems.Item(i).Text, RVAtoGo, vbBinaryCompare) > 0 Then
            ListView1.SetFocus
            ListView1.ListItems.Item(i).EnsureVisible
            ListView1.ListItems.Item(i).Selected = True
            Exit For
        End If
    Next i
End Sub

Private Sub Command3_Click()
'cherche un texte dans la listview
Dim Wtxt As String, i As Long, j As Long, l As Long, p As Long, d As Long
'reprend à partir de la selection + 1 offset
ListView1.SetFocus
d = ListView1.SelectedItem.Index + 1
If d < 1 Then d = 1
If d >= ListView1.ListItems.Count Then d = ListView1.ListItems.Count - 1
    
    Wtxt = InputBox("Texte/code/hexdump à trouver :", "VB Analyser")
    If Wtxt = "" Or Wtxt = "2" Then Exit Sub
    For i = d To ListView1.ListItems.Count
        p = InStr(1, ListView1.ListItems.Item(i), Wtxt, vbBinaryCompare)
        If p <> 0 Then
            ListView1.SetFocus
            ListView1.ListItems.Item(i).EnsureVisible
            ListView1.ListItems.Item(i).Selected = True
            Exit Sub
        End If
        j = 1: l = ListView1.ListItems.Item(i).ListSubItems.Count
        Do Until j > l
            p = InStr(1, ListView1.ListItems.Item(i).ListSubItems.Item(j).Text, Wtxt, vbBinaryCompare)
            If p <> 0 Then
                ListView1.SetFocus
                ListView1.ListItems.Item(i).EnsureVisible
                ListView1.ListItems.Item(i).Selected = True
                Exit Sub
            End If
        j = j + 1
        Loop
    Next i
    MsgBox Wtxt & " non trouvé.", vbExclamation + vbOKOnly, "VB Analyser"
End Sub

Private Sub Command4_Click()
'liste des strings trouvées
Dim i As Long
Dim tFrm As frmList
Set tFrm = New frmList
    
    tFrm.inList.Clear
    For i = 1 To UBound(ASM_StrIdx())
        tFrm.inList.AddItem unASM.ASM_LIST(unASM.ASM_StrIdx(i)).sData
    Next i
    tFrm.Caption = "Liste des strings trouvées."
    tFrm.Show

End Sub

Private Sub Command5_Click()
'liste des références aux objets trouvées
Dim i As Long
Dim tFrm As frmList
Set tFrm = New frmList
    
    tFrm.inList.Clear
    For i = 1 To UBound(ASM_ObjIdx())
        tFrm.inList.AddItem Hex$(unASM.ASM_LIST(unASM.ASM_ObjIdx(i)).rvaCode) & " : " & unASM.ASM_LIST(unASM.ASM_ObjIdx(i)).sData
    Next i
    tFrm.Caption = "Liste des ref objets trouvées."
    tFrm.Show

End Sub

Private Sub Form_Load()
'précharge, désassemble, analyse, liste le code compilé
Dim i, j, Csub, fp As Integer
lImageBase = PEexe.exeOPHEAD.ImageBase
sFileName = PEexe.exeFILENAMElong
lFileLen = PEexe.exeFILENAMEsize


    Me.Visible = True
    Me.Caption = "VBanalyse - Désassemblage..."
    Me.MousePointer = vbHourglass
    DoEvents

    'listage des subs
    Combo1.Clear
    Combo1.AddItem "Code Start : " & Hex$(PEexe.exeVB_CODEENTRY)
    Csub = UBound(PEexe.exeVB_SUBS())
    For i = 1 To Csub
        If PEexe.exeVB_SUBS(i).sName = "" Then
            Combo1.AddItem "Sub " & i & " : " & Hex$(PEexe.exeVB_SUBS(i).rvaCode)
        Else
            Combo1.AddItem "Sub " & PEexe.exeVB_SUBS(i).sName & " " & i & " : " & Hex$(PEexe.exeVB_SUBS(i).rvaCode)
        End If
    Next i
    
    DoEvents

    'désassemblage (peut être appelé depuis n'importe quel sub)
    If frmPeExe.menu_util(4).Checked Then
        fp = FreeFile
        Open PEexe.exeFILENAMElong For Binary Access Read As #fp
        unASM.VBCODE_DeAsm PEexe.exeVB_CODEENTRY + 1, fp, PEexe.exeVB_CODELEN, PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase
        Analyse.AsmListingParse fp
        Close #fp
    Else
        PeReaderExt.ExecPERDR PEexe.exeFILENAMElong
        PeReaderExt.ParsePERDR
        fp = FreeFile
        Open PEexe.exeFILENAMElong For Binary Access Read As #fp
        PeReaderExt.PeRDR_UnAsmVB PEexe.exeVB_CODEENTRY + 1, fp, PEexe.exeVB_CODELEN, PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase
        Analyse.AsmListingParse fp
        Close #fp
    End If
    
    
    Me.Caption = "VBanalyse - Affichage..."
    DoEvents
    
    'affichage
    ListView1.ListItems.Clear
    j = UBound(unASM.ASM_LIST)
    For i = 1 To j
        ListView1.ListItems.Add i, , "00" & Hex$(unASM.ASM_LIST(i).rvaCode)
        
        ListView1.ListItems.Item(i).SubItems(1) = unASM.ASM_LIST(i).sHexDump
        ListView1.ListItems.Item(i).SubItems(2) = unASM.ASM_LIST(i).sUnAsm
        ListView1.ListItems.Item(i).SubItems(3) = unASM.ASM_LIST(i).sStruct
        ListView1.ListItems.Item(i).SubItems(4) = unASM.ASM_LIST(i).sData
        If i Mod 10000 = 0 Then
            Me.Caption = "VBanalyse - Affichage... " & Int(i / j * 100) & "%"
            DoEvents
        End If
    Next i
    
    Me.Caption = "Vue désassemblée de " & PEexe.exeFILENAMElong
    Me.MousePointer = 0
    
    ListView1.SetFocus
    
End Sub

Private Sub Unused()
'vieux code de l'interpréteur
Dim i, j, k, p
Dim fva As String, hexc As String, vjmp As Long, subs As Long, Csub As Long, Casm As Long
Dim fp As Integer, Offs As Long, dump As Long


    'création du descripteur de RVA : TOUTES les entrées possibles...
    p = 1
    'For i = 1 To UBound(PEexe.exeVB_FORMS())
    '   ReDim Preserve Rva_Descriptor(p)
    '   Rva_Descriptor(p).Lngv = PEexe.exeVB_FORMS(i).rvaPtr + PEexe.exeOPHEAD.ImageBase
    '   Rva_Descriptor(p).Strg = "EP Form " & PEexe.exeVB_FORMS(i).sName
    '   p = p + 1
    'Next i
    'faudrai que je simplifie...
    For i = 1 To UBound(PEexe.exeVB6_APICALLS())
       ReDim Preserve Rva_Descriptor(p)
       Rva_Descriptor(p).Lngv = PEexe.exeVB6_APICALLS(i).rva + PEexe.exeOPHEAD.ImageBase
       If PEexe.exeVB6_APICALLS(i).ApiVbDefPtr <> 0 Then
            Rva_Descriptor(p).Strg = "API. " & PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName
            Rva_Descriptor(p).Stru = PEexe.VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, 12)), PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, fva)
        Else
            Rva_Descriptor(p).Strg = "API.???"
        End If
       p = p + 1
    Next i
    For i = 1 To UBound(PEexe.exeVB_API())
       ReDim Preserve Rva_Descriptor(p)
       Rva_Descriptor(p).Lngv = PEexe.exeVB_API(i).RvaOffset + PEexe.exeOPHEAD.ImageBase
       Rva_Descriptor(p).Strg = PEexe.exeVB_API(i).sDll & "." & PEexe.exeVB_API(i).sName
       p = p + 1
    Next i
    'pour appel via FF15h et 8BxDh (duplicata...)
    For i = 1 To UBound(PEexe.exeVB6_APICALLS())
       ReDim Preserve Rva_Descriptor(p)
       If PEexe.exeVB6_APICALLS(i).ApiVbDefPtr <> 0 Then
        Rva_Descriptor(p).Lngv = PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).VaTbl + PEexe.exeOPHEAD.ImageBase
        Rva_Descriptor(p).Strg = "API. " & PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName
        Rva_Descriptor(p).Stru = PEexe.VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, 12)), PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, fva)
        End If
       p = p + 1
    Next i
    'For i = 1 To UBound(PEexe.exeVB_CONTROL())
    '   ReDim Preserve Rva_Descriptor(p)
    '   Rva_Descriptor(p).Lngv = PEexe.exeVB_CONTROL(i).Offset + PEexe.exeOPHEAD.ImageBase
    '   Rva_Descriptor(p).Strg = "Objet " & PEexe.exeVB_CONTROL(i).sType & "." & PEexe.exeVB_CONTROL(i).sName
    '   p = p + 1
    'Next i
    For i = 1 To Csub
       ReDim Preserve Rva_Descriptor(p)
       Rva_Descriptor(p).Lngv = PEexe.exeVB_SUBS(i).rvaCode + PEexe.exeOPHEAD.ImageBase
       Rva_Descriptor(p).Strg = "Call Sub " & PEexe.exeVB_SUBS(i).sName
       p = p + 1
    Next i
    p = UBound(Rva_Descriptor())
    

    
    'ouvre le fichier pour les besoins de l'interpréteur
    fp = FreeFile
    Open PEexe.exeFILENAMElong For Binary Access Read As #fp
        
    
    'affichage du listing et interprétation partielle
    ListView1.ListItems.Clear
    Casm = UBound(unASM.StrDEASM())
    For i = 1 To Casm
        fva = Left$(unASM.StrDEASM(i), 8)
        hexc = Trim$(Mid$(unASM.StrDEASM(i), 11, 22))
        ListView1.ListItems.Add i, , fva
        ListView1.ListItems.Item(i).SubItems(1) = hexc
        ListView1.ListItems.Item(i).SubItems(2) = Trim$(Mid$(unASM.StrDEASM(i), 31))
        
        
        'analyse de la structure : point d'entrée des subs
        k = Val("&H" & fva) - PEexe.exeOPHEAD.ImageBase
        For j = 1 To Csub
            If PEexe.exeVB_SUBS(j).rvaCode = k Then
                ListView1.ListItems.Item(i).SubItems(3) = "Sub " & j
                'ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = &HA04040
                Exit For
            End If
        Next j
        
        'analyse des PUSH et CALL : ils font souvent références aux API et subs internes
        subs = Val("&h" & Left$(hexc, 2))
        If subs = &H68 Or subs = &HE8 Or subs = &HBA Or subs = &H8B Or subs = &HFF Or subs = &HC7 Then
        
            If subs = &H68 Or subs = &HE8 Or subs = &HBA Then
                '68h, E8h, BAh
                vjmp = Val("&h" & Mid$(Trim$(Mid$(unASM.StrDEASM(i), 30)), 7))
            Else
                'FF15h, 8BxDh
                'ces deux codes pointent directement vers la tables des imports
                'et non pas sur l'etrange table FF25h de VB...
                vjmp = InStr(34, unASM.StrDEASM(i), "[", vbBinaryCompare)
                If vjmp > 0 Then
                    vjmp = Val("&h" & Mid$(unASM.StrDEASM(i), vjmp + 1, 8))
                End If
            End If
            
            If vjmp > 255 Then
                For k = 1 To p
                    If Rva_Descriptor(k).Lngv = vjmp Then
                        ListView1.ListItems.Item(i).SubItems(3) = Rva_Descriptor(k).Stru
                        ListView1.ListItems.Item(i).SubItems(4) = Rva_Descriptor(k).Strg
                        'ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = &H800000
                        GoTo NCasm
                    End If
                Next k
            End If
            
            'recherche d'une string (youpi la manip de str :/ ) '##wwxxyyzz 3 5 7 9
            If subs = &H68 Or subs = &HBA Or subs = &HC7 Then
                k = 0
                If subs = &HC7 Then k = 4
                subs = Val("&h" & Mid$(hexc, 9 + k, 2) & Mid$(hexc, 7 + k, 2) & Mid$(hexc, 5 + k, 2) & Mid$(hexc, 3 + k, 2))
                If subs > 0 And subs < (PEexe.exeOPHEAD.ImageBase + PEexe.exeFILENAMEsize) Then
                    Offs = subs - PEexe.exeOPHEAD.ImageBase
                    If Offs > 4128 And Offs < PEexe.exeVB_CODEENTRY Then
                        Get #fp, Offs - 3, dump
                        If dump > 0 And dump < 32000 Then
                            fva = PEexe.ScanUnicode(fp, Offs + 1, dump)
                            If fva <> "" And Asc(fva & " ") > 15 Then
                                ListView1.ListItems.Item(i).SubItems(3) = "String :"
                                ListView1.ListItems.Item(i).SubItems(4) = Chr$(34) & fva & Chr$(34)
                                'ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = &H9000&
                                GoTo NCasm
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'analyse d'un MOV particulier : le C7h, qui est utilisé pour charger un rva de sub!
        If subs = &HC7 Then
            'ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = &HA01010
            vjmp = Val("&h" & Mid$(unASM.StrDEASM(i), 57, 8))
            If vjmp > 0 Then
                vjmp = vjmp - PEexe.exeOPHEAD.ImageBase
                For k = 1 To Csub
                    If PEexe.exeVB_SUBS(k).SubFrom = vjmp Then
                        ListView1.ListItems.Item(i).SubItems(3) = "Charge Sub"
                        ListView1.ListItems.Item(i).SubItems(4) = ListView1.ListItems.Item(i).SubItems(4) & k & ","
                    End If
                Next k
            End If
        End If
NCasm:
    Next i
    Close #fp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Caption = "Purge mémoire..."
    DoEvents
    ListView1.ListItems.Clear
    DoEvents
End Sub

Private Sub Form_Resize()
ListView1.Width = Me.Width - 120
If Me.Height > 1000 Then ListView1.Height = Me.Height - 780
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim bufs As String, cmps As String, Offs As Long, p As Long
'détermine s'il y a quelque chose à suivre (jump, call, push, mov... vers une adresse rva)
    Command2.Enabled = False
On Local Error GoTo Wrongness
    bufs = ListView1.SelectedItem.SubItems(2)
    cmps = "004" '00 40 00 00 : rva des exes VB -il existe des rva étendu mébon
    p = InStr(4, bufs, cmps, vbBinaryCompare)
    If p > 0 And Len(bufs) >= (p + 7) Then
        cmps = Mid$(bufs, p, 8)
        Offs = Val("&H" & cmps)
        If Offs > PEexe.exeOPHEAD.ImageBase And (Offs - PEexe.exeOPHEAD.ImageBase) < PEexe.exeFILENAMEsize Then
            Command1.Enabled = True
            If Offs > PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase And Offs < PEexe.exeVB_CODELEN + PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase Then
                RVAtoGo = cmps
                Command2.Enabled = True
            End If
            Exit Sub
        End If
    End If
    
    Command1.Enabled = False
    Exit Sub
Wrongness:
    Err.Clear
    Command1.Enabled = False
    
    
End Sub
