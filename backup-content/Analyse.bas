Attribute VB_Name = "Analyse"
'=================
' ANALYSE
'*********
'
'Par Proger
'juin 2006
'fonctions d'analyse du code VB6 désassemblé
'les autres modules devenant trop gros pour être pratique, j'ai dérivé des fonctions ici

Option Explicit
DefLng A-Z

Public CC_ctrl As CONTROL_DEF
Public CC_ctrlID As Long
Public CC_FrmID As Long

Sub AsmListingParse(ByVal fp As Integer)
Dim i, c, o, rva, AdV   'variables d'itération et de progression dans le code
Dim ta, ip, sp, ap, op  'variables de progression d'index
Dim p, jrva, cl, cs, cp 'variables de l'interpréteur/analyseur
Dim sub_en_cours        'index du sub qui est en cours de désassemblage
Dim last_obj, freeobj   'indique le dernier objet appelé, index indicateur de l'api freeobj
Dim bDump As Byte, iDump As Integer, lDump As Long
Dim bArray(1 To 14) As Byte
Dim sBuffer As String, sB2 As String

Erase ASM_SubIdx(): Erase ASM_StrIdx(): Erase ASM_ApiIdx(): Erase ASM_ObjIdx()
ReDim ASM_StrIdx(0) 'anticrash
cs = UBound(PEexe.exeVB_SUBS)
cp = UBound(PEexe.exeVB_EP())

    'etape 1 : préparation des descripteurs de saut hors code
    o = 1
    'liste des API et zones d'appel
    ReDim RVAT_LIST(1 To (UBound(PEexe.exeVB6_APICALLS()) * 2 + UBound(PEexe.exeVB_API()) + cs + cp))
    For i = 1 To UBound(PEexe.exeVB6_APICALLS())
       RVAT_LIST(o).rva = PEexe.exeVB6_APICALLS(i).rva + PEexe.exeOPHEAD.ImageBase
       If PEexe.exeVB6_APICALLS(i).ApiVbDefPtr <> 0 Then
            RVAT_LIST(o).StrS = "API VB"
            sBuffer = ""
            sB2 = PEexe.VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, 12)), PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName, sBuffer)
            If sBuffer <> "" Then sBuffer = "  - " & sBuffer
            RVAT_LIST(o).StrD = sB2 & sBuffer & " ( " & PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName & ")"
            If freeobj = 0 Then
                If PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).ApiName = "__vbaFreeObj" Then
                    freeobj = o
                End If
            End If
            RVAT_LIST(o + 1).rva = PEexe.exeIMPORT_APINAME(PEexe.exeVB6_APICALLS(i).ApiVbDefPtr).VaTbl + PEexe.exeOPHEAD.ImageBase
            RVAT_LIST(o + 1).StrS = RVAT_LIST(o).StrS
            RVAT_LIST(o + 1).StrD = RVAT_LIST(o).StrD

            o = o + 2
        Else
            RVAT_LIST(o).StrD = "API.?"
            o = o + 1
        End If
    Next i
    'api déclaré
    For i = 1 To UBound(PEexe.exeVB_API())
       RVAT_LIST(o).rva = PEexe.exeVB_API(i).RvaCall + PEexe.exeOPHEAD.ImageBase
       RVAT_LIST(o).StrS = "API user"
       RVAT_LIST(o).StrD = PEexe.exeVB_API(i).sName & " Lib " & PEexe.exeVB_API(i).sDll
       o = o + 1
    Next i
    'liste des subs objets connus...
    For i = 1 To cs
       RVAT_LIST(o).rva = PEexe.exeVB_SUBS(i).rvaCode + PEexe.exeOPHEAD.ImageBase
       RVAT_LIST(o).StrS = "Call Sub"
       RVAT_LIST(o).StrD = "Sub " & i '& " (" & PEexe.exeVB_SUBS(i).sName & ")"
       o = o + 1
    Next i
    'liste des EP (expérimental)
    For i = 1 To cp
       RVAT_LIST(o).rva = PEexe.exeVB_EP(i).rvaJmp
       RVAT_LIST(o).StrS = "EP " & PEexe.exeVB_EP(i).oriSub & " jmp"
       'RVAT_LIST(o).StrD = i & " (" & PEexe.exeVB_SUBS(i).sName & ")"
       o = o + 1
    Next i
    'redimensionne précisément le tableau (pour eviter les entrées vide en fin)
    ReDim Preserve RVAT_LIST(1 To o) As RVATARGET
    RVAT_Len = o
       
    'liste des EP loaders (expérimental)
    cl = 0
    For i = 1 To UBound(PEexe.exeVB_EP())
        If PEexe.exeVB_EP(i).oriRva <> cl Then
            o = o + 1
            ReDim Preserve RVAT_LIST(1 To o) As RVATARGET
            RVAT_LIST(o).rva = PEexe.exeVB_EP(i).oriRva
            RVAT_LIST(o).StrS = "EP " & PEexe.exeVB_EP(i).oriSub & " load"
            RVAT_LIST(o).StrD = "EP " & PEexe.exeVB_EP(i).oriSub & " type " & Hex$(PEexe.exeVB_EP(i).oriCode)
            cl = PEexe.exeVB_EP(i).oriRva
        End If
    Next i
    RVAT_Len = o


    'étape 2 : analyse du listing désassemblé
    CC_FrmID = 1
    rva = UBound(ASM_LIST)
    For ta = 1 To rva - 1
    
        'lit
        iDump = ASM_LIST(ta).imDump
        bDump = (CLng(iDump) And &HFF&)
        
        jrva = ASM_LIST(ta).rvaJump
        sBuffer = ASM_LIST(ta).sUnAsm '= sBuffer
            
            'si c'est un 55h, c'est probablement un Sub. lequel ?
            If bDump = &H55 Then
                c = ASM_LIST(ta).rvaCode - PEexe.exeOPHEAD.ImageBase
                For i = 1 To cs
                    If PEexe.exeVB_SUBS(i).rvaCode = c Then
                        ASM_LIST(ta).sStruct = "Sub " & i
                        sub_en_cours = i
                        last_obj = 0
                        If (exeVB_SUBS(i).ObjFrom = 0 Or exeVB_SUBS(i).SubType = -1) And exeVB_SUBS(i).SubFrom >= 1 Then
                            ASM_LIST(ta).sData = PEexe.exeVB_MODULES(exeVB_SUBS(i).SubFrom).sName & ".user()"
                        ElseIf exeVB_SUBS(i).ObjFrom = -1 And exeVB_SUBS(i).SubType = -10 Then
                            ASM_LIST(ta).sData = "(module).user()"
                        ElseIf exeVB_SUBS(i).SubFrom = 0 And exeVB_SUBS(i).SubType = 1 Then
                            ASM_LIST(ta).sData = "Sub Main()"
                        Else
                            ASM_LIST(ta).sData = PEexe.exeVB_MODULES(exeVB_SUBS(i).SubFrom).sName & ".Obj-" & Hex$(exeVB_SUBS(i).SubType) & "_event()"
                        End If
                        Analyse.CC_FrmID = exeVB_SUBS(i).SubFrom
                        'ListView1.ListItems.Item(i).ListSubItems.Item(1).ForeColor = &HA04040
                        'mise à jour index sub
                        ip = ip + 1
                        ReDim Preserve ASM_SubIdx(ip)
                        ASM_SubIdx(ip) = ta
                        GoTo NAsm
                    End If
                Next i
                
                'sub non listé dans les structures VB : il s'agit d'un sub utilisateur dans un modules .bas compilé
                'frmPeExe.AddInfo "debug : sub présumé trouvé @ " & Hex$(c)
                'ASM_LIST(ta).sStruct = "Sub EP"
                'ASM_LIST(ta).sData = "(module).user()"
                'rajout dans la liste des subs
                'sub_en_cours = UBound(exeVB_SUBS()) + 1
                'ReDim Preserve PEexe.exeVB_SUBS(sub_en_cours)
                'With PEexe.exeVB_SUBS(sub_en_cours)
                '    .SubFrom = -1: .ObjFrom = -1
                '    .rvaCode = c
                '    .SubType = -10
                'End With
                
                'màj index
                ip = ip + 1
                ReDim Preserve ASM_SubIdx(ip)
                ASM_SubIdx(ip) = ta
            End If
            

            'détermine si le saut est un appel vers qq chose d'existant
            If bDump = &H68 Or bDump = &HE8 Or bDump = &HBA Or bDump = &H8B Or bDump = &HFF Or bDump = &HC7 Or bDump = &HB8 Then
                For i = 1 To RVAT_Len
                    If RVAT_LIST(i).rva = jrva Then
                        ASM_LIST(ta).sStruct = RVAT_LIST(i).StrS
                        ASM_LIST(ta).sData = RVAT_LIST(i).StrD
                        'mise à jour index api
                        ap = ap + 1
                        ReDim Preserve ASM_ApiIdx(ap)
                        ASM_ApiIdx(ap) = ta
                        'libère objet ?
                        If i = freeobj Then last_obj = 0
                    End If
                Next i
                
                'aucunes correspondances API. soit c'est une string, soit c'est un appel de définition de Form
                
                If i > RVAT_Len Then
                    'est-ce une string ?
                    jrva = jrva - PEexe.exeOPHEAD.ImageBase
                    If jrva > 4128 And jrva < PEexe.exeVB_CODEENTRY Then
                        Get #fp, jrva - 3, lDump
                        If lDump > 0 And lDump < 32000 Then
                            sBuffer = PEexe.ScanUnicode(fp, jrva + 1, lDump)
                            If sBuffer <> "" And Asc(sBuffer & " ") > 15 Then
                                ASM_LIST(ta).sStruct = "String"
                                ASM_LIST(ta).sData = Chr$(34) & sBuffer & Chr$(34)
                                'mise à jour index string
                                sp = sp + 1
                                ReDim Preserve ASM_StrIdx(sp)
                                ASM_StrIdx(sp) = ta
                            End If
                        End If
                    End If
                End If
                
                If i > RVAT_Len Then
                    'est-ce un form ?
                    For i = 1 To UBound(PEexe.exeVB_MODULES())
                        If (ASM_LIST(ta).rvaJump - PEexe.exeOPHEAD.ImageBase) = PEexe.exeVB_MODULES(i).RvaOffset Then
                            ASM_LIST(ta).sStruct = "Mod Load"
                            ASM_LIST(ta).sData = "Read Struct " & PEexe.exeVB_MODULES(i).sName & " >  " & i
                            If PEexe.exeVB_MODULES(i).lType = 98435 Then
                                CC_FrmID = i
                                Call WhoIsThatControl(&H2F8)
                            End If
                            Exit For
                        End If
                    Next i
                End If
                
                
            End If
            
            'appel à un objet/controle (experimental)
            If bDump = &HFF Then
                If (iDump And &HF000) = &H9000 Or (iDump And &HF000) = &H5000 Then
                    p = InStr(4, sBuffer, "+", vbBinaryCompare) '
                    If p > 0 Then
                    p = Val("&h" & Mid$(sBuffer, p + 1, InStr(5, sBuffer, "]", vbBinaryCompare) - p - 1))
                    If p < &H61C And p > &H2C4 Then
                        'objet
                        If Analyse.WhoIsThatControl(p) Then
                            ASM_LIST(ta).sStruct = "Ref Obj"
                            ASM_LIST(ta).sData = Hex$(p) & "h : " & PEexe.exeVB_FORMS(exeVB_MODULES(CC_FrmID).frmidx).sName & "." & Analyse.CC_ctrl.sName
                        Else
                            ASM_LIST(ta).sStruct = "Ref Obj"
                            ASM_LIST(ta).sData = Hex$(p) & "h : (?)"
                        End If
                        last_obj = p
                        op = op + 1
                        ReDim Preserve ASM_ObjIdx(op)
                        ASM_ObjIdx(op) = ta
                    ElseIf p <= &H2C4 And p > &H47 Then
                        'propriété d'un objet
                        If last_obj > 0 Then
                            If Analyse.CC_ctrlID = 0 Then
                                ASM_LIST(ta).sData = "Propriété ref_" & Hex$(last_obj) & "." & PEexe.exeVB_Prop(p)
                            Else
                                ASM_LIST(ta).sData = "Propriété " & Analyse.CC_ctrl.sName & "." & PEexe.exeVB_Prop(p) & " (" & Analyse.CC_ctrl.sType & ") "
                            End If
                        Else
                            If CC_FrmID > 0 Then
                                If exeVB_MODULES(CC_FrmID).frmidx > 0 Then ASM_LIST(ta).sData = "Propriété " & PEexe.exeVB_FORMS(exeVB_MODULES(CC_FrmID).frmidx).sName & "." & PEexe.exeVB_Prop(p) & " (form)"
                            Else
                                ASM_LIST(ta).sData = "Propriété " & PEexe.exeVB_Prop(p) & " (form)"
                            End If
                        End If
                        op = op + 1
                        ReDim Preserve ASM_ObjIdx(op)
                        ASM_ObjIdx(op) = ta
                    End If
                    End If
                End If
            End If
                
            
            
            'EP jmp (experimental) (gouffre de performances :/ )
            For i = 1 To cp
                If rva = PEexe.exeVB_EP(i).rvaJmp Then
                    ASM_LIST(ta).sStruct = "EP " & PEexe.exeVB_EP(i).oriSub & " jmp"
                End If
            Next i
            
NAsm:
    If ta Mod 12000 = 0 Then
        'indicateur de progression non bloquant
        frmPeExe.AddInfo "Analyse du listing à " & Int(ta / rva * 100) & "%...", True
    End If
    
    Next ta
            

End Sub


Function WhoIsThatControl(ByVal rid As Integer) As Boolean
'retrouve le nom d'un contrôle en fonction du code d'appel en hexa
Dim i As Long, t As Long
    If (CC_FrmID <= UBound(PEexe.exeVB_MODULES())) And CC_FrmID > 0 Then
        t = PEexe.exeVB_MODULES(CC_FrmID).frmidx
        If t > 0 Then
        For i = PEexe.exeVB_FORMS(t).DefPtr To UBound(PEexe.exeVB_CONTROL())
            If rid = PEexe.exeVB_CONTROL(i).frmID Then
            WhoIsThatControl = True
            CC_ctrl = PEexe.exeVB_CONTROL(i)
            CC_ctrlID = i
            'CC_FrmID = t
            Exit Function
            End If
        Next i
        End If
    End If
    CC_ctrlID = 0

End Function
