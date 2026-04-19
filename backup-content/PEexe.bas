Attribute VB_Name = "PEexe"
'=======================
' PE EXE
'********
'
'Par Proger
'aout 2003
'reprise septembre 2005
'
' Explore l'intérieur d'un fichier exécutable portable (PE), optimisé pour VB6
'
' sources : documentations trouvés sur wotsit.org à propos des PEs.
'  pour ce qui est du "format vb", c'est du 100% manuel avec editeur hexa :)
'  analyser les format de fichier c'est mon dada :)
'
'dire que je n'ai pas "du tout" utiliser les sources "décompilateur" que l'on trouve ici est là
'serai faux : en effet, elles m'ont donné un aperçu de ce que ce code devra minimum faire ;)
'néanmoins j'ai volontairement refusé d'utiliser du code de ces sources pour éviter de copier leurs bogues.
'
'Pour des raisons de copyright, ancienneté et droit d'auteur :
'L'auteur de cette source (connu sous le pseudonyme Proger) ne pourra en aucun cas être tenu pour responsable
'de l'utilisation de cette source dans le cadre de perte ou destruction de données, ingéniérie a rebourt, vol
'ou détournement de propriété intellectuel, contrefaçon ou plagiat.

DefLng A-Z
Option Explicit

'en-tête PE
Type PE_HEADER
    Signature As Long  ' "PE" en ansi
    CpuType As Integer
    Objects As Integer
    
    TimeDate As Long    'time-f (seconde écoulé depuis le 1/1/1970)
    PointerToSymbolTable As Long
    
    NumberOfSymbols As Long
    NThdrSize As Integer
    Flags As Integer    'conditions d'exécution du PE

End Type

'en-tête "optionnel" (mais toujours présent!)
Type OPTIONAL_HEADER
    
    SizeOfOptionalHeader As Integer
    LinkMajor As Byte
    LinkMinor As Byte
    Reserved1 As Long
    
    Reserved2 As Long
    Reserved3 As Long
    
    EntryPointRVA As Long   ' offset du début du code proprement dit
    Reserved4 As Long
    
    Reserved5 As Long
    ImageBase As Long   ' adressage relatif dans le code (en général 00400000h)
    
    ObjectAlign As Long
    FileAlign As Long
    
    OsMajor As Integer
    OsMinor As Integer
    UserMajor As Integer
    UserMinor As Integer
    
    SubSysMajor As Integer
    SubSysMinor As Integer
    Reserved6 As Long
    
    ImageSize As Long
    HeaderSize As Long
    
    FileCheckSum As Long
    SubSystemNT As Integer
    DLLflags As Integer
    
    StackReserveSize As Long
    StackCommitSize As Long
    
    HeapReserveSize As Long
    HeapCommitSize As Long
    
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long  'nombre de data directories (toujours 16)

End Type

'data directories
Type HEAD_DIRECTORIES
    '0 : fonctions exportés
    RvaEXPORT_TABLE As Long
    RvaTOTAL_EXPORT_DATA_SIZE As Long
    '1 : fonctions importés
    RvaIMPORT_TABLE As Long
    RvaTOTAL_IMPORT_DATA_SIZE As Long
    '2 : ressources (les copytight vb sont dans ce directory)
    RvaRESOURCE_TABLE As Long
    RvaTOTAL_RESOURCE_DATA_SIZE As Long
    '3
    RvaEXCEPTION_TABLE As Long
    RvaTOTAL_EXCEPTION_DATA_SIZE As Long
    '4
    RvaSECURITY_TABLE As Long
    RvaTOTAL_SECURITY_DATA_SIZE As Long
    '5
    RvaFIXUP_TABLE As Long
    RvaTOTAL_FIXUP_DATA_SIZE As Long
    '6
    RvaDEBUG_TABLE As Long
    RvaTOTAL_DEBUG_DIRECTORIES As Long
    '7 : copyright
    RvaIMAGE_DESCRIPTION As Long
    RvaTOTAL_DESCRIPTION_SIZE As Long
    '8
    RvaMACHINE_SPECIFIC As Long
    RvaMACHINE_SPECIFIC_SIZE As Long
    '9
    RvaTHREAD_LOCAL_STORAGE As Long
    RvaTOTAL_TLS_SIZE As Long
    '10
    RvaENTRY_LOAD_CONFIG As Long
    RvaTOTAL_ENTRY_LOAD_CONFIG_SIZE As Long
    '11
    RvaENTRY_BOUND_IMPORT As Long
    RvaTOTAL_ENTRY_BOUND_IMPORT_SIZE As Long
    '12 : table d'adresses d'importation (en rapport avec les fonctions importées)
    RvaENTRY_IAT As Long
    RvaTOTAL_ENTRY_IAT_SIZE As Long
    '13
    RvaENTRY_D13 As Long
    RvaTOTAL_ENTRY_D13_SIZE As Long
    '14
    RvaENTRY_D14 As Long
    RvaTOTAL_ENTRY_D14_SIZE As Long
    '15
    RvaENTRY_D15 As Long
    RvaTOTAL_ENTRY_D15_SIZE As Long

End Type

'section
Type OBJECT_TABLE
    oName(1 To 8) As Byte
    
    VirtualSize As Long
    VirtualAddress As Long
    
    SizeOfRawData As Long
    PointerToRawData As Long
    
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    oFlags As Long
    
End Type



'descripteurs :

'Importations
Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk As Long
    TimeDateStamp As Long
    ForwarderChain As Long
    RvaName As Long
    FirstThunk As Long
End Type
    Private exeIMPORT As IMAGE_IMPORT_DESCRIPTOR
    'déclarations pour analyse (n'existe pas dans l'exe pe)
    Type IMPORT_DLL_LOOKUP
        DllName As String
        HashPtr As Long
    End Type
    Type IMPORT_API_LOOKUP
        ApiName As String
        Address As Long
        VaTbl As Long
    End Type
    Public exeIMPORT_DLLNAME() As IMPORT_DLL_LOOKUP
    Public exeIMPORT_APINAME() As IMPORT_API_LOOKUP

    'DEASM : table de recherche inverse via call+jmp vers dll vb
    Public Type VB_APICALLS
        rva As Long
        ApiVbDefPtr As Long
    End Type
    Public exeVB6_APICALLS() As VB_APICALLS

'Exportation (en dev)
Type IMAGE_EXPORT_DESCRIPTOR
    Characteristics As Long
    TimeDateStamp As Long 'ou version
    NamePtr As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions_Rva As Long
End Type
    Type EXPORT_FUNCTIONS
        AddressOfFunction As Long
        AddressOfName As Long
        AddressOfNameOrdinal As Long
    End Type
    Private exeEXPORT_NAME() As String
    
'Ressources
Type IMAGE_RESOURCE
    Flags As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NumName As Integer
    NumIDentry As Integer
End Type
    Type IMAGE_RESOURCE_ENTRY
        RvaName_ID As Long
        RvaEntry As Long
    End Type
    Type RESOURCE_ENTRY
        RvaDATA As Long
        Size As Long
        CodePage As Long
        Reserved As Long
    End Type
    Private exeTBL_RESOURCE() As RESOURCE_ENTRY
    ' stockage pour le prog (n'est pas un type de structure contenu dans un .exe PE)
    Type RESOURCE_INFO
        RvaEntry As Long
        EntrySize As Long
        ResourceType As Long
    End Type
    Private exeRES_INFO() As RESOURCE_INFO
    Private exeRES_UNICODE() As String


'déclarations pour les contrôles vb (objets)
Type CONTROL_FORM
    sName As String  'nom de la form
    rvaPtr As Long   'point d'entrée de la form dans le fichier exe
    DefPtr As Long   'pointeur vers la liste control_def
    DefLen As Long   'nombre de contrôle contenu
End Type
    Type CONTROL_DEF
        sName As String  'nom "name" du contrôle
        id As Integer    'identifiant type de contrôle
        sType As String  'type de contrôle
        offset As Long   'offset physique (pour récupérer les propriétés/attributs)
        LenTr As Long    'nombre d'octets définissant l'objet
        frmID As Integer 'identifiant ordre d'apparition dans le form
    End Type
        Type CONTROL_IDTYPE 'type de contrôle (pure VB seulement)
            inID As Integer
            cType As String
        End Type
        Public vbDEFCTRL() As CONTROL_IDTYPE
    Type CONTROL_PROPERTY 'attribut les plus communs
        pTop As Long
        pLeft As Long
        pHeight As Long
        pWidth As Long
        pBackColor As Long
        pForeColor As Long
        pRank As Long 'sous-bloc de controle,...
        sCaption As String
    End Type
Public exeVB_CONTROL() As CONTROL_DEF
    Public exeVB_PROJECTNAME As String      'nom du projet :)
    Public exeVB_FORMS() As CONTROL_FORM
    Public exeVB_CTRL_PRP() As CONTROL_PROPERTY
    
'déclarations pour les API importé ("declare") vb
'en dev
Private Type APISTRUCT  'structure de déclaration d'une API tel que c'est physiquement dans l'exe
    sDll As Long
    sAPI As Long
    num As Long
    rvan As Long
End Type
Type APIDECLARE
    sName As String     'nom de l'api
    sDll As String      'nom de la DLL origine
    lFrom As Long       'index de la feuille dans laquel l'API est déclaré (en debug)
    RvaOffset As Long   'offset où l'api est déclaré dans l'exe
    RvaCall As Long     'offset demandant l'appel à cet API par le code compilé
End Type
    Public exeVB_API() As APIDECLARE



'déclaration pour un morceau de fonction compilé (en dev)
Public exeVB_CODEENTRY As Long 'pointeur vers le début des fonctions compilées
Public exeVB_CODELEN As Long   'longueur de l'ensemble des fonctions compilées
Public exeVB_CODEMAIN As Long  'point d'entrée vers la fonction appelé au démarrage (Sub Main)
'Table de strucutre de feuille (modules, forms...) tel qu'elle est enregistré physiquement
'dans l'exe compilé vb6. Les informations associés à la variables sont "ce qui semble être" et non pas "sûre"
Private Type VBTBLSTRUCT
    rvaS As Long    'pointeur vers définition des subs de la feuille
    fill1 As Long   'toujours &hffffffff
    rva2 As Long
    rva3 As Long
    rva4 As Long
    rva5 As Long
    rva6 As Long    'pointe vers le nom "Name" de la feuille (string ansi)
    nSubs As Long   'nombre de subs()
    rva7 As Long
    num2 As Long
    mType As Long   'type de feuille (modules, forms, classes)
    eot1 As Long
    'rva1 As Long    'pointeur d'association
End Type
Private Type VBPAGESTRUCT
    Starts As Long
    rva1 As Long
    null1 As Long
    rva2 As Long
    fill1 As Long
    null2 As Long
    rva3 As Long
    rva4 As Long
    null3 As Long
    value1 As Long
    null4 As Long
    null5 As Long
    null6 As Long
    rva5 As Long
    vone1 As Long
    rva6 As Long
    null7 As Long
    rva7 As Long
    vone2 As Long
    rva8 As Long
    null8 As Long
    rva9 As Long
    vthree As Long
    rva10 As Long
    numsub As Integer
    const1 As Integer
    const2 As Integer
    const3 As Integer
    subjmp As Long
End Type
    'type pour vbanalyse
    Type VBMODULE
        sName As String    'nom
        lType As Long      'type de feuille
        RvaOffset As Long  'rva debut structure dans fichier
        FullLen As Long    'longueur de la structure
        numsub As Long     'nombre de sub
        frmidx As Long     'index vers la table des forms/userctl si le module est un frm/uctl
    End Type
    Type VBSUB  'description d'un sub() utilisateur (programmé par)
        sName As String     'si possible
        sParams As String   'si possible
        rvaCode As Long     'point d'entrée dans le code compilé
        codelen As Long     'longueur du code compilé (jusqu'a une instruction RET)
        rvaEnd  As Long     'rva de fin du code
        SubType As Long     'type de sub
        SubFrom As Long     'origine du sub (feuille dans laquelle il est codé...)
        ObjFrom As Long     'objet d'origine du sub (s'il y a lieu)
    End Type
    Type VBEP  '"entry points" dans le code compilé
        oriSub As Long    'index du sub appelant
        oriCode As Long   'code du pattern
        oriRva As Long    'position du pattern
        rvaJmp As Long    'cible dans le code compilé
    End Type
        Public exeVB_MODULES() As VBMODULE
        Public exeVB_SUBS() As VBSUB
        Public exeVB_EP() As VBEP



'analyse du copyright-ressource (en dev)
Type RES_COPYRIGHT
    CompanyName As String
    FileDescription As String
    LegalCopyright As String
    LegalTrademarks As String
    ProductName As String
    FileVersion As String
    ProductVersion As String
    InternalName As String
    OriginalFilename As String
End Type
    Private exeVB_COPYRIGHT As RES_COPYRIGHT


'analyse des API VB6 importé
Type API_VBDEF
    rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Public exeVB6_APIDEF() As API_VBDEF

Public exeVB_Prop() As String 'liste des propriétés d'objets VB


'l'api qui sauve la vie :)
Private Declare Sub MemCpy Lib "Kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, From As Any, ByVal Length As Long)

Private exePEHEAD As PE_HEADER
Public exeOPHEAD As OPTIONAL_HEADER
Private exeHEADIR As HEAD_DIRECTORIES
Private exeOTABLE() As OBJECT_TABLE
Public exeISVB As Boolean
Public exeISPACKED As Boolean

Public exeVB_VBEP As Long 'point d'entrée des données VB

Public exeFILENAMElong As String  'nom du dernier fichier étudié
Public exeFILENAMEsize As Long    'taille du fichier
Public exeFILENAMEdir As String

'=================================================================
'=================================================================
'=================================================================

Sub OpenEXE_PK(ByVal FFname As String)
'cette fonction alternative essaye de récupérer des informations VB6 dans les exe dépacké manuellement.
'coder un tel sub, c'est vraiment "limite-limite" par rapport aux licenses...
Dim fp As Integer, pefp As Long
Dim PEl As Long, OPl As Long, HDl As Long, OTl As Long
Dim i As Long, j, n As Long, vb5 As Long, oep As Long, bcheck As Byte
Dim bArray() As Byte

    fp = FreeFile
    PEl = Len(exePEHEAD)
    OPl = Len(exeOPHEAD)
    HDl = Len(exeHEADIR)
    ReDim exeOTABLE(1)
    OTl = Len(exeOTABLE(1))
    
    ReDim bArray(1 To PEl + OPl + HDl)
    
    exeFILENAMElong = FFname
    Open FFname For Binary Access Read As #fp
        exeFILENAMEsize = LOF(fp)
        Get #fp, 61, pefp   'offset du début PE
        Get #fp, pefp + 1, bArray() 'récupère les octets
        'copie les octets dans les structures
        MemCpy exePEHEAD.Signature, bArray(1), PEl
        MemCpy exeOPHEAD.SizeOfOptionalHeader, bArray(PEl + 1), OPl
        MemCpy exeHEADIR.RvaEXPORT_TABLE, bArray(PEl + OPl + 1), HDl
        ReDim exeIMPORT_DLLNAME(0 To 0)
        
        'ces informations de header sont celle du pack... donc inutile...
        
        exeISVB = False
        For i = 2048 To LOF(fp)
            'recherche "VB5!"
            Get #fp, i, vb5
            If vb5 = 557138518 Then
                exeVB_VBEP = i - 1
                'recherche l'entrypoint qui a appelé "VB5!"
                n = i - 1 + exeOPHEAD.ImageBase
                For j = i To 2048 Step -1
                    Get #fp, j, oep
                    If oep = n Then
                        Get #fp, j - 1, bcheck
                        's'il s'agit bien de l'instruction JMP, c'est bien l'appel à "VB5!" donc l'entrypoint
                        If bcheck = &H68 Then
                            exeISVB = True
                            Call FindControl(fp, j - 2)
                            Call FindModules(fp, j - 2)
                            Call AssocieSubObj
                            GoTo WasVB
                        End If
                    End If
                Next j
                Exit For
            End If
        Next i
        
    Close #fp
        
    Exit Sub

WasVB:
    'servira...

End Sub

Function IsExe(FFname As String) As Boolean
'détermine si c'est un .exe ou non en cherchant "MZ".
Dim eHead As Integer

    Open FFname For Binary Access Read As #10
        Get #10, 1, eHead
    Close #10
    
    IsExe = (eHead = 23117)

End Function

Sub OpenEXE(ByVal FFname As String)
'attention, ca ne vérifie pas si l'exe est valide (bonjour l'erreur en cas de forcing :op )
Dim fp As Integer, pefp As Long
Dim PEl As Long, OPl As Long, HDl As Long, OTl As Long
Dim i As Long, n
Dim bArray() As Byte

    fp = FreeFile
    PEl = Len(exePEHEAD)
    OPl = Len(exeOPHEAD)
    HDl = Len(exeHEADIR)
    ReDim exeOTABLE(1)
    OTl = Len(exeOTABLE(1))
    
    ReDim bArray(1 To PEl + OPl + HDl)
    
    exeFILENAMElong = FFname
    Open FFname For Binary Access Read As #fp
        exeFILENAMEsize = LOF(fp)
        Get #fp, 61, pefp   'offset du début PE
        Get #fp, pefp + 1, bArray() 'récupère les octets
        'copie les octets dans les structures
        MemCpy exePEHEAD.Signature, bArray(1), PEl
        MemCpy exeOPHEAD.SizeOfOptionalHeader, bArray(PEl + 1), OPl
        MemCpy exeHEADIR.RvaEXPORT_TABLE, bArray(PEl + OPl + 1), HDl
        
        'récupère les "objets" (aka "sections", rien a voir avec les contrôles vb)
        ReDim exeOTABLE(1 To exePEHEAD.Objects)
        ReDim bArray(1 To OTl)
        exeISPACKED = False 'pour savoir si l'exe est protégé
        For i = 1 To exePEHEAD.Objects
            Get #fp, , bArray()
            MemCpy exeOTABLE(i).oName(1), bArray(1), OTl
            If Left$(PEobj_Name(exeOTABLE(i).oName()), 2) = "UP" Then exeISPACKED = True
        Next i
        
        
        
        'récupère les directories


        'importations
        If exeHEADIR.RvaIMPORT_TABLE > 0 Then
            'récupère les fonctions importés
                ReDim exeIMPORT_APINAME(1 To 1)
                ReDim exeIMPORT_DLLNAME(0 To 0)
            ReDim bArray(1 To exeHEADIR.RvaTOTAL_IMPORT_DATA_SIZE + 1)
            Get #fp, exeHEADIR.RvaIMPORT_TABLE + 1, bArray()
            
            'si le PE contient un objet ".idata", l'offset vers l'import est dedans (ex : vb5)
            For i = 1 To exePEHEAD.Objects
                If exeOTABLE(i).oName(2) = 105 And exeOTABLE(i).oName(3) = 100 Then
                    Get #fp, exeOTABLE(i).PointerToRawData + 1, bArray()
                    Exit For
                End If
            Next i
            
            'première dll :
            MemCpy exeIMPORT.OriginalFirstThunk, bArray(1), Len(exeIMPORT)
            i = 0: n = 0
            If exeIMPORT.OriginalFirstThunk <> 0 Then ReDim exeIMPORT_DLLNAME(1 To 1)
            Do While (exeIMPORT.OriginalFirstThunk <> 0)
                n = n + 1
                Call ParseImportTable(fp, n, exeIMPORT)
                i = Len(exeIMPORT) + i
                If (i > exeHEADIR.RvaTOTAL_IMPORT_DATA_SIZE) Then Exit Do 'antibug
                MemCpy exeIMPORT.OriginalFirstThunk, bArray(1 + i), Len(exeIMPORT)
                'boucle jusqu'a la dernière dll de la liste import
            Loop
            If n = 0 Then frmPeExe.AddInfo "Erreur de lecture des imports de DLL! fichier truqué!"
        End If
        
        'exporation (dll/API externe)
        ' non utilisé par les programmes VB / pas d'équivalent dans les formats PEs


        'ressources
        If exeHEADIR.RvaRESOURCE_TABLE > 0 Then
            'le point d'entrée (offset) vers les ressources est dans l'objet ".rsrc"
            For i = 1 To exePEHEAD.Objects
                If exeOTABLE(i).oName(2) = 114 And exeOTABLE(i).oName(3) = 115 Then
                    Call ParseResourceTable(fp, exeOTABLE(i).PointerToRawData + 1)
                    Exit For
                End If
            Next i
            'Call ParseResourceTable(FP, exeHEADIR.RvaRESOURCE_TABLE + 1)
        End If


        'vb seulement :
        If CheckVBexe(fp, exeOPHEAD.EntryPointRVA) Then
            Get #fp, exeOPHEAD.EntryPointRVA + 2, exeVB_VBEP
            
            'objets, contrôles
            Call FindControl(fp, exeOPHEAD.EntryPointRVA)
            'copyright
            'Call ParseCopyright(fp, exeTBL_RESOURCE(5).RvaDATA + 1, exeTBL_RESOURCE(5).Size)
            
            'module
            Call FindModules(fp, exeOPHEAD.EntryPointRVA)
            
            'table de points d'entrés (anciennement cru : sub)
            Call ParseVBEP(fp)
            
            'associe les subs trouvés aux objets/modules
            Call AssocieSubObj
            
        End If
        
    Close #fp
    
    
    
    
End Sub

Sub ParseCopyright(FilePointer As Integer, ByVal ep As Long, ByVal MaxLen As Long)
'je vous ai dit que c'est en dev !
Dim i, j, k
Dim Tstr As String
Dim iArray(1 To 50) As Integer

    ep = ep + 6
    Tstr = ScanUnicode(FilePointer, ep)
    If Tstr = "VS_VERSION_INFO" Then
        ep = ep + 92
        Tstr = ScanUnicode(FilePointer, ep)
        If Tstr = "VarFileInfo" Then
            ep = ep + 128: j = 10
            For i = 1 To 9
            Tstr = ScanUnicode(FilePointer, ep)
            Select Case Tstr
            Case "CompanyName"
                ep = ep + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.CompanyName = Tstr
            Case "FileDescription"
                ep = ep + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.FileDescription = Tstr
            Case "LegalCopyright"
                ep = ep + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.LegalCopyright = Tstr
            Case "LegalTrademarks"
                ep = ep + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.LegalTrademarks = Tstr
                j = 8
            Case "ProductName"
                ep = ep + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.ProductName = Tstr
                j = 10
            Case "FileVersion"
                ep = ep + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.FileVersion = Tstr
            Case "ProductVersion"
                ep = ep + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.ProductVersion = Tstr
            Case "InternalName"
                ep = ep + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.InternalName = Tstr
                j = 8
            Case "OriginalFilename"
                ep = ep + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, ep)
                exeVB_COPYRIGHT.OriginalFilename = Tstr
            End Select
            ep = ep + Len(Tstr) * 2 + j
            Next i
        End If
    End If

End Sub
Function ScanUnicode(fp As Integer, ByVal offset As Long, Optional EOS As Long = 32767) As String
'renvoi la chaine commençant à l'offset Ofs lorsque elle est de type unicode dans le fichier
Dim B1 As Byte, B2 As Byte, i As Long
i = 1

    Get #fp, offset, B1
    Get #fp, offset + 1, B2
    
    Do Until ((CLng(B1) + CLng(B2)) = 0) Or (i > EOS)
        ScanUnicode = ScanUnicode & " "
        MidB$(ScanUnicode, i, 1) = Chr$(B1)
        MidB$(ScanUnicode, i + 1, 1) = Chr$(B2)
        i = i + 2
        offset = offset + 2
        Get #fp, offset, B1
        Get #fp, offset + 1, B2
        If (CLng(B1) + CLng(B2)) > 255 Then
            ScanUnicode = ""
            Exit Function
        End If
    Loop

End Function

Function CheckVBexe(FilePointer As Integer, ByVal EntryPoint As Long) As Boolean
'vérifie qu'il s'agit d'un programme vb
Dim bB As Byte, bL As Long, Rel As Long

    Get #FilePointer, EntryPoint + 1, bB
    
    If bB = 104 Then 'instruction "push"
        Get #FilePointer, EntryPoint + 2, bL
        Rel = bL - exeOPHEAD.ImageBase
        Get #FilePointer, Rel + 1, bL
        
            CheckVBexe = (bL = 557138518) ' "VB5!"
            exeISVB = CheckVBexe
            
    End If
        
End Function

Sub FindModules(FilePointer As Integer, ByVal ProgEntryPoint As Long)
Dim PtrStr, PtrTbl, PtrBas, PtrSubMain
Dim TBLdeb, ASMdeb, ASMend
Dim i, j, k, l, m
Dim bPad As Byte
Dim TblLst(1 To 4) As Long
Dim TblDef(1 To 22) As Long
Dim TblStruct() As VBTBLSTRUCT
Erase exeVB_SUBS
ReDim exeVB_SUBS(0)

    'récupère le pointeur vers "VB5!"
    Get #FilePointer, ProgEntryPoint + 2, PtrStr
    
    'récupère le pointeur vers les defs de situation
    PtrTbl = PtrStr - exeOPHEAD.ImageBase + 49
    Get #FilePointer, PtrTbl, PtrBas
    
    'récupère le pointeur vers le sub Main() (départ d'exécution) s'il existe (sinon = 0)
    Get #FilePointer, PtrTbl - 4, PtrSubMain
    If PtrSubMain > exeOPHEAD.ImageBase Then
        'Sub Main()
        PtrSubMain = PtrSubMain - exeOPHEAD.ImageBase
        exeVB_CODEMAIN = PtrSubMain
        ReDim exeVB_SUBS(0 To 1)
        exeVB_SUBS(1).sName = "Main"
        exeVB_SUBS(1).SubType = 1
        exeVB_SUBS(1).rvaCode = exeVB_CODEMAIN
        exeVB_SUBS(1).SubFrom = 0
    Else
        exeVB_CODEMAIN = 0
    End If
    
    'récupère les 3 pointeurs de situations
    PtrBas = PtrBas - exeOPHEAD.ImageBase
    Get #FilePointer, PtrBas + 5, TblLst()
    
    TBLdeb = TblLst(1) - exeOPHEAD.ImageBase    'tables de structure
    ASMdeb = TblLst(3) - exeOPHEAD.ImageBase    'début du code ASM compilé
    ASMend = TblLst(4) - exeOPHEAD.ImageBase    'fin du code ASM compilé
    'le début du code compilé est décalé avec E9h comme padding, le vrai début est après E9h
    'la fin du code compilé est marqué avec 9Eh. Idem, il s'agit d'un padding a supprimer.
    i = ASMdeb: bPad = &HE9
    Do While bPad = &HE9
        i = i + 1
        Get #FilePointer, i, bPad
    Loop
    exeVB_CODEENTRY = i - 1
    exeVB_CODELEN = ASMend - i
    
    'récupération des tables de structure
    Get #FilePointer, TBLdeb + 1, TblDef()
    'le nombre de table de structure est un Integer, pour le récupérer a partir d'un Long, on fait un bitmasking
    k = (TblDef(12) And &HFFFF0000) / 65536
    ReDim TblStruct(1 To k)
    ReDim exeVB_MODULES(1 To k): i = 1
    Get #FilePointer, TblDef(13) - exeOPHEAD.ImageBase + 1, TblStruct()
    'initialise le tableau local s'il y a des API
    ReDim exeVB_API(0)
    
    'récupère les informations sur les feuilles dans l'exe
    For i = 1 To k
        exeVB_MODULES(i).lType = TblStruct(i).mType
        exeVB_MODULES(i).numsub = TblStruct(i).nSubs 'pas forcément précis => voir ParseVBSubs()
        exeVB_MODULES(i).sName = ScanString(FilePointer, TblStruct(i).rva6 - exeOPHEAD.ImageBase + 1)
        exeVB_MODULES(i).RvaOffset = TblStruct(i).rvaS - exeOPHEAD.ImageBase 'TblDef(22) - exeOPHEAD.ImageBase + ((i - 1) * 48)
        'Debug.Print Hex$(exeVB_MODULES(i).RvaOffset)
        'IMPORTANT : ce RVA est le point d'entré vers des tables et RVA...
        '...qui contiennent API ainsi que point d'entrée des SUBS dans le code !!
        If exeVB_MODULES(i).lType <> 98305 Then 'si le type est un module .bas, il n'y a pas de table.
            ParseVBSubs FilePointer, i, exeVB_MODULES(i).RvaOffset, exeVB_MODULES(i).numsub
        End If
        'Attention : seul les subs des évènements objets sont listés.
        'Les subs indépendants ne peuvent être trouvés qu'en analysant le code compilé désassemblé
        'un sub commence toujours par : 55 8B EC 6A/83
        ' 55    = push ebp
        ' 8B EC = mov ebp, esp
        ' 83 () = sub esp, (byte) ==> dans le cas d'un sub objet
        ' 6A () = push (byte) ==> dans le cas d'un sub indépendant
        
        'If TblStruct(i).rva1 <= exeOPHEAD.ImageBase Then Exit For 'mouais
        PtrBas = TblStruct(i).rvaS - exeOPHEAD.ImageBase + 1
        
        GoSub ScanTab
        
    Next i
    
    Call ParseCode(FilePointer)
    
    Exit Sub

    
ScanTab:
    'recherche de déclarations d'API via declare sub
    Get #FilePointer, PtrBas, m
    Get #FilePointer, PtrBas - 8, l
    Get #FilePointer, PtrBas - 4, j
    If l > exeOPHEAD.ImageBase Then 'mouais (antibug)
    l = l - exeOPHEAD.ImageBase + 1
    If l > LOF(FilePointer) Then Return
    'recherche s'il n'y a pas des déclarations dans la feuille en cours : API, strings, ...
    Do
        Get #FilePointer, l, PtrStr
        Get #FilePointer, l + 4, PtrTbl
        l = l + 8
        Select Case PtrStr
            Case 6
                '???
            Case 7
                'déclaration d'API
                Call ParseDeclares(FilePointer, PtrTbl - exeOPHEAD.ImageBase + 1, i)
            Case Else
                Exit Do
        End Select
    Loop
    End If
Return

End Sub

Sub ParseDeclares(fp As Integer, ByVal DeclareEntryPoint As Long, ByVal FormOrigin As Long)
'récupère les API déclaré sous VB et les classe par ordre de dll d'origine
Dim Tapi As APISTRUCT
Dim i, j, k
    
    Get #fp, DeclareEntryPoint, Tapi
    j = UBound(exeVB_API()) + 1

    ReDim Preserve exeVB_API(j)
    exeVB_API(j).RvaOffset = DeclareEntryPoint - 1
    exeVB_API(j).sName = ScanString(fp, Tapi.sAPI - exeOPHEAD.ImageBase + 1)
    exeVB_API(j).sDll = ScanString(fp, Tapi.sDll - exeOPHEAD.ImageBase + 1)
    exeVB_API(j).lFrom = FormOrigin
    exeVB_API(j).RvaCall = DeclareEntryPoint - 1 + 24
    
'    i = UBound(exeAPI_DLL())
'    If i = 0 Then
'        ReDim Preserve exeAPI_DLL(1)
'        exeAPI_DLL(1).DefPtr = Tapi.sDll
'        exeAPI_DLL(1).sName = ScanString(FP, Tapi.sDll - exeOPHEAD.ImageBase + 1)
'        exeAPI_DLL(1).ApiPtr = j
'    Else
'        For k = 1 To i
'            If exeAPI_DLL(k).DefPtr = Tapi.sDll Then Exit Sub
'        Next k
'        ReDim Preserve exeAPI_DLL(i + 1)
'        exeAPI_DLL(i + 1).DefPtr = Tapi.sDll
'        exeAPI_DLL(i + 1).sName = ScanString(FP, Tapi.sDll - exeOPHEAD.ImageBase + 1)
'        exeAPI_DLL(i + 1).ApiPtr = j
'    End If

End Sub

Sub FindControl(FilePointer As Integer, ByVal ProgEntryPoint As Long)
'recherche les blocs d'infos sur contrôles (les forms)
Dim i, j, k, PtrForm
Dim NumForm As Integer
Dim ObjOffset As Long
ReDim exeVB_CONTROL(1 To 1) 'purge


    'récupère le nom du projet
    exeVB_PROJECTNAME = ScanString(FilePointer, ProgEntryPoint + 61)

    'récupère le pointeur vers "VB5!"
    Get #FilePointer, ProgEntryPoint + 2, i
    
    'récupère un indice après "VB5!" indiquant le nombre de forms dans le projet
    ObjOffset = i - exeOPHEAD.ImageBase + 69
    Get #FilePointer, ObjOffset, NumForm
    
    'récupère le pointeur vers la table des forms inclusent dans le exe vb
    Get #FilePointer, ObjOffset + 8, PtrForm
    PtrForm = PtrForm - exeOPHEAD.ImageBase
    
    'récupère les différents pointeurs vers les bloc de forms
    ReDim exeVB_FORMS(1 To NumForm)
    PtrForm = PtrForm + 72 '(la table pointé recèle de diverses infos, dont en offset 72, le pointeur)
    For i = 1 To NumForm
        
        Get #FilePointer, PtrForm + 1, j
        exeVB_FORMS(i).rvaPtr = j - exeOPHEAD.ImageBase
        PtrForm = PtrForm + 80
    Next i
    
    'décomposition des forms
    ParseControl FilePointer, exeVB_FORMS()
    
    'récupération des attributs
    ParseControlParams FilePointer, exeVB_CONTROL()
    
   

End Sub


Sub ParseResourceTable(FilePointer As Integer, ByVal OffsetStart As Long)
'lit l'arbre conteneur de ressources PE, et récupère la taille et l'offset de chaque ressource
Dim i, j, k, l, m, n
Dim tArray() As Byte, lS As Long, lE As Long
Dim OffsetRes As Long, OffsetSub As Long, OffsetD As Long
Dim exeRESOURCE As IMAGE_RESOURCE
Dim BaseRS As IMAGE_RESOURCE_ENTRY
Dim TempRS As IMAGE_RESOURCE_ENTRY, SBtempRS As IMAGE_RESOURCE_ENTRY
Dim exeTBL_SUBDIRECTORY() As IMAGE_RESOURCE

    OffsetRes = OffsetStart
    lS = Len(exeRESOURCE)
    ReDim tArray(1 To lS)
    ReDim exeTBL_RESOURCE(1 To 1)
    lE = Len(BaseRS)

    'scanner de la table de ressources
    Get #FilePointer, OffsetRes, tArray()
    MemCpy exeRESOURCE.Flags, tArray(1), lS
    If exeRESOURCE.Flags > 0 Then Exit Sub
    OffsetRes = OffsetRes + lS
    j = 0: l = 0
    For i = 1 To exeRESOURCE.NumIDentry 'nombre de subdir "primaire"
        Get #FilePointer, OffsetRes, tArray()
        MemCpy BaseRS.RvaName_ID, tArray(1), lE
        
            OffsetSub = OffsetStart + (BaseRS.RvaEntry And &H7FFFFFFF)
            Get #FilePointer, OffsetSub, tArray()
            j = j + 1
            ReDim Preserve exeTBL_SUBDIRECTORY(1 To j)
            MemCpy exeTBL_SUBDIRECTORY(j).Flags, tArray(1), lS
            
                OffsetSub = OffsetSub + lS
                n = j
                For k = 1 To exeTBL_SUBDIRECTORY(n).NumIDentry 'nombre d'entrée dans le subdir i
                    Get #FilePointer, OffsetSub, tArray()
                    MemCpy TempRS.RvaName_ID, tArray(1), lE
                    If (TempRS.RvaEntry And &H80000000) = 0 Then
                        'ressource direct
                        l = l + 1
                        ReDim Preserve exeTBL_RESOURCE(1 To l)
                        If TempRS.RvaEntry < 1 Then Exit For 'antibug
                        Get #FilePointer, TempRS.RvaEntry, tArray()
                        MemCpy exeTBL_RESOURCE(l).RvaDATA, tArray(1), lE
                    Else
                        'sub-subdir :(
                        
                        OffsetD = OffsetStart + (TempRS.RvaEntry And &H7FFFFFFF)
                        Get #FilePointer, OffsetD, tArray()
                        j = j + 1
                        ReDim Preserve exeTBL_SUBDIRECTORY(1 To j)
                        MemCpy exeTBL_SUBDIRECTORY(j).Flags, tArray(1), lS
                        OffsetD = OffsetD + lS
                        For m = 1 To exeTBL_SUBDIRECTORY(j).NumIDentry
                            Get #FilePointer, OffsetD, tArray()
                            MemCpy SBtempRS.RvaName_ID, tArray(1), lE
                            If (SBtempRS.RvaEntry And &H8000000) = 0 Then
                                l = l + 1
                                ReDim Preserve exeTBL_RESOURCE(1 To l)
                                Get #FilePointer, OffsetStart + SBtempRS.RvaEntry, tArray()
                                MemCpy exeTBL_RESOURCE(l).RvaDATA, tArray(1), lE
                                OffsetD = OffsetD + lE
                            Else
                                'il est possible qu'il y ai encore 1 niveau de subdir
                                'mais ce n'est pas le cas dans un exe vb6
                            End If
                        Next m
                        
                    End If
                    OffsetSub = OffsetSub + lE
                Next k
        
        OffsetRes = OffsetRes + lE
    Next i

End Sub

Sub ParseImportTable(FilePointer As Integer, Idx As Long, ImportTable As IMAGE_IMPORT_DESCRIPTOR)
'récupère les API importés.
Dim i As Long, j As Long
    
    
    i = UBound(exeIMPORT_DLLNAME())
    If Idx > i Then ReDim Preserve exeIMPORT_DLLNAME(1 To Idx)
    
        exeIMPORT_DLLNAME(Idx).DllName = ScanString(FilePointer, ImportTable.RvaName + 1)
        exeIMPORT_DLLNAME(Idx).HashPtr = UBound(exeIMPORT_APINAME())
        
    Call ScanTable(FilePointer, ImportTable.FirstThunk + 1, ImportTable.OriginalFirstThunk + 1, exeIMPORT_APINAME())
    
End Sub

Private Sub ParseVBSubs(ByVal fp As Integer, ByVal ModuleFrom As Long, ByVal RvaOffs As Long, ByRef nSubs As Long)
Dim i, j, k, p, r, v, d
Dim sAsm As String, bAsm(1 To 10) As Byte, iAsm As Integer
Dim TblPage As VBPAGESTRUCT
    
    Get #fp, RvaOffs + 1, TblPage
    If (TblPage.Starts And &HFFFF&) <> 1 Then Exit Sub
    
    'nombre de pointeur de subs = TblPage.numsub
    'rva vers les pointeurs = TblPage.subjmp
    r = TblPage.subjmp - exeOPHEAD.ImageBase + 1
    
    j = UBound(exeVB_SUBS)
    'ReDim Preserve exeVB_SUBS(0 To j + TblPage.numsub)
    nSubs = TblPage.numsub
    
    For i = 0 To TblPage.numsub - 1
    j = j + 1
        d = r + i * 4
        If d > 1024 And d < exeFILENAMEsize Then
            Get #fp, r + i * 4, p
            'p pointe vers le code asm référant objet est jmp vers le code asm du sub
            d = p - exeOPHEAD.ImageBase + 1 - 8
            If d > 1024 And d < exeFILENAMEsize Then
                Get #fp, d, bAsm()
                'fix : le pointeur n'indique pas toujours un déclarateur de sub (sub + jmp)
                If bAsm(1) = 129 Then
                    ReDim Preserve exeVB_SUBS(0 To j)
                    sAsm = unASM.CodeToStr(bAsm(), unASM.GetVASM(TblPtrASM(bAsm(1)), bAsm(1)), p, k, v)
                    exeVB_SUBS(j).SubFrom = ModuleFrom
                    exeVB_SUBS(j).SubType = Val("&h" & Right$(sAsm, 8))
                    Get #fp, p - exeOPHEAD.ImageBase + 1, bAsm()
                    sAsm = unASM.CodeToStr(bAsm(), unASM.GetVASM(TblPtrASM(bAsm(1)), bAsm(1)), p, k, v)
                    'exeVB_SUBS(j).rvaCode = Val("&h" & Right$(sAsm, 8)) - exeOPHEAD.ImageBase
                    exeVB_SUBS(j).rvaCode = v - exeOPHEAD.ImageBase
                Else
                    'mauvais pointage du sub / sub faux / bug
                    j = j - 1
                End If
            End If
        Else
        Stop
        End If
    Next i

End Sub

Private Sub ParseVBEP(fp As Integer)
'recherche les points d'entrée dans le code compilé
'+ analyse la table des saut absolu (jmp FF25h) pointant vers les API vb importés
Dim i, j, k, cn, ce, sBe, rvao, ct
Dim t1, t2, t3, t4
Dim lj As Integer
t1 = 1

    i = 0
    ReDim exeVB_EP(1)
    ce = UBound(exeIMPORT_APINAME())    'pointeur fin des import d'api = début du tableau des points d'entrées
    cn = 4096 + (4 * ce) + 1
    
    'ETAPE 1
    'analyse de la table des points d'entrés dans le code compilé (faussement appelé "table des subs" antérieurement
    ct = 0
    Do
DebLoop:
        cn = cn + 4
        Get #fp, cn, j
        rvao = cn - 1
        If (j And &H25FF&) = &H25FF& Then 'fin de la liste (repéré par un 25ffh ...)
            Get #fp, cn + 2, t1
            If (t1 - exeOPHEAD.ImageBase) < cn And (t1 - exeOPHEAD.ImageBase) > 0 Then Exit Do
        End If
        ct = ct + 1
        
        'récupère la liste.
        'pattern 1 : 0x 00 08 00 (avec x impair) définition de Sub Object, avec x l'identifieur
        ' ===> groupement de 3 rva ou 00
        'pattern 2 : 0x 00 04 00 (avec x pair)
        ' ===> groupement de 3 00 ou rva
        'pattern 3 : xx 00 14 00
        ' ===> gorupement de 5 rva ou 00, suivi de pattern 4
        'pattern 4 : nn 00 00 00 (nn)
        ' ===> rva (long) répèté nn fois.
        'pattern 5 : 00 00 xx 4x
        ' ===> marque une autre section (autre feuille ?)
        'pattern 6 : xx 00 08 00 (rare, xx = 85)
        ' ===> 1 rva, 00, 1 rva
        '
        'pattern autres possible : xx xx 0F F3, xx xx CF F3, xx xx 3F F3, ...
        
        'détecteur de pattern 1 à 3
        t1 = (j And &H1F00FF)
        t2 = (j And &HFF)
        If (j > 0) And (t1 = j) And ((t1 - t2) > 0) Then 'evite 00h AND confirme pattern 1/2/3 AND evite pattern 4
            If (j And &H80001) = t1 Then
                'pattern 1 avec 1 seule entrée
                Get #fp, cn + 4, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                cn = cn + 4
            ElseIf ((j And &H8000F) = t1) Or ((j And &H4000F) = t1) Then
                'pattern 1 ou pattern 2
                Get #fp, cn + 4, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 8, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 12, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                cn = cn + 12
                
            ElseIf (j And &H1400FF) = t1 Then
                'pattern 3
                Get #fp, cn + 4, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 8, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 12, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 16, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 20, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                
                cn = cn + 20
                Get #fp, cn + 4, t2
                If t2 > 10240 Then
                    'bug à tracer
                    frmPeExe.AddInfo "Problème dans ParseVBEP() : t2 = " & Hex$(t2) & "h, ignoré."
                    Exit Do
                End If
                For k = 1 To t2
                    Get #fp, cn + 4 + (4 * k), t4
                    GoSub AddNSub
                Next k
                cn = cn + 4 + (4 * t2)
            
            ElseIf t2 = 133 Then
            'pattern 6
                Get #fp, cn + 4, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                Get #fp, cn + 12, t4
                If t4 > 0 Then
                    GoSub AddNSub
                End If
                cn = cn + 12
          
            Else
                'nouveau pattern ???
                't2 = Val(Hex(j))
                'Stop
            End If
        End If
    
    
    Loop
    
    
    'ETAPE 2
    'analyse des JMP DWORD PTR vers les APIs (FF 25 [long])
    
    If t2 > 10240 Then cn = 4096 + (4 * ce) + 1 'bug retro controle
    cn = cn - 2
    Do
        cn = cn + 2
        Get #fp, cn, lj
    Loop Until lj = &H25FF
    'les FF25& sont les JMP vers l'appel d'une fonction de la DLL VB.
    t3 = UBound(exeIMPORT_APINAME())
    i = 0
    Do
        Get #fp, cn, lj
        If lj = &H25FF Then 'jmp dword ptr ...
            Get #fp, cn + 2, j
            'jmp dword ptr [j]
            If j < exeOPHEAD.ImageBase Then
                j = Val("&h" & Hex$(j) & "0000")
            End If
            Get #fp, (j - exeOPHEAD.ImageBase) + 1, t1  'fait le jmp et recup l'addresse de l'api appelée.
            
            i = i + 1
            ReDim Preserve exeVB6_APICALLS(1 To i) As VB_APICALLS
            'cherche le nom de l'api correspondant (ou non-correspondance !).
            For t2 = 1 To t3
                If t1 = exeIMPORT_APINAME(t2).Address Then
                    exeVB6_APICALLS(i).ApiVbDefPtr = t2
                    exeVB6_APICALLS(i).rva = cn - 1
                    Exit For
                End If
            Next t2
            
        Else
            Exit Do
        End If
        
        cn = cn + 6
    Loop
    
    
    
Exit Sub

AddNSub:
    i = i + 1
    ReDim Preserve exeVB_EP(i)
    exeVB_EP(i).rvaJmp = t4 '- exeOPHEAD.ImageBase
    exeVB_EP(i).oriCode = j
    exeVB_EP(i).oriRva = rvao + exeOPHEAD.ImageBase
    exeVB_EP(i).oriSub = ct
    
Return
    
End Sub

Private Sub ParseCode(ByVal fp As Integer)
'réalise une analyse primaire du code compilé
Dim i, j, lvbs, va
Dim aByte() As Byte
ReDim aByte(PEexe.exeVB_CODELEN)
lvbs = UBound(PEexe.exeVB_SUBS())

    Get #fp, PEexe.exeVB_CODEENTRY + 1, aByte()
    

    For i = 1 To PEexe.exeVB_CODELEN
        If aByte(i) <> &H55 Then 'recherche des subs non listé dans les tables de structures VB
        Else
            If i < (PEexe.exeVB_CODELEN - 8) Then
                If aByte(i + 1) = &H8B And aByte(i + 2) = &HEC Then
                    va = PEexe.exeVB_CODEENTRY + i
                    For j = 1 To lvbs
                        If exeVB_SUBS(j).rvaCode = va Then GoTo Known
                    Next j
                    'nouveau sub trouvé
                    lvbs = lvbs + 1
                    ReDim Preserve exeVB_SUBS(lvbs)
                    exeVB_SUBS(lvbs).rvaCode = va
                    exeVB_SUBS(lvbs).SubType = -10
                    exeVB_SUBS(lvbs).SubFrom = -1
                    exeVB_SUBS(lvbs).ObjFrom = -1
                    
                End If
            End If
        End If
Known:
    Next i
                        

End Sub

Private Sub AssocieSubObj()
'associe les subs trouvés aux objets/controles trouvés
Dim i, j, k
Dim v, w

'exeVB_SUBS()
'PEexe.exeVB_CONTROL()
'PEexe.exeVB_MODULES()


    'mise en relation entre les FORMS et MODULE_FORM NAME (rafistolage ...)
    v = UBound(exeVB_FORMS())
    w = UBound(exeVB_MODULES())
    For i = 1 To w
        For j = 1 To v
            If exeVB_MODULES(i).sName = exeVB_CONTROL(exeVB_FORMS(j).DefPtr).sName Then Exit For
        Next j
        If j <= v Then
            exeVB_FORMS(j).sName = exeVB_MODULES(i).sName
            exeVB_MODULES(i).frmidx = j
        End If
    Next i

    'EN DEV !! buggé à mort
    For i = 1 To UBound(exeVB_SUBS())
    
        'retrouve l'index dans la liste des objets a partir du nom du form
        If exeVB_SUBS(i).SubFrom <= 0 Then Exit For
        For j = 1 To UBound(exeVB_FORMS())
            If exeVB_MODULES(exeVB_SUBS(i).SubFrom).sName = exeVB_CONTROL(exeVB_FORMS(j).DefPtr).sName Then Exit For
        Next j
        If j > UBound(exeVB_FORMS()) Then Exit For
        
        'principe : SubType contient soit un 00000033, 00000037 soit un FFFFFFFF
        'si v = FFFF, alors le sub est indépendant d'un objet, mais contenu dans le form
        'si v = 33h ou 37h, alors on soustrait 32h ce qui laisse 1 et 4 : c'est l'index relatif de l'objet associé
        'pour faire la correspondance entre l'index relatif et la liste d'objet trouvé par ParseControl(), on divise par 4
        v = exeVB_SUBS(i).SubType
        If v = &HFFFFFFFF Then
            exeVB_SUBS(i).ObjFrom = 0
        Else
            'EN DEBUG
            'v = (v - 51) / 4
            'If v = 0 Then
            '    exeVB_SUBS(i).ObjFrom = exeVB_FORMS(j).DefPtr
            'Else
            '    exeVB_SUBS(i).ObjFrom = exeVB_FORMS(j).DefPtr + v ' exeVB_FORMS(j).DefLen - v
            'End If
            
            'exeVB_SUBS(i).ObjFrom = -1
            'w = v + &H2F8 - (&H33 + &H2F8) - (&H33 + 4 * (exeVB_FORMS(j).DefLen - 1))
            'w = (Abs(w) + 709)
            
            'exeVB_SUBS(i).ObjFrom = v + exeVB_FORMS(j).DefLen - 1
            For k = exeVB_FORMS(j).DefPtr To exeVB_FORMS(j).DefPtr + exeVB_FORMS(j).DefLen - 1
                If (exeVB_CONTROL(k).frmID - 709) = v Then
                    exeVB_SUBS(i).ObjFrom = k
                End If
            Next k
        End If
    Next i
    
    'Associe les modules de classes, crée des controles virtuelles correspondant au sub
    'v = UBound(exeVB_FORMS())
    'For i = 1 To UBound(exeVB_MODULES())
    '    If exeVB_MODULES(i).lType = 1146883 Then
    '        v = v + 1
    '        ReDim Preserve exeVB_FORMS(1 To v)
    '        exeVB_MODULES(i).frmidx = v
    '        exeVB_FORMS(v).sName = exeVB_MODULES(i).sName
    '
    '    End If
    'Next i
    
    
End Sub

Function ScanString(fp As Integer, ByVal offset As Long) As String
'scanne une chaine de caractère ANSI, se termine par un 8-bits NULL
Dim b As Byte
If offset < 256 Then Exit Function
Get #fp, offset, b
Do Until b = 0
    ScanString = ScanString & Chr$(b)
    offset = offset + 1
    Get #fp, offset, b
Loop

End Function

Sub VBfunc_Description_Init(ByVal fRes As String)
'charge la liste de définition des API VB6
Dim lfp As Integer, i As Long
Dim sAdr As String, sOrd As String, sName As String, sDef As String
lfp = FreeFile
Erase exeVB6_APIDEF()

    Open fRes For Input Access Read As #lfp
        i = 0
        Do
        i = i + 1
            Input #lfp, sAdr, sOrd, sName, sDef
            If LCase$(sAdr) <> "eof" Then
                'If sOrd > 700 Then Stop
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).rva = Val("&H" & sAdr)
                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub

Function VBfunc_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'renvoi la fonction vb associé à l'api vb appelé
Dim i As Long


If inOrdinal > 99 Then ' And inAPIname = "" Then
    'par ordinal :
    For i = 1 To UBound(exeVB6_APIDEF())
        If exeVB6_APIDEF(i).Ordinal = inOrdinal Then
            VBfunc_Description = exeVB6_APIDEF(i).uDescr
            outRName = exeVB6_APIDEF(i).uName
            Exit Function
        End If
    Next i

Else
    'par nom:
    For i = 1 To UBound(exeVB6_APIDEF())
        If exeVB6_APIDEF(i).uName = inAPIname Then
            VBfunc_Description = exeVB6_APIDEF(i).uDescr
            Exit Function
        End If
    Next i
End If

VBfunc_Description = "API inconnu / non présente dans msvbvm60.dll / erreur"

End Function

Private Sub ScanTable(fp As Integer, ByVal OffsetADR As Long, ByVal OffsetSTR As Long, ByRef outADRarray() As IMPORT_API_LOOKUP)
'scanne une table d'addresse, se termine par un 32-bits NULL
'OffsetADR : position de la table d'adressage (argument d'appel vers la DLL importé)
'OffsetSTR : position dans la table des noms
Dim l As Long, i As Long, s As Long

    
    i = UBound(outADRarray()) - 1
    Get #fp, OffsetADR, l
    Do
        i = i + 1
        ReDim Preserve outADRarray(1 To i)
        outADRarray(i).Address = l
        Get #fp, OffsetSTR, s
        outADRarray(i).VaTbl = OffsetADR - 1
        If (s And &H80000000) = 0 Then
            'importation par nom
            outADRarray(i).ApiName = ScanString(fp, s + 3)
        Else
            'importation par ordinal
            outADRarray(i).ApiName = "!ordinal : " & (s And &H7FFFFFFF)
        End If
        
        'récupère la prochaine addresse
        OffsetSTR = OffsetSTR + 4
        OffsetADR = OffsetADR + 4
        Get #fp, OffsetADR, l
    Loop Until l = 0

End Sub



Sub PrintRessources(ByRef Alist As ListBox)
'affiche les ressources PE dans une liste
Dim i

With Alist
    .Clear
    
    .AddItem UBound(exeTBL_RESOURCE()) & " ressources trouvés."
    .AddItem ""
    
    For i = 1 To UBound(exeTBL_RESOURCE())
    
        .AddItem "Ressource " & i & " :"
        .AddItem " - Offset : " & exeTBL_RESOURCE(i).RvaDATA
        .AddItem " - Taille : " & exeTBL_RESOURCE(i).Size

    Next i

End With

End Sub

Sub PrintVBAPI(ByRef Alist As ListBox)
'affiche une description des API VB6 utilisé
Dim i, j, ll, fl, e, f
Dim TDs As String, ouR As String
With Alist
    .Clear
    
    ll = UBound(exeIMPORT_DLLNAME())
    j = 1
    If ll > 0 Then
        For i = 1 To ll

            If exeIMPORT_DLLNAME(i).DllName = "MSVBVM60.DLL" Then
            
                f = exeIMPORT_DLLNAME(i).HashPtr
                If i < ll Then
                    e = exeIMPORT_DLLNAME(i + 1).HashPtr - 1
                Else
                    e = UBound(exeIMPORT_APINAME())
                End If
                .AddItem e & " fonctions VB60 trouvés :"
                .AddItem ""
                
                Do While j <= e
                    'compare le nom des API aux fonctions reconnu
                    .AddItem ""
                    If Left$(exeIMPORT_APINAME(j).ApiName, 8) = "!ordinal" Then
                        'via ordinal
                        TDs = VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(j).ApiName, 12)), "", ouR)
                        If TDs = "undef" Then
                            .AddItem "" & exeIMPORT_APINAME(j).ApiName & " (" & ouR & ") - description non dispo"
                        Else
                            .AddItem "" & exeIMPORT_APINAME(j).ApiName & " (" & ouR & ") :"
                            .AddItem "     " & TDs
                        End If
                    Else
                        'via directname
                        TDs = VBfunc_Description(0, exeIMPORT_APINAME(j).ApiName, ouR)
                        If TDs = "undef" Then
                            .AddItem "" & exeIMPORT_APINAME(j).ApiName & " - description non dispo"
                        Else
                            .AddItem "" & exeIMPORT_APINAME(j).ApiName & " :"
                            .AddItem "     " & TDs
                        End If
                    End If
                    j = j + 1
                Loop
                
            End If
            
        Next i
        
    Else
        .AddItem "No Info."
    End If

End With

End Sub

Sub PrintControls(ByRef Alist As ListBox)
'affiche les contrôles VB trouvés dans une liste
Dim i, j

With Alist
    .Clear
    
    If exeISVB Then
        .AddItem "PROJET " & Chr$(34) & exeVB_PROJECTNAME & Chr$(34)
        .AddItem UBound(exeVB_FORMS()) & " forms trouvés, " & UBound(exeVB_CONTROL()) & " objets."
        .AddItem ""
        
        For j = 1 To UBound(exeVB_FORMS())
        
        
            .AddItem "Form " & j & " : " & exeVB_FORMS(j).DefLen & " contrôles trouvés."
            .AddItem ""
            For i = exeVB_FORMS(j).DefPtr To exeVB_FORMS(j).DefPtr + exeVB_FORMS(j).DefLen - 1
                .AddItem " -Contrôle " & i & " :"
                .AddItem "  - nom : " & exeVB_CONTROL(i).sName & VBCTRL_CollectionSpare(exeVB_CONTROL(i).sName)
                .AddItem "  - type : " & exeVB_CONTROL(i).id & "  (" & exeVB_CONTROL(i).sType & ")"
                .AddItem "  - offset " & exeVB_CONTROL(i).offset
                
                If exeVB_CTRL_PRP(i).sCaption <> "" Then
                    .AddItem "    .Caption=" & exeVB_CTRL_PRP(i).sCaption
                End If
                If exeVB_CTRL_PRP(i).pLeft > 0 Then
                    .AddItem "    .Left=" & exeVB_CTRL_PRP(i).pLeft
                End If
                If exeVB_CTRL_PRP(i).pTop > 0 Then
                    .AddItem "    .Top=" & exeVB_CTRL_PRP(i).pTop
                End If
                If exeVB_CTRL_PRP(i).pHeight > 0 Then
                    .AddItem "    .Height=" & exeVB_CTRL_PRP(i).pHeight
                End If
                If exeVB_CTRL_PRP(i).pWidth > 0 Then
                    .AddItem "    .Width=" & exeVB_CTRL_PRP(i).pWidth
                End If
                
                
            Next i
            
            .AddItem ""
            
        Next j
        
    Else
    
        .AddItem "exe non-VB"
        
    End If
    
End With

End Sub

Sub PrintImport(ByRef Alist As ListBox)
'affiche la table d'importation PE de fonctions dans une liste
Dim i, j, f, e, ll

With Alist
    .Clear
    
    ll = UBound(exeIMPORT_DLLNAME())
    j = 1
    If ll > 0 Then
        For i = 1 To ll
            .AddItem " -- Fichier : " & exeIMPORT_DLLNAME(i).DllName
            f = exeIMPORT_DLLNAME(i).HashPtr
            If i < ll Then
                e = exeIMPORT_DLLNAME(i + 1).HashPtr - 1
            Else
                e = UBound(exeIMPORT_APINAME())
            End If
            
            Do While j <= e
                .AddItem Hex(exeIMPORT_APINAME(j).Address) & " : " & exeIMPORT_APINAME(j).ApiName
                j = j + 1
            Loop
            
        Next i
        
    Else
        .AddItem "No Import."
    End If

End With

End Sub

Sub PrintDeclares(ByRef Alist As ListBox)
Dim i, j, es
Dim k, p As Long, op As Boolean
Dim AlreadyPrint() As Long

With Alist
    .Clear
    
    es = UBound(exeVB_API())
    If es = 0 Then
        .AddItem "Pas d'API déclarées"
    Else
        .AddItem es & " DLL(s) déclarée(s) :"
        .AddItem ""
        ReDim AlreadyPrint(es): p = 0
        For i = 1 To es
            j = i
            GoSub AmPrinted
            If op = False Then
                .AddItem ""
                .AddItem "  " & exeVB_API(i).sDll
                
                For j = i To es
                    'affiche les API par ordre de dll parent.
                    GoSub AmPrinted
                    If (exeVB_API(j).sDll = exeVB_API(i).sDll) And (Not op) Then
                        p = p + 1
                        AlreadyPrint(p) = j
                        .AddItem "     " & exeVB_API(j).sName
                    End If
                Next j
                
            End If
        Next i
        
    End If

End With

Exit Sub

AmPrinted:

For k = 1 To es
    If AlreadyPrint(k) = j Then
        op = True
        Return
    End If
Next k
op = False
Return

End Sub


Sub PrintOutPE(ByRef Alist As ListBox)
'affiche les informations PE trouvés dans l'en-tête de l'exe
Dim i

With Alist
    .Clear
    
    .AddItem "===== PE header ====="
    
    .AddItem "Signature : " & exePEHEAD.Signature & "  (" & PE_Check(exePEHEAD.Signature) & ")"
    .AddItem "CpuType : " & exePEHEAD.CpuType & " : pour CPU " & PE_CpuType(exePEHEAD.CpuType)
    .AddItem "Objets : " & exePEHEAD.Objects & " objets (alias sections)"
    
    .AddItem "TimeDate : " & exePEHEAD.TimeDate
    .AddItem "Pointeur symboles : " & exePEHEAD.PointerToSymbolTable
    
    .AddItem "Nombre de symboles: " & exePEHEAD.NumberOfSymbols
    .AddItem "Taille header NT : " & exePEHEAD.NThdrSize
    .AddItem "Flags : " & exePEHEAD.Flags & " : " & PE_Flags(exePEHEAD.Flags)
    
    .AddItem "===== optional header ====="
    
    .AddItem "Taille : " & exeOPHEAD.SizeOfOptionalHeader
    .AddItem "Link Major : " & exeOPHEAD.LinkMajor
    .AddItem "Link Minor : " & exeOPHEAD.LinkMinor
    .AddItem "Reserved 1 : " & exeOPHEAD.Reserved1
    
    .AddItem "Reserved 2 : " & exeOPHEAD.Reserved2
    .AddItem "Reserved 3 : " & exeOPHEAD.Reserved3
    
    .AddItem "Point d'entrée code : " & exeOPHEAD.EntryPointRVA
    .AddItem "Reserved 4 : " & exeOPHEAD.Reserved4
    
    .AddItem "Reserved 5 : " & exeOPHEAD.Reserved5
    .AddItem "ImageBase : " & exeOPHEAD.ImageBase & " (" & Hex(exeOPHEAD.ImageBase) & "h)"
    
    .AddItem "Object Align : " & exeOPHEAD.ObjectAlign
    .AddItem "File Align : " & exeOPHEAD.FileAlign
    
    .AddItem "OS Major : " & exeOPHEAD.OsMajor
    .AddItem "OS Minor : " & exeOPHEAD.OsMinor
    .AddItem "User Major : " & exeOPHEAD.UserMajor
    .AddItem "User Minor : " & exeOPHEAD.UserMinor
    
    .AddItem "SubSys Major : " & exeOPHEAD.SubSysMajor
    .AddItem "SubSys Major : " & exeOPHEAD.SubSysMinor
    .AddItem "Reserved 6 : " & exeOPHEAD.Reserved6
    
    .AddItem "Taille image : " & exeOPHEAD.ImageSize
    .AddItem "Taille header : " & exeOPHEAD.HeaderSize
    
    .AddItem "Checksum : " & Hex(exeOPHEAD.FileCheckSum)
    .AddItem "SubSystem NT : " & exeOPHEAD.SubSystemNT
    .AddItem "Flag de dll : " & exeOPHEAD.DLLflags
    
    .AddItem "Stack Reserve Size : " & exeOPHEAD.StackReserveSize
    .AddItem "Stack Commit Size : " & exeOPHEAD.StackCommitSize
    
    .AddItem "Heap Reserve Size : " & exeOPHEAD.HeapReserveSize
    .AddItem "Heap Commit Size : " & exeOPHEAD.HeapCommitSize
    
    .AddItem "Loader Flags : " & exeOPHEAD.LoaderFlags
    .AddItem "Nb of rva and sizes : " & exeOPHEAD.NumberOfRvaAndSizes

    .AddItem "===== data directories ====="

    .AddItem "Export rva : " & Hex$(exeHEADIR.RvaEXPORT_TABLE)
    .AddItem "Export size : " & exeHEADIR.RvaTOTAL_EXPORT_DATA_SIZE
    
    .AddItem "Import rva : " & Hex$(exeHEADIR.RvaIMPORT_TABLE)
    .AddItem "Import size : " & exeHEADIR.RvaTOTAL_IMPORT_DATA_SIZE
    
    .AddItem "Resource rva : " & exeHEADIR.RvaRESOURCE_TABLE
    .AddItem "Resource size : " & exeHEADIR.RvaTOTAL_RESOURCE_DATA_SIZE
    
    .AddItem exeHEADIR.RvaEXCEPTION_TABLE
    .AddItem exeHEADIR.RvaTOTAL_EXCEPTION_DATA_SIZE
    
    .AddItem exeHEADIR.RvaSECURITY_TABLE
    .AddItem exeHEADIR.RvaTOTAL_SECURITY_DATA_SIZE
    
    .AddItem exeHEADIR.RvaFIXUP_TABLE
    .AddItem exeHEADIR.RvaTOTAL_FIXUP_DATA_SIZE
    
    .AddItem exeHEADIR.RvaDEBUG_TABLE
    .AddItem exeHEADIR.RvaTOTAL_DEBUG_DIRECTORIES
    
    .AddItem exeHEADIR.RvaIMAGE_DESCRIPTION
    .AddItem exeHEADIR.RvaTOTAL_DESCRIPTION_SIZE
    
    .AddItem exeHEADIR.RvaMACHINE_SPECIFIC
    .AddItem exeHEADIR.RvaMACHINE_SPECIFIC_SIZE
    
    .AddItem exeHEADIR.RvaTHREAD_LOCAL_STORAGE
    .AddItem exeHEADIR.RvaTOTAL_TLS_SIZE
    
    .AddItem exeHEADIR.RvaENTRY_IAT
    .AddItem exeHEADIR.RvaTOTAL_ENTRY_IAT_SIZE
    
    .AddItem exeHEADIR.RvaENTRY_D13
    .AddItem exeHEADIR.RvaTOTAL_ENTRY_D13_SIZE
    
    .AddItem exeHEADIR.RvaENTRY_D14
    .AddItem exeHEADIR.RvaTOTAL_ENTRY_D14_SIZE
    
    .AddItem exeHEADIR.RvaENTRY_D15
    .AddItem exeHEADIR.RvaTOTAL_ENTRY_D15_SIZE
    
    .AddItem " "
    .AddItem "===== objects table ====="
    
    For i = 1 To exePEHEAD.Objects
    
        .AddItem PEobj_Name(exeOTABLE(i).oName())
        .AddItem Hex$(exeOTABLE(i).VirtualSize)
        .AddItem Hex$(exeOTABLE(i).VirtualAddress)
    
        .AddItem Hex$(exeOTABLE(i).SizeOfRawData)
        .AddItem Hex$(exeOTABLE(i).PointerToRawData)
    
        .AddItem exeOTABLE(i).PointerToRelocations
        .AddItem exeOTABLE(i).PointerToLinenumbers
    
        .AddItem exeOTABLE(i).NumberOfRelocations
        .AddItem exeOTABLE(i).NumberOfLinenumbers
        .AddItem exeOTABLE(i).oFlags & " : " & PEobj_Flags(exeOTABLE(i).oFlags)
    
    Next i

End With

End Sub

Sub PrintSummary(ByRef AText As TextBox)
'affiche un bilan du fichier, sous forme de texte, dans une textbox
Dim p1 As String, p2 As String, p3 As String, p4 As String
Dim i, j, k, l, es, nb
Dim TmS As String, CatS As String, TDs As String
Dim BGS As String
Dim UniqueID() As Integer

    
    'texte : Import de DLL VB6 et nom de la fonction vb associé
    p1 = "Nombres de DLL importé : " & UBound(exeIMPORT_DLLNAME())
    CatS = "": k = 0
    For i = 1 To UBound(exeIMPORT_APINAME())
        If Left$(exeIMPORT_APINAME(i).ApiName, 8) = "!ordinal" Then
            TDs = VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(i).ApiName, 12)), "", TmS)
        Else
            TDs = VBfunc_Description(0, exeIMPORT_APINAME(i).ApiName, TmS)
        End If
        
        If TDs = "undef" Or Left$(TDs, 4) = "API " Then
        Else
            k = k + 1
            If k Mod 4 = 0 Then
                'retour a la ligne toute les 4 func trouvés
                CatS = CatS & vbCrLf & "   " & TDs & ", "
            Else
                CatS = CatS & TDs & ", "
            End If
        End If
        
        If (exeIMPORT_APINAME(i).ApiName = "DllFunctionCall") Then
            p3 = "    Le programme appelle des dll via la méthode Declare ... Lib ()" & vbCrLf
        End If
        
    Next i
    
        es = UBound(exeVB_API())
        If es > 0 Then
            'il y a des API déclarés via vb !
            For l = 1 To es
                If exeVB_API(l).lFrom > 0 Then
                    BGS = exeVB_API(l).lFrom
                    p3 = p3 & "      Déclaré dans " & BGS & " (info non garantie)" & vbCrLf
                End If
                p3 = p3 & "      - " & UCase$(exeVB_API(l).sDll) & "." & exeVB_API(l).sName & vbCrLf
            Next l
        End If

    p1 = p1 & vbCrLf & "   fonctions VB reconnues : " & vbCrLf & "   " & CatS
    p1 = p1 & vbCrLf & "   (s'il persiste des fonctions non reconnus, cela signifie que le compilateur les a entièrement traduite en ASM)"


    'texte : nom du projet, forms, contrôles et appel de contrôles externe
        'liste des feuilles, classes, modules...
        CatS = ""
        For i = 1 To UBound(exeVB_MODULES())
            Select Case exeVB_MODULES(i).lType
            Case 98305
                CatS = CatS & "    Module : " & exeVB_MODULES(i).sName & vbCrLf
            Case 98435
                CatS = CatS & "    Form   : " & exeVB_MODULES(i).sName & vbCrLf
            Case 1146883
                CatS = CatS & "    Classe : " & exeVB_MODULES(i).sName & vbCrLf
            Case 1941507
                CatS = CatS & "    Contrôle utilisateur : " & exeVB_MODULES(i).sName & vbCrLf
            Case Else
                'Stop
            End Select
        Next i
    
    p2 = "Nom du projet original : " & exeVB_PROJECTNAME & vbCrLf & _
         CatS & _
         "   - forms trouvés : " & UBound(exeVB_FORMS()) & vbCrLf & _
         "   - contrôles trouvés : " & UBound(exeVB_CONTROL())
        
        'crée la liste des type de contrôle existant
        ReDim UniqueID(1 To 1)
        UniqueID(1) = exeVB_CONTROL(1).id
        CatS = ""
        For i = 1 To UBound(exeVB_CONTROL())
            
            For j = 1 To UBound(UniqueID())
                If UniqueID(j) = exeVB_CONTROL(i).id Then GoTo nidx
            Next j
            ReDim Preserve UniqueID(1 To UBound(UniqueID()) + 1)
            UniqueID(UBound(UniqueID())) = exeVB_CONTROL(i).id
            TDs = Get_VBCTRL(UniqueID(j))
            If TDs = "inconnu" Then TDs = exeVB_CONTROL(i).sType
            CatS = CatS & "     " & TDs & vbCrLf
nidx:
        Next i
    p2 = p2 & vbCrLf & "   - type de contrôles :" & vbCrLf & CatS
        
    
    'info sur le code compilé
    p4 = "Compilation :" & vbCrLf & _
         "   Offset de début du code compilé :" & vbCrLf & _
         "    octet " & exeVB_CODEENTRY & vbCrLf & _
         "   Longueur : " & exeVB_CODELEN & " octets." & vbCrLf
    If exeVB_CODEMAIN > 0 Then
        p4 = p4 & "   Le programme commençe via un Sub Main()" & vbCrLf & "    offset " & exeVB_CODEMAIN & vbCrLf
    End If
    p4 = p4 & "  Total " & UBound(exeVB_SUBS) & " subs() utilisateurs trouvés."
    
    'concaténation finale et affichage (note : je vous dit que ce prog' n'est pas optimisé, mais alors pas du tout!!!)
    BGS = "Résultat de l'analyse de " & exeFILENAMElong & vbCrLf & vbCrLf & _
          p2 & vbCrLf & vbCrLf & p1 & vbCrLf & p3 & vbCrLf & p4
    AText.Text = BGS
    
    'ces lignes de code barbares, de concaténation de texte façon porky, c'est vraiment horrible...

End Sub

Sub PrintJolieInterface(ByRef ATree As TreeView)
'sub pour afficher les résultats dans une jolie interface (me fait délirer, ce nom...)
Dim i, j, k, l
Dim cf, cm, cd, cu, cs
Dim op As Boolean
Dim TDs As String

With ATree
    .Nodes.Clear

    .Nodes.Add , 0, "root", exeVB_PROJECTNAME, 3
    .Nodes.Add "root", 4, , "Point d'entrée VB : " & Hex$(exeVB_VBEP - exeOPHEAD.ImageBase) & "h", 16
    .Nodes.Add "root", 4, , "Point d'entrée structure : " & Hex$(exeVB_VBEP - exeOPHEAD.ImageBase + 48) & "h", 16
    
    
    'affiche les feuilles
    Dim TNs As String
    For i = 1 To UBound(exeVB_MODULES())
        Select Case exeVB_MODULES(i).lType
        Case 98305  'module BAS
            cm = cm + 1
            If cm = 1 Then
                .Nodes.Add "root", 4, "modu", "Modules", 1
                .Nodes.Item("modu").ExpandedImage = 2
            End If
            .Nodes.Add "modu", 4, LCase$(exeVB_MODULES(i).sName), exeVB_MODULES(i).sName, 10
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Subs contenu : " & exeVB_MODULES(i).numsub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16

        Case 98435  'form
            cf = cf + 1
            If cf = 1 Then
                .Nodes.Add "root", 4, "form", "Feuilles", 1
                .Nodes.Item("form").ExpandedImage = 2
            End If
            TNs = LCase$(exeVB_MODULES(i).sName)
            .Nodes.Add "form", 4, TNs, exeVB_MODULES(i).sName, 6
            .Nodes.Add TNs, 4, TNs & "-subs", "Subs contenu : " & exeVB_MODULES(i).numsub, 19
            'sub dans la form
                For l = 1 To UBound(exeVB_SUBS())
                    If exeVB_SUBS(l).SubFrom = i Then
                        If exeVB_SUBS(l).SubType = &HFFFFFFFF Then
                            TDs = "Sub " & l & " = " & exeVB_MODULES(i).sName & ".user() : Offset " & Hex$(exeVB_SUBS(l).rvaCode) & "h"
                        Else
                            TDs = "Sub " & l & " = " & exeVB_MODULES(i).sName & ".Obj-" & Hex$(exeVB_SUBS(l).SubType) & "_event() : Offset " & Hex$(exeVB_SUBS(l).rvaCode) & "h"
                        End If
                        .Nodes.Add TNs & "-subs", 4, , TDs, 19
                    End If
                Next l
                
            .Nodes.Add TNs, 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            
            'For j = 1 To UBound(exeVB_FORMS())
            '    If TNs = LCase$(exeVB_CONTROL(exeVB_FORMS(j).DefPtr).sName) Then Exit For
            'Next j
                j = exeVB_MODULES(i).frmidx
            
            .Nodes.Add TNs, 4, , "Offset objets : " & Hex$(exeVB_FORMS(j).rvaPtr) & "h", 16
            'rajoute les objets dans la form
            .Nodes.Add TNs, 4, "form_" & TNs, exeVB_FORMS(j).DefLen & " contrôles trouvés.", 18
            For k = exeVB_FORMS(j).DefPtr To exeVB_FORMS(j).DefPtr + exeVB_FORMS(j).DefLen - 1
                If exeVB_CONTROL(k).id = 276 Then .Nodes.Item(TNs).Image = 7    'icone MDI
                
                .Nodes.Add "form_" & TNs, 4, TNs & "_" & k, exeVB_CONTROL(k).sName & VBCTRL_CollectionSpare(exeVB_CONTROL(k).sName), 5
                .Nodes.Add TNs & "_" & k, 4, , "Longueur " & exeVB_CONTROL(k).LenTr & " octets @ offset " & Hex$(exeVB_CONTROL(k).offset) & "h", 16
                .Nodes.Add TNs & "_" & k, 4, , "Type : " & exeVB_CONTROL(k).id & "  (" & exeVB_CONTROL(k).sType & ")", 18
                .Nodes.Add TNs & "_" & k, 4, , "systID : " & Hex$(exeVB_CONTROL(k).frmID) & "  (" & Hex$(exeVB_CONTROL(k).frmID - 709) & ")", 18
                
                If exeVB_CTRL_PRP(k).sCaption <> "" Then
                    .Nodes.Add TNs & "_" & k, 4, , "Caption=" & exeVB_CTRL_PRP(k).sCaption, 18
                End If
                If exeVB_CTRL_PRP(k).pLeft > 0 Then
                    .Nodes.Add TNs & "_" & k, 4, , "Left=" & exeVB_CTRL_PRP(k).pLeft, 18
                End If
                If exeVB_CTRL_PRP(k).pTop > 0 Then
                    .Nodes.Add TNs & "_" & k, 4, , "Top=" & exeVB_CTRL_PRP(k).pTop, 18
                End If
                If exeVB_CTRL_PRP(k).pHeight > 0 Then
                    .Nodes.Add TNs & "_" & k, 4, , "Height=" & exeVB_CTRL_PRP(k).pHeight, 18
                End If
                If exeVB_CTRL_PRP(k).pWidth > 0 Then
                    .Nodes.Add TNs & "_" & k, 4, , "Width=" & exeVB_CTRL_PRP(k).pWidth, 18
                End If
                
                'subs
                For l = 1 To UBound(exeVB_SUBS())
                    If exeVB_SUBS(l).SubFrom = i And exeVB_SUBS(l).ObjFrom = k Then
                        TDs = "Sub " & l & " = " & exeVB_CONTROL(k).sName & "_() : Offset " & Hex$(exeVB_SUBS(l).rvaCode) & "h"
                        .Nodes.Add TNs & "_" & k, 4, , TDs, 19
                    End If
                Next l
                
            Next k
            
        Case 1146883  'classe
            cs = cs + 1
            If cs = 1 Then
                .Nodes.Add "root", 4, "clas", "Modules de classe", 1
                .Nodes.Item("clas").ExpandedImage = 2
            End If
            .Nodes.Add "clas", 4, LCase$(exeVB_MODULES(i).sName), exeVB_MODULES(i).sName, 9
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, LCase$(exeVB_MODULES(i).sName) & "-subs", "Subs contenu : " & exeVB_MODULES(i).numsub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            
            For l = 1 To UBound(exeVB_SUBS())
                If exeVB_SUBS(l).SubFrom = i Then
                    TDs = "Sub " & l & " = " & exeVB_MODULES(i).sName & ".user() : Offset " & Hex$(exeVB_SUBS(l).rvaCode) & "h"
                    .Nodes.Add LCase$(exeVB_MODULES(i).sName) & "-subs", 4, , TDs, 19
                End If
            Next l
            
        Case 1941507  'contrôles utilisateur
            cu = cu + 1
            If cu = 1 Then
                .Nodes.Add "root", 4, "uctl", "Contrôles utilisateur", 1
                .Nodes.Item("uctl").ExpandedImage = 2
            End If
            .Nodes.Add "uctl", 4, LCase$(exeVB_MODULES(i).sName), exeVB_MODULES(i).sName, 11
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, LCase$(exeVB_MODULES(i).sName) & "-subs", "Subs contenu : " & exeVB_MODULES(i).numsub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            
            For l = 1 To UBound(exeVB_SUBS())
                If exeVB_SUBS(l).SubFrom = i Then
                    TDs = "Sub " & l & " = " & exeVB_MODULES(i).sName & ".user() : Offset " & Hex$(exeVB_SUBS(l).rvaCode) & "h"
                    .Nodes.Add LCase$(exeVB_MODULES(i).sName) & "-subs", 4, , TDs, 19
                End If
            Next l
            
            
        Case Else
            'Stop
            '.Nodes.Add "root", 4, "prop", "Pages de propriétés", 1, 2
        End Select
    Next i

    'affiche plus d'info :
    .Nodes.Add "root", 4, "info", "Informations supplémentaires", 1
    .Nodes.Item("info").ExpandedImage = 2
    
        'DLL importés
        Dim ouR As String
        l = UBound(exeIMPORT_DLLNAME())
        j = 1
        If l > 0 Then
            .Nodes.Add "info", 4, "import", "DLL importées", 1
            .Nodes.Item("import").ExpandedImage = 2
            For i = 1 To l
                .Nodes.Add "import", 4, LCase$(exeIMPORT_DLLNAME(i).DllName), exeIMPORT_DLLNAME(i).DllName, 15
                If i < l Then
                    k = exeIMPORT_DLLNAME(i + 1).HashPtr - 1
                Else
                    k = UBound(exeIMPORT_APINAME())
                End If
                
                Do While j <= k
                
                    .Nodes.Add LCase$(exeIMPORT_DLLNAME(i).DllName), 4, exeIMPORT_APINAME(j).ApiName, exeIMPORT_APINAME(j).ApiName, 14
                    .Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Offset " & Hex(exeIMPORT_APINAME(j).Address) & "h", 16
                    
                    If Left$(LCase$(exeIMPORT_DLLNAME(i).DllName), 8) = "msvbvm60" Then
                    If Left$(exeIMPORT_APINAME(j).ApiName, 8) = "!ordinal" Then
                        'via ordinal
                        TDs = VBfunc_Description(Val(Mid$(exeIMPORT_APINAME(j).ApiName, 12)), "", ouR)
                        If TDs = "undef" Then
                            .Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Nom complet : " & ouR, 18
                        Else
                            .Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , "Nom complet : " & ouR, 18
                            .Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                    Else
                        'via directname
                        TDs = VBfunc_Description(0, exeIMPORT_APINAME(j).ApiName, ouR)
                        If TDs = "undef" Then
                        Else
                            .Nodes.Add exeIMPORT_APINAME(j).ApiName, 4, , TDs, 19
                        End If
                    End If
                    End If
                    j = j + 1
                Loop
            Next i
            
        Else
            
        End If

        
        'API déclarés
        Dim AlreadyPrint() As Long
        cd = UBound(exeVB_API())
        If cd = 0 Then
        Else
            .Nodes.Add "info", 4, "declares", "API déclarées", 1
            .Nodes.Item("declares").ExpandedImage = 2
            
            ReDim AlreadyPrint(cd): k = 0
            For i = 1 To cd
                j = i
                GoSub AmPrinted
                If op = False Then
                    .Nodes.Add "declares", 4, LCase$(exeVB_API(i).sDll), exeVB_API(i).sDll, 15
                    
                    For j = i To cd
                        'affiche les API par ordre de dll parent.
                        GoSub AmPrinted
                        If (exeVB_API(j).sDll = exeVB_API(i).sDll) And (Not op) Then
                            k = k + 1
                            AlreadyPrint(k) = j
                            .Nodes.Add LCase$(exeVB_API(i).sDll), 4, , exeVB_API(j).sName, 14
                        End If
                    Next j
                    
                End If
            Next i
        End If

        'code compilé
        .Nodes.Add "info", 4, "codebase", "Code compilé", 1
        .Nodes.Item("codebase").ExpandedImage = 2
    
        .Nodes.Add "codebase", 4, "codeep", "Point d'entrée : " & Hex$(exeVB_CODEENTRY) & "h", 16
        .Nodes.Add "codeep", 4, , "Longueur : " & exeVB_CODELEN & " octets   (fin du code à " & Hex$(exeVB_CODEENTRY + exeVB_CODELEN) & "h)", 18
        'affiche les infos s'il y a un Sub Main()
        'If exeVB_CODEMAIN > 0 Then
        '    .Nodes.Add "codeep", 4, "sub_main", "Sub Main", 19
        '    .Nodes.Add "sub_main", 4, , "Offset " & Hex$(exeVB_CODEMAIN) & "h", 16
        'End If
        'affiche la liste des subs trouvés
        k = UBound(exeVB_SUBS)
        For i = 1 To k
            
            'Select Case exeVB_SUBS(i).SubType
            'Case 1
            '    TDs = "Sub "
            'Case 2
            '    TDs = "Sub "
            'Case Else
            '    TDs = "Sub "
            'End Select
            TDs = "Sub "
            If exeVB_SUBS(i).SubFrom > 0 Then
                If exeVB_SUBS(i).SubType = &HFFFFFFFF Then
                    TNs = " (" & exeVB_MODULES(exeVB_SUBS(i).SubFrom).sName & ".user() :"
                Else
                    TNs = " (" & exeVB_MODULES(exeVB_SUBS(i).SubFrom).sName & ".Obj-" & Hex$(exeVB_SUBS(i).SubType) & "_event() :"
                End If
            ElseIf exeVB_SUBS(i).SubFrom = -1 Then
                TNs = " (module).user() :"
            Else
                TNs = " (Sub Main) :"
            End If
            .Nodes.Add "codeep", 4, , TDs & i & TNs & " Offset " & Hex$(exeVB_SUBS(i).rvaCode) & "h", 19
            
        Next i

    
    
    'formate l'affichage
    .Nodes.Item("root").Expanded = True

End With

Exit Sub

'misub utile
AmPrinted:
    For l = 1 To cd
        If AlreadyPrint(l) = j Then
            op = True
            Return
        End If
    Next l
    op = False
Return


End Sub

Private Function PE_Check(NID As Long) As String
    If NID = 17744 Then
        PE_Check = "valide"
    Else
        PE_Check = "invalide"
    End If
End Function

Private Function PE_CpuType(ByVal CPID As Long) As String

Select Case CPID
    Case &H14C
        PE_CpuType = "80386"

    Case &H14D
        PE_CpuType = "80486"
        
    Case &H14E
        PE_CpuType = "80586 (pentium)"
    
    Case &H162
        PE_CpuType = "MIPS Mark I (R2000, R3000)"

    Case &H163
        PE_CpuType = "MIPS Mark II (R6000)"

    Case &H166
        PE_CpuType = "MIPS Mark III (R4000)"
    
    Case &H168
        PE_CpuType = "MIPS Mark IV (R10000)"
    
    Case &H184
        PE_CpuType = "DEC Alpha AXP"
    
    Case &H1F0
        PE_CpuType = "IBM Power PC"
    
    Case Else
        PE_CpuType = "inconnu"
        
End Select
End Function

Private Function PE_Flags(ByVal Flag As Long) As String
' interprète le flag du PE HEADER
    If CBool(Flag And 1) = True Then
        PE_Flags = "no relocation"
    End If
    If CBool(Flag And 2) = True Then
        PE_Flags = PE_Flags & ", executable"
    End If
    If CBool(Flag And 4) = True Then
        PE_Flags = PE_Flags & ", stripped line" 'noms remplacé par ordinaux
    End If
    If CBool(Flag And 8) = True Then
        PE_Flags = PE_Flags & ", no local symbol info"
    End If
    If CBool(Flag And 16) = True Then
        PE_Flags = PE_Flags & ", OS aggressive trim"
    End If
    If CBool(Flag And 128) = True Then
        PE_Flags = PE_Flags & ", bytes swapped"
    End If
    If CBool(Flag And 256) = True Then
        PE_Flags = PE_Flags & ", for 32bit machine"
    End If
    If CBool(Flag And 512) = True Then
        PE_Flags = PE_Flags & ", no debug info"
    End If
    If CBool(Flag And 1024) = True Then
        PE_Flags = PE_Flags & ", removable : run from swap"
    End If
    If CBool(Flag And 2048) = True Then
        PE_Flags = PE_Flags & ", net : run from swap"
    End If
    If CBool(Flag And 4096) = True Then
        PE_Flags = PE_Flags & ", system file"
    End If
    If CBool(Flag And 8192) = True Then
        PE_Flags = PE_Flags & ", dll"
    End If
    If CBool(Flag And 16384) = True Then
        PE_Flags = PE_Flags & ", not for smp"
    End If

End Function

Private Function PEobj_Name(ByRef inBytes() As Byte) As String
'renvoi le nom de la section PE (toujours sur 8 octet)
    PEobj_Name = Space$(8)
    Mid$(PEobj_Name, 1, 1) = Chr$(inBytes(1))
    Mid$(PEobj_Name, 2, 1) = Chr$(inBytes(2))
    Mid$(PEobj_Name, 3, 1) = Chr$(inBytes(3))
    Mid$(PEobj_Name, 4, 1) = Chr$(inBytes(4))
    Mid$(PEobj_Name, 5, 1) = Chr$(inBytes(5))
    Mid$(PEobj_Name, 6, 1) = Chr$(inBytes(6))
    Mid$(PEobj_Name, 7, 1) = Chr$(inBytes(7))
    Mid$(PEobj_Name, 8, 1) = Chr$(inBytes(8))
    
End Function

Private Function PEobj_Flags(ByVal Flag As Long) As String
'interprétation des flags des sections PE! (incomplet : fait ch...)
If CBool(Flag And 32) = True Then
        PEobj_Flags = "executable"
    End If
    If CBool(Flag And 64) = True Then
        PEobj_Flags = PEobj_Flags & ", data to initialize"
    End If
    If CBool(Flag And 128) = True Then
        PEobj_Flags = PEobj_Flags & ", data uninitialized (bss)"
    End If
    If CBool(Flag And 512) = True Then
        PEobj_Flags = PEobj_Flags & ", link info"
    End If
    If CBool(Flag And 2048) = True Then
        PEobj_Flags = PEobj_Flags & ", link remove"
    End If
    If CBool(Flag And 4096) = True Then
        PEobj_Flags = PEobj_Flags & ", link common data"
    End If
    If CBool(Flag And 32768) = True Then
        PEobj_Flags = PEobj_Flags & ", mem far data"
    End If
    If CBool(Flag And 131072) = True Then
        PEobj_Flags = PEobj_Flags & ", mem purgeable"
    End If
    If CBool(Flag And 262144) = True Then
        PEobj_Flags = PEobj_Flags & ", mem locked"
    End If
    If CBool(Flag And 524288) = True Then
        PEobj_Flags = PEobj_Flags & ", mem preload"
    End If
    '...
    If CBool(Flag And 268435456) = True Then
        PEobj_Flags = PEobj_Flags & ", mem share"
    End If
    If CBool(Flag And 536870912) = True Then
        PEobj_Flags = PEobj_Flags & ", mem execute"
    End If
    If CBool(Flag And 1073741824) = True Then
        PEobj_Flags = PEobj_Flags & ", mem read"
    End If
    If CBool(Flag And &H80000000) = True Then
        PEobj_Flags = PEobj_Flags & ", mem write"
    End If

End Function

Sub PrintExe(Fichier As String, ByVal OffsetMod16 As Long, Ligne As Long, oPCB As PictureBox)
'affichage dans une picturebox du fichier en hexa
Dim i, j, k, l, m, n
Dim bArray() As Byte
Dim cArray(1 To 15) As Long
Dim Lp As Integer
Dim Tstr As String
'i = incrémenteur de ligne
'j = valeur offset dans le fichier de la ligne en cours (correspondance de i)
'k = index du tableau tampon contenant les octets du fichier
'l = place mémoire pour placer la valeur hexa d'un octet dans une string
'm = place mémoire pour placer la valeur ANSI d'un octet dans la string
'n = valeur offset dans le fichier de l'octet lu

'marqueur coloré :
'en-tête exe & PE = gris très foncé
cArray(1) = &H202020
'table API import, noms des API = bleu
cArray(2) = vbBlue
'table API déclaré VB = bleu clair
cArray(3) = RGB(128, 128, 255)
'structure de définition VB = marron
cArray(4) = RGB(64, 64, 32)
'structure VB form = rouge
cArray(5) = vbRed
'structure VB modu = orange
cArray(6) = RGB(192, 128, 128)
'structure VB clas = magenta
cArray(7) = vbMagenta
'structure VB ctrl = rose
cArray(8) = RGB(192, 112, 112)
'variables/données VB = vert
cArray(9) = RGB(32, 192, 32)
'table de sub VB = vert clair
cArray(10) = RGB(0, 192, 0)
'code compilé VB = vert foncé
cArray(11) = RGB(0, 128, 0)
'ressources PE = gris fonçé
cArray(12) = &H505050


With oPCB

    ReDim bArray(1 To (16 * Ligne) + 1)
    Lp = FreeFile
    
    Open Fichier For Binary Access Read As #Lp
        Seek #Lp, OffsetMod16 * 16 + 1
        Get #Lp, , bArray()
    Close #Lp
    
    .AutoRedraw = True
    .Cls
    .FontBold = True
    .FontSize = 8
    .FontName = "Courier New"
    .CurrentX = 15
    .CurrentY = 0
    .ForeColor = &H888888
    oPCB.Print "-Offset- 00 01 02 03 04 05 06 07-08 09 0A 0B 0C 0D 0E 0F ------ANSI------"
    .ForeColor = vbBlack
    
    j = OffsetMod16 * 16
    
    k = 0
    For i = 1 To Ligne
        .CurrentY = i * 190
        .CurrentX = 10
        l = 1: m = 1
        Tstr = Right$("0000000" & Hex$(j), 8) & Space$(70)
        Do
            k = k + 1
            Mid$(Tstr, l * 3 + 7, 2) = Right$("0" & Hex$(bArray(k)), 2)
            Mid$(Tstr, m + 57, 1) = IIf(bArray(k) > 30, Chr$(bArray(k)), ".")
            l = l + 1: m = m + 1
        Loop Until (k Mod 16) = 0
        oPCB.Print Tstr 'groumf
        
        j = OffsetMod16 * 16 + (i * 16)
    Next i

    .AutoRedraw = False

End With
End Sub

Sub PrintExeCompare(Fichier1 As String, Fichier2 As String, ByVal OffsetMod16 As Long, Ligne As Long, oPCB1 As PictureBox, oPCB2 As PictureBox)
'affichage dans deux picturebox de deux fichiers en hexa, avec différence en surbrillance
'EN DEV
Dim i, j, k, l, m
Dim bArray1() As Byte, bArray2() As Byte
Dim Lp As Integer
Dim Tstr1 As String, Tstr2 As String

    ReDim bArray1(1 To (16 * Ligne) + 1)
    ReDim bArray2(1 To (16 * Ligne) + 1)
    Lp = FreeFile
    
    Open Fichier1 For Binary Access Read As #Lp
        Seek #Lp, OffsetMod16 * 16 + 1
        Get #Lp, , bArray1()
    Close #Lp
    Open Fichier2 For Binary Access Read As #Lp
        Seek #Lp, OffsetMod16 * 16 + 1
        Get #Lp, , bArray2()
    Close #Lp

    oPCB1.AutoRedraw = True
    oPCB1.Cls
    oPCB1.FontSize = 8
    oPCB1.FontName = "Courier"
    oPCB1.ForeColor = vbBlack
    oPCB2.AutoRedraw = True
    oPCB2.Cls
    oPCB2.FontSize = 8
    oPCB2.FontName = "Courier"
    oPCB2.ForeColor = vbBlack

    j = OffsetMod16 * 16
    
    k = 0
    For i = 1 To Ligne
        oPCB1.CurrentY = i * 150
        oPCB1.CurrentX = 10
        oPCB2.CurrentY = i * 150
        oPCB2.CurrentX = 10
        l = 1: m = 1
        Tstr1 = Right$("0000000" & Hex$(j), 8) & Space$(70)
        Tstr2 = Right$("0000000" & Hex$(j), 8) & Space$(70)
        Do
            k = k + 1
            Mid$(Tstr1, l * 3 + 7, 2) = Right$("0" & Hex$(bArray1(k)), 2)
            Mid$(Tstr1, m + 57, 1) = IIf(bArray1(k) > 30, Chr$(bArray1(k)), ".")
            Mid$(Tstr2, l * 3 + 7, 2) = Right$("0" & Hex$(bArray2(k)), 2)
            Mid$(Tstr2, m + 57, 1) = IIf(bArray2(k) > 30, Chr$(bArray2(k)), ".")
            l = l + 1: m = m + 1
        Loop Until (k Mod 16) = 0
        oPCB1.Print Tstr1
        oPCB2.Print Tstr2
        
        j = OffsetMod16 * 16 + (i * 16)
    Next i

    oPCB1.AutoRedraw = False
    oPCB2.AutoRedraw = False

End Sub



Public Function VBCTRL_CollectionSpare(inStrN As String) As String
'petite routine bidon pour l'affichage des collections d'objets vb
Dim i, j, k

    i = InStr(2, inStrN, "(", vbBinaryCompare)
    If i > 0 Then
        j = InStr(i, inStrN, ")", vbBinaryCompare)
        k = InStr(i, inStrN, "-", vbBinaryCompare)
        VBCTRL_CollectionSpare = "  Collection " & Mid$(inStrN, i + 1, k - i - 1) & _
                                 "  Index " & Mid$(inStrN, k + 1, j - k - 1)
    Else
        VBCTRL_CollectionSpare = vbNullString
    End If
    
End Function

Public Sub VBCTRL_VbpRebuild(inTblF() As CONTROL_FORM, ByVal DestDir As String)
'reconstruit un fichier vbp permettant d'ouvrir toutes les forms reconstruite ...
'(oui, ce n'est pas un vrai vbp... pour l'instant)

'NOTA : CE SUB N'EST PAS NECESSAIRE - je l'ai codé pour le fun :)
'   utilisez VBReFormer (codé par Warning) pour reconstruire intégralement les feuilles
Dim i
Dim Lp As Integer
Dim TNs As String

SubDeb:
If Dir$(DestDir, vbDirectory) <> "" Then
    If Right$(DestDir, 1) <> "\" Then DestDir = DestDir & "\"
    Lp = FreeFile
    Open DestDir & exeVB_PROJECTNAME & ".vbp" For Output As #Lp

        Print #Lp, "Type=Exe"

        For i = 1 To UBound(inTblF())
        
            Print #Lp, "Form=" & exeVB_CONTROL(inTblF(i).DefPtr).sName & ".frm"

        Next i
        
        TNs = exeVB_CONTROL(inTblF(1).DefPtr).sName
        Print #Lp, "IconForm=" & Chr$(34) & TNs & Chr$(34)
        Print #Lp, "Startup=" & Chr$(34) & TNs & Chr$(34)
        Print #Lp, "Name=" & Chr$(34) & exeVB_PROJECTNAME & Chr$(34)

        
    Close #Lp
    
Else
    
    MkDir DestDir
    GoTo SubDeb
    
End If

End Sub

Public Sub VBCTRL_FrmRebuild(inTblF As CONTROL_FORM, ByVal DestDir As String)
'reconstruit une feuille frm d'un a partir de ses paramètres retrouvés
'(en standby : il faut récupérer les paramètres correctement d'abord!)

'NOTA : CE SUB N'EST PAS NECESSAIRE - je l'ai codé pour le fun :)
'   utilisez VBReFormer (codé par Warning) pour reconstruire intégralement les feuilles

Dim i, j, p, k, l
Dim Lp As Integer
Dim Tstr As String, Dstr As String
Dim UsePos As Long
Dim AutoLeft As Long, AutoTop As Long

SubDeb:
If Dir$(DestDir, vbDirectory) <> "" Then

    p = inTblF.DefPtr

    If Right$(DestDir, 1) <> "\" Then DestDir = DestDir & "\"
    Lp = FreeFile
    Open DestDir & exeVB_CONTROL(p).sName & ".frm" For Output As #Lp
    
        
    
        Print #Lp, "VERSION 5.00"
        
        Print #Lp, "BEGIN " & exeVB_CONTROL(p).sType & " " & exeVB_CONTROL(p).sName
        Print #Lp, "   Caption         =   " & Chr$(34) & exeVB_CTRL_PRP(p).sCaption & Chr$(34)
        Print #Lp, "   Left            =   " & exeVB_CTRL_PRP(p).pLeft
        Print #Lp, "   Top             =   " & exeVB_CTRL_PRP(p).pTop
        Print #Lp, "   Width           =   " & exeVB_CTRL_PRP(p).pWidth
        Print #Lp, "   Height          =   " & exeVB_CTRL_PRP(p).pHeight
    
        For i = p + 1 To p + inTblF.DefLen - 1
        If exeVB_CONTROL(i).sType <> "inconnu" Then
        
            'DEBUT
            Print #Lp, "   BEGIN " & exeVB_CONTROL(i).sType & " " & VBCTRL_CollectionPrint(exeVB_CONTROL(i).sName, Dstr)
            
            'PARAMETRES COMMUNS
                If exeVB_CTRL_PRP(i).sCaption <> "" Then
                    'caption?
                    Print #Lp, "      Caption         =   " & Chr$(34) & exeVB_CTRL_PRP(i).sCaption & Chr$(34)
                End If
                If Dstr <> "" Then
                    'index?
                    Print #Lp, "      " & Dstr
                End If
            
            UsePos = (exeVB_CONTROL(i).id And 255) <> 19 'suis-je un menu ?
            
            'POSITION
            If UsePos Then
                If exeVB_CTRL_PRP(i).pLeft > 0 Then
                    'left
                    Print #Lp, "      Left            =   " & exeVB_CTRL_PRP(i).pLeft
                Else
                    AutoLeft = AutoLeft + 100
                    Print #Lp, "      Left            =   " & AutoLeft
                End If
                If exeVB_CTRL_PRP(i).pTop > 0 Then
                    'top
                    Print #Lp, "      Top             =   " & exeVB_CTRL_PRP(i).pTop
                Else
                    AutoTop = AutoTop + 100
                    Print #Lp, "      Top             =   " & AutoTop
                End If
                If exeVB_CTRL_PRP(i).pWidth > 0 Then
                    'width
                    Print #Lp, "      Width           =   " & exeVB_CTRL_PRP(i).pWidth
                End If
                If exeVB_CTRL_PRP(i).pHeight > 0 Then
                    'height
                    Print #Lp, "      Height          =   " & exeVB_CTRL_PRP(i).pHeight
                End If
            Else
            'MENU
                k = i
                
                Do
                    'scanne les sous-menus :
                    k = k + 1
                    If k > (p + inTblF.DefLen - 1) Then Exit Do
                    If (exeVB_CTRL_PRP(k).pRank = -249) Or ((exeVB_CONTROL(k).id And 255) <> 19) Then Exit Do
                    
                    Print #Lp, "      BEGIN " & exeVB_CONTROL(k).sType & " " & VBCTRL_CollectionPrint(exeVB_CONTROL(k).sName, Dstr)
                    Print #Lp, "         Caption         =   " & Chr$(34) & exeVB_CTRL_PRP(k).sCaption & Chr$(34)
                    If Dstr <> "" Then
                        'subindex
                        Print #Lp, "         " & Dstr
                    End If
                    Print #Lp, "      END"
                Loop
                i = k - 1

            End If
            
            'FIN
            
            Print #Lp, "   END"
            
        End If
        Next i
        
        Print #Lp, "END"
        Print #Lp, "Attribute VB_Name = " & Chr$(34) & exeVB_CONTROL(p).sName & Chr$(34)
        Print #Lp, "Attribute VB_GlobalNameSpace = False"
        Print #Lp, "Attribute VB_Creatable = False"
        Print #Lp, "Attribute VB_PredeclaredId = True"
        Print #Lp, "Attribute VB_Exposed = False"
        
    Close #Lp

Else
    
    MkDir DestDir
    GoTo SubDeb
    
End If
End Sub

Private Function VBCTRL_CollectionPrint(inName As String, outIndex As String) As String
Dim i, j, k

    i = InStr(2, inName, "(", vbBinaryCompare)
    If i > 0 Then
        j = InStr(i, inName, ")", vbBinaryCompare)
        k = InStr(i, inName, "-", vbBinaryCompare)
        VBCTRL_CollectionPrint = Left$(inName, i - 2)
        outIndex = "Index           =   " & Mid$(inName, k + 1, j - k - 1)
    Else
        VBCTRL_CollectionPrint = inName
        outIndex = vbNullString
    End If

End Function

Function VBCTRL_AsmProp(ByVal AsmValue As Long, objType As Variant)

End Function

Sub VBCTRL_AsmProperty()
'renvoi la propriété appelé en fonction de la valeur...
ReDim exeVB_Prop(1 To &H2C4)

'reste a établir la liste exhaustive pour les 26 contrôle/objets de VB ...
'a noté que &HAC peut être textbox.fontname ou filelistbox.path, et d'autres encore!
'il faut donc faire un tableau à double entrée...

exeVB_Prop(&H54) = "Caption"
'call dword ptr [eax+0000009C] = Visible
exeVB_Prop(&H9C) = "Visible"
'call dword ptr [eax+0000016C] = TextBox.Alignment
'call dword ptr [eax+000000EC] = Label.Alignment
exeVB_Prop(&HEC) = "Alignment (label)"
'call dword ptr [eax+000001CC] = TextBox.Appearance
'call dword ptr [eax+00000184] = Label.Appearance
'call dword ptr [eax+0000014C] = AutoRedraw
exeVB_Prop(&H14C) = "AutoRedraw"
'call dword ptr [eax+5C] = TextBox.BackColor
'call dword ptr [eax+5C] = Timer.Enabled
exeVB_Prop(&H5C) = "Enabled (timer)"
'call dword ptr [eax+000000E4] = Borderstyle
'call dword ptr [eax+00000234] = CausesValidation
'call dword ptr [eax+000001D0] = Container
'call dword ptr [eax+000001A4] = DataChanged
'call dword ptr [eax+0000019C] = DataField
'call dword ptr [eax+00000240] = DataFormat
'call dword ptr [eax+0000023C] = DataMember
'call dword ptr [eax+0000013C] = DragIcon
'call dword ptr [eax+00000134] = DragMode
'call dword ptr [eax+0000008C] = Enabled
'call dword ptr [eax+000001BC] = Font
'call dword ptr [eax+000000BC] = FontBold
'call dword ptr [eax+000000C4] = FontItalic
'call dword ptr [eax+000000AC] = FontName
exeVB_Prop(&HAC) = "FontName (textbox)"
'call dword ptr [eax+000000B4] = FontSize
'call dword ptr [eax+000000CC] = FontStrikethru
'call dword ptr [eax+000000D4] = FontUnderline
'call dword ptr [eax+64] = ForeColor
exeVB_Prop(&H64) = "ForeColor"
'call dword ptr [eax+00000084] = Height
exeVB_Prop(&H84) = "Height"
'call dword ptr [eax+0000017C] = HelpContextID
'call dword ptr [eax+00000160] = HideSelection
'call dword ptr [eax+00000180] = hWnd
'call dword ptr [eax+50] = Index
'call dword ptr [eax+6C] = Left
exeVB_Prop(&H6C) = "Left"
'call dword ptr [eax+000000F4] = LinkItem
'call dword ptr [eax+000000FC] = LinkMode
'call dword ptr [eax+00000144] = LinkTimeout
'call dword ptr [eax+000000EC] = LinkTopic
'call dword ptr [eax+000001B4] = TextBox.Locked
exeVB_Prop(&H1B4) = "Locked"
'call dword ptr [eax+00000174] = MaxLength
'call dword ptr [eax+000001AC] = MouseIcon
'call dword ptr [eax+0000009C] = MousePointer
'call dword ptr [eax+00000100] = MultiLine
'call dword ptr [eax+48] = Name
'call dword ptr [eax+000001EC] = OLEDragMode
'call dword ptr [eax+000001F4] = OLEDropMode
'call dword ptr [eax+00000128] = Parent
'call dword ptr [eax+0000015C] = PasswordChar
'call dword ptr [eax+000001DC] = RightToLeft
'call dword ptr [eax+00000108] = ScrollBars
'call dword ptr [eax+0000011C] = SelLength
'call dword ptr [eax+00000114] = SelStart
'call dword ptr [eax+000000DC] = TabIndex
'call dword ptr [eax+0000014C] = TabStop
'call dword ptr [eax+00000154] = Tag
'call dword ptr [eax+000000A4] = Text
exeVB_Prop(&HA4) = "Text"
'call dword ptr [eax+000001E4] = ToolTipText
'call dword ptr [eax+74] = Top
exeVB_Prop(&H74) = "Top"
'call dword ptr [eax+00000094] = TextBox.Visible
exeVB_Prop(&H94) = "Visible (textbox)"
'call dword ptr [eax+000001C4] = WhatsThisHelpID
'call dword ptr [eax+7C] = Width
exeVB_Prop(&H7C) = "Width"

'call dword ptr [eax+000001A4] = FileListBox.Appearance
'call dword ptr [eax+000000D4] = Archive
'call dword ptr [eax+000000BC] = FileName
exeVB_Prop(&HBC) = "FileName (filelistbox)"
'call dword ptr [eax+000000DC] = Hidden
'call dword ptr [eax+000000AC] = Path
exeVB_Prop(&HAC) = "Path (filelistbox)"
'call dword ptr [eax+000000B4] = Pattern
exeVB_Prop(&HB4) = "Pattern (filelistbox)"
'call dword ptr [eax+000000F8] = List()
exeVB_Prop(&HF8) = "List ()"
'call dword ptr [eax+000000E8] = ListCount
exeVB_Prop(&HE8) = "ListCount"
'call dword ptr [eax+000000F0] = ListIndex
exeVB_Prop(&HF0) = "ListIndex"
'call dword ptr [eax+00000168] = MultiSelect
'call dword ptr [eax+000000C4] = FileListBox.Normal
'call dword ptr [eax+00000170] = Selected
exeVB_Prop(&H170) = "Selected"
'call dword ptr [eax+000000E4] = FileListBox.System
'call dword ptr [eax+000000CC] = FileListBox.ReadOnly

'call dword ptr [eax+000000C0] = DirListBox.Path"
exeVB_Prop(&HC0) = "List (dirlistbox)"

exeVB_Prop(&H1E8) = "Clear (listbox)"


'call dword ptr [eax+5C] = Timer.Enabled
'call dword ptr [eax+64] = Timer.Interval
exeVB_Prop(&H64) = "Interval (timer)"

'call dword ptr [eax+000000E0] = CheckBox.Value
exeVB_Prop(&HE0) = "Value (checkbox)"
'call dword ptr [eax+00000190] = Style

exeVB_Prop(&H8C) = "Enabled (commandbutton)"

'call dword ptr [eax+58] = Form.hWnd
exeVB_Prop(&H58) = "hWnd (form)"
'call dword ptr [eax+000002C4] = Form.Cls
exeVB_Prop(&H2C4) = "Cls (form)"
'call dword ptr [eax+00000254] = Form.Appearance
exeVB_Prop(&H254) = "Appearance (form)"

exeVB_Prop(&H278) = "Cls (picturebox)"

exeVB_Prop(&H2B0) = "Show (form)"
exeVB_Prop(&H2B4) = "Hide (form)"


End Sub

Function ByteToStr(inB() As Byte) As String
'converti des bytes en hex dump
Dim i, j, n, l, r
n = UBound(inB())  'longueur
l = n \ 16         'nombre de lignes de 16 octets
r = n Mod 16       'longueur derniere ligne
ByteToStr = Space$(l * 16 * 3 + r * 3 + l * 2)
    
    j = 1
    i = 1
    Do While j <= n
        Mid$(ByteToStr, i, 2) = Right$("0" & Hex$(inB(j)), 2)
        i = i + 3
        If j Mod 16 = 0 Then 'retour à la ligne
            Mid$(ByteToStr, i, 2) = vbCrLf
            i = i + 2
        End If
        j = j + 1
    Loop

End Function


Public Sub Util_SnifStart(frmJI As TreeView, ByVal epva As Long)
'scanner recursif de RVA
Dim fp As Integer
Dim lDump As Long, min, max
Dim followup
Dim lkey As String
min = exeOPHEAD.ImageBase + 1024
max = exeOPHEAD.ImageBase + exeFILENAMEsize

    fp = FreeFile
    Open exeFILENAMElong For Binary Access Read As #fp
    followup = 1
    lkey = "a" & CStr(epva + exeOPHEAD.ImageBase)
    frmJI.Nodes.Add , , lkey, "Offset " & Hex$(epva), 1
    Call Util_SnifNext(min, max, fp, followup, lkey, epva + exeOPHEAD.ImageBase, frmJI)
    Close #fp

End Sub

Private Sub Util_SnifNext(ByVal min As Long, ByVal max As Long, ByVal fp As Integer, ByVal fu As Long, skey As String, ByVal rva As Long, frmJI As TreeView)
'frmJI.Arbre.Nodes.Add fu, 4, , offset, 16
Dim i As Long, oFs As Long, lDump As Long
Dim lkey As String, decs As String
On Local Error Resume Next

    oFs = rva - exeOPHEAD.ImageBase + 1
    For i = 1 To 12
        decs = "+" & Format$((i - 1) * 4, "00") & " "
        Get #fp, oFs, lDump
        If lDump > min And lDump < max Then
            lkey = "a" & lDump
            frmJI.Nodes.Add skey, 4, lkey, decs & Hex$(lDump), 16
            If Err.Number <> 35602 Then
                Call Util_SnifNext(min, max, fp, fu + 1, lkey, lDump, frmJI)
            Else
                Err.Clear
                frmJI.Nodes.Add skey, 4, , decs & Hex$(lDump), 16
            End If
        Else
            frmJI.Nodes.Add skey, 4, , decs & Right$("0000000" & Hex$(lDump), 8) & " - " & ScanString(fp, oFs), 18
        End If
        oFs = oFs + 4
    Next i
    
End Sub

Public Sub Util_BackStart(frmJI As TreeView, ByVal epva As Long)
'scanner inverse d'appel RVA
Dim fp As Integer
Dim lDump As Long, lng, i, j, nva, cva, ptr, rva
Dim followup
Dim lkey As String, pkey As String
Dim bDump() As Long
Dim stxt As String
ReDim bDump(1 To (exeFILENAMEsize - 4096) / 4)
lng = UBound(bDump)

    fp = FreeFile
    Open exeFILENAMElong For Binary Access Read As #fp
    followup = 1
    Get #fp, 4097, bDump()
    lDump = bDump(epva / 4 - 1023)
    pkey = "a" & CStr(epva + exeOPHEAD.ImageBase)
    frmJI.Nodes.Add , , pkey, "Offset " & Hex$(epva) & " : " & Right$("0000000" & Hex$(lDump), 8) & " - " & ScanString(fp, epva + 1), 1
    nva = epva + exeOPHEAD.ImageBase
    
    Call Util_BackNext(nva, bDump(), lng, pkey, frmJI)
    
    Close #fp

End Sub

Private Sub Util_BackNext(ByVal rva_to_find As Long, ByRef bDump() As Long, ByVal lng As Long, pkey As String, frmJI As TreeView)
Dim i, j, cva, nrva
Dim lkey As String
Static fu As Long
On Local Error Resume Next
fu = fu + 1
If fu > 150 Then Exit Sub

    For i = 0 To -10 Step -1
        cva = rva_to_find + (i * 4)
        For j = 1 To lng
            If bDump(j) = cva Then
                nrva = j * 4 + 4092
                lkey = "a" & nrva
                frmJI.Nodes.Add pkey, 4, lkey, "Offset " & Hex$(nrva) & " ask jmp to " & Hex$(cva) & " ( soit " & Hex$(rva_to_find - exeOPHEAD.ImageBase) & " + " & (i * 4) & ")", 16
                If Err.Number = 35602 Then
                    Err.Clear
                    Exit Sub
                End If
                'If nrva = (exeVB_VBEP - exeOPHEAD.ImageBase + 48) Then Stop
                Call Util_BackNext(nrva + exeOPHEAD.ImageBase, bDump(), lng, lkey, frmJI)
            End If
        Next j
    Next i

End Sub

