Attribute VB_Name = "PEexe"
'=======================
' PE EXE
'********
'
'Par Proger
'aout 2003
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

'source à usage privé (et hop, l'excuse universelle pour dire basta aux commentaires)

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
    End Type
    Private exeIMPORT_DLLNAME() As IMPORT_DLL_LOOKUP
    Private exeIMPORT_APINAME() As IMPORT_API_LOOKUP

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
        sName As String 'nom "name" du contrôle
        id As Integer   'identifiant type de contrôle
        sType As String 'type de contrôle
        Offset As Long  'offset physique (pour récupérer les propriétés/attributs)
    End Type
        Private Type CONTROL_IDTYPE 'type de contrôle (pure VB seulement)
            inID As Integer
            cType As String
        End Type
        Private vbDEFCTRL() As CONTROL_IDTYPE
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
Private exeVB_CONTROL() As CONTROL_DEF
    Public exeVB_PROJECTNAME As String      'nom du projet :)
    Public exeVB_FORMS() As CONTROL_FORM
    Private exeVB_CTRL_PRP() As CONTROL_PROPERTY
    
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
    sFrom As String     'nom de la feuille dans laquel l'API est déclaré (en debug)
    RvaOffset As Long   'offset où l'api est déclaré dans l'exe
End Type
    Private exeVB_API() As APIDECLARE



'déclaration pour un morceau de fonction compilé (en dev)
Private exeVB_CODEENTRY As Long 'pointeur vers le début des fonctions compilées
Private exeVB_CODELEN As Long   'longueur de l'ensemble des fonctions compilées
Public exeVB_CODEMAIN As Long  'point d'entrée vers la fonction appelé au démarrage (Sub Main)
'Table de strucutre de feuille (modules, forms...) tel qu'elle est enregistré physiquement
'dans l'exe compilé vb6. Les informations associés à la variables sont "ce qui semble être" et non pas "sûre"
Private Type VBTBLSTRUCT
    Fill1 As Long   'toujours &hffffffff
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
    rva1 As Long    'pointeur d'association
End Type
    'type pour vbanalyse
    Type VBMODULE
        sName As String
        lType As Long
        RvaOffset As Long
        FullLen As Long
        NumSub As Long
    End Type
    Type VBSUB  'description d'un sub() utilisateur (programmé par)
        sName As String     'si possible
        sParams As String   'si possible
        rvaCode As Long     'point d'entrée dans le code compilé
        CodeLen As Long     'longueur du code compilé (jusqu'a une instruction RET)
        SubType As Long     'type de sub
        SubFrom As Long     'origine du sub (feuille dans laquelle il est codé...)
    End Type
        Private exeVB_MODULES() As VBMODULE
        Private exeVB_SUBS() As VBSUB



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
    Rva As Long
    Ordinal As Long
    uName As String
    uDescr As String
End Type
Private exeVB6_APIDEF() As API_VBDEF



'l'api qui sauve la vie :)
Private Declare Sub MemCpy Lib "Kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, From As Any, ByVal Length As Long)

Private exePEHEAD As PE_HEADER
Private exeOPHEAD As OPTIONAL_HEADER
Private exeHEADIR As HEAD_DIRECTORIES
Private exeOTABLE() As OBJECT_TABLE
Public exeISVB As Boolean
Public exeISPACKED As Boolean

Public exeFILENAMElong As String  'nom du dernier fichier étudié

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
        Get #fp, 61, pefp   'offset du début PE
        Get #fp, pefp + 1, bArray() 'récupère les octets
        'copie les octets dans les structures
        MemCpy exePEHEAD.Signature, bArray(1), PEl
        MemCpy exeOPHEAD.SizeOfOptionalHeader, bArray(PEl + 1), OPl
        MemCpy exeHEADIR.RvaEXPORT_TABLE, bArray(PEl + OPl + 1), HDl
        
        'ces informations de header sont celle du pack... donc inutile...
        
        exeISVB = False
        For i = 2048 To LOF(fp)
            'recherche "VB5!"
            Get #fp, i, vb5
            If vb5 = 557138518 Then
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
    ParseVBFunc (fp)
    
    Close #fp
    
End Sub

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
                ReDim exeIMPORT_DLLNAME(1 To 1)
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
            Do While (exeIMPORT.OriginalFirstThunk <> 0)
                n = n + 1
                Call ParseImportTable(fp, n, exeIMPORT)
                i = Len(exeIMPORT) + i
                If (i > exeHEADIR.RvaTOTAL_IMPORT_DATA_SIZE) Then Exit Do 'antibug
                MemCpy exeIMPORT.OriginalFirstThunk, bArray(1 + i), Len(exeIMPORT)
                'boucle jusqu'a la dernière dll de la liste import
            Loop
        End If


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
            'objets, contrôles
            Call FindControl(fp, exeOPHEAD.EntryPointRVA)
            'copyright
            'Call ParseCopyright(fp, exeTBL_RESOURCE(5).RvaDATA + 1, exeTBL_RESOURCE(5).Size)
            
            'module
            Call FindModules(fp, exeOPHEAD.EntryPointRVA)
            
            'table de sub
            Call ParseVBFunc(fp)
            
        End If
        
    Close #fp
    
    
    
    
End Sub

Sub ParseCopyright(FilePointer As Integer, ByVal EP As Long, ByVal MaxLen As Long)
'je vous ai dit que c'est en dev !
Dim i, j, k
Dim Tstr As String
Dim iArray(1 To 50) As Integer

    EP = EP + 6
    Tstr = ScanUnicode(FilePointer, EP)
    If Tstr = "VS_VERSION_INFO" Then
        EP = EP + 92
        Tstr = ScanUnicode(FilePointer, EP)
        If Tstr = "VarFileInfo" Then
            EP = EP + 128: j = 10
            For i = 1 To 9
            Tstr = ScanUnicode(FilePointer, EP)
            Select Case Tstr
            Case "CompanyName"
                EP = EP + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.CompanyName = Tstr
            Case "FileDescription"
                EP = EP + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.FileDescription = Tstr
            Case "LegalCopyright"
                EP = EP + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.LegalCopyright = Tstr
            Case "LegalTrademarks"
                EP = EP + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.LegalTrademarks = Tstr
                j = 8
            Case "ProductName"
                EP = EP + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.ProductName = Tstr
                j = 10
            Case "FileVersion"
                EP = EP + Len(Tstr) * 2 + 4
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.FileVersion = Tstr
            Case "ProductVersion"
                EP = EP + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.ProductVersion = Tstr
            Case "InternalName"
                EP = EP + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.InternalName = Tstr
                j = 8
            Case "OriginalFilename"
                EP = EP + Len(Tstr) * 2 + 2
                Tstr = ScanUnicode(FilePointer, EP)
                exeVB_COPYRIGHT.OriginalFilename = Tstr
            End Select
            EP = EP + Len(Tstr) * 2 + j
            Next i
        End If
    End If

End Sub
Private Function ScanUnicode(fp As Integer, ByVal Offset As Long) As String
'renvoi la chaine commençant à l'offset Ofs lorsque elle est de type unicode dans le fichier
Dim B1 As Byte, B2 As Byte, i As Long
i = 1

    Get #fp, Offset, B1
    Get #fp, Offset + 1, B2
    
    Do
        ScanUnicode = ScanUnicode & " "
        MidB$(ScanUnicode, i, 1) = Chr$(B1)
        MidB$(ScanUnicode, i + 1, 1) = Chr$(B2)
        i = i + 2
        Offset = Offset + 2
        Get #fp, Offset, B1
        Get #fp, Offset + 1, B2
    Loop Until (B1 + B2) = 0

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
Dim TblLst(1 To 4) As Long
Dim TblDef(1 To 22) As Long
Dim TblStruct() As VBTBLSTRUCT

    'récupère le pointeur vers "VB5!"
    Get #FilePointer, ProgEntryPoint + 2, PtrStr
    
    'récupère le pointeur vers les defs de situation
    PtrTbl = PtrStr - exeOPHEAD.ImageBase + 49
    Get #FilePointer, PtrTbl, PtrBas
    
    'récupère le pointeur vers le sub Main() (départ d'exécution) s'il existe (sinon = 0)
    Get #FilePointer, PtrTbl - 4, PtrSubMain
    If PtrSubMain > 0 Then
        PtrSubMain = PtrSubMain - exeOPHEAD.ImageBase
        exeVB_CODEMAIN = PtrSubMain
    Else
        exeVB_CODEMAIN = 0
    End If
    
    'récupère les 3 pointeurs de situations
    PtrBas = PtrBas - exeOPHEAD.ImageBase
    Get #FilePointer, PtrBas + 5, TblLst()
    
    TBLdeb = TblLst(1) - exeOPHEAD.ImageBase    'tables de structure
    ASMdeb = TblLst(3) - exeOPHEAD.ImageBase    'début du code ASM compilé
    ASMend = TblLst(4) - exeOPHEAD.ImageBase    'fin du code ASM compilé
    exeVB_CODEENTRY = ASMdeb
    exeVB_CODELEN = ASMend - ASMdeb
    
    'récupération des tables de structure
    Get #FilePointer, TBLdeb + 1, TblDef()
    'le nombre de table de structure est un Integer, pour le récupérer a partir d'un Long, on fait un bitmasking
    k = (TblDef(12) And &HFFFF0000) / 65536
    ReDim TblStruct(1 To k)
    ReDim exeVB_MODULES(1 To k): i = 1
    Get #FilePointer, TblDef(13) - exeOPHEAD.ImageBase + 5, TblStruct()
    
    'initialise le tableau local s'il y a des API
    ReDim exeVB_API(0)
    
    'antibug... mouais
    PtrBas = TblDef(22) - exeOPHEAD.ImageBase + 1
    exeVB_MODULES(1).sName = ScanString(FilePointer, TblStruct(1).rva6 - exeOPHEAD.ImageBase + 1)
    GoSub ScanTab   'certaines API sont appelé ici
    
    'récupère les informations sur les feuilles dans l'exe
    For i = 1 To k
        exeVB_MODULES(i).lType = TblStruct(i).mType
        exeVB_MODULES(i).sName = ScanString(FilePointer, TblStruct(i).rva6 - exeOPHEAD.ImageBase + 1)
        exeVB_MODULES(i).RvaOffset = TblDef(22) - exeOPHEAD.ImageBase + ((i - 1) * 48)
        exeVB_MODULES(i).NumSub = TblStruct(i).nSubs
        
        If TblStruct(i).rva1 <= exeOPHEAD.ImageBase Then Exit For 'mouais
        PtrBas = TblStruct(i).rva1 - exeOPHEAD.ImageBase + 1
        
        GoSub ScanTab
        
    Next i
    
    Exit Sub

    
ScanTab:
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
                Call ParseDeclares(FilePointer, PtrTbl - exeOPHEAD.ImageBase + 1, exeVB_MODULES(i).sName)
            Case Else
                Exit Do
        End Select
    Loop
    End If
Return

End Sub

Sub ParseDeclares(fp As Integer, ByVal DeclareEntryPoint As Long, ByVal sFormOrigine As String)
'récupère les API déclaré sous VB et les classe par ordre de dll d'origine
Dim Tapi As APISTRUCT
Dim i, j, k
    
    Get #fp, DeclareEntryPoint, Tapi
    j = UBound(exeVB_API()) + 1

    ReDim Preserve exeVB_API(j)
    exeVB_API(j).RvaOffset = DeclareEntryPoint
    exeVB_API(j).sName = ScanString(fp, Tapi.sAPI - exeOPHEAD.ImageBase + 1)
    exeVB_API(j).sDll = ScanString(fp, Tapi.sDll - exeOPHEAD.ImageBase + 1)
    exeVB_API(j).sFrom = sFormOrigine
    
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

Sub ParseControl(FilePointer As Integer, ByRef FormsDef() As CONTROL_FORM)
'récupère les différents contrôles utilisés dans les forms de l'exe vb
Dim OffsetStart As Long 'début du bloc info contrôle
Dim SizeBlock As Long   'taille du bloc contenant l'info sur les conrôles
Dim NumCtrl As Long     'nombre de contrôles dans le bloc
Dim ObjBlock As Long    'taille d'un segment (définition d'1 contrôle)
Dim NameLen As Integer  'longueur du nom d'un contrôle (sur 2 octets)
Dim bCol As Byte, bNum As Integer  'collection : identifiant, numéro
Dim i As Long, j As Long, BlkEnd As Long, k As Long
Dim BugCheck As Byte '(... ben oui ya des truc que je pige pas encore)


    Call Init_VBCTRL    'charge l'identificateur de contrôles

i = 1
For j = 1 To UBound(FormsDef())
    
    
    OffsetStart = FormsDef(j).rvaPtr
    FormsDef(j).DefPtr = i
    
    'nombre de contrôles (????)
    Get #FilePointer, OffsetStart + 4, NumCtrl
    'on inverse les octets (ben oui il faut, sinon ya faux!), seul les 2 derniers octets sont significatifs
    NumCtrl = (((NumCtrl And 255) * 256) Or ((NumCtrl And 65280) / 256)) + 1 '(le +1 est pour le form inclu)
    
    
    'taille du bloc
    OffsetStart = OffsetStart + 90
    Get #FilePointer, OffsetStart, SizeBlock
    BlkEnd = OffsetStart + SizeBlock - 1
    
    OffsetStart = OffsetStart + 4
    
    k = 0 '<<== sert a compter le nb d'objets trouvé (encore en phase de test)
    Do
        'If k > NumCtrl Then Exit Do

Cscan:
        Get #FilePointer, OffsetStart, ObjBlock
        
        If ObjBlock = 68 And OffsetStart + 4 > BlkEnd Then Exit Do
        
        Get #FilePointer, OffsetStart + 7, BugCheck
        If (ObjBlock > SizeBlock) Or BugCheck = 0 Then
            'bizarre : parfois il y a un décalage d'octet (conteneur ??)
            OffsetStart = OffsetStart + 1
            If OffsetStart > BlkEnd Then Exit Do
            GoTo Cscan
            
        ElseIf ObjBlock < -1 Then
            'il y a un flag "1" sur le bit 31
            'en général, il s'agit un groupe d'objets.
            ObjBlock = (ObjBlock And &H7FFFFFFF)
            ReDim Preserve exeVB_CONTROL(1 To i)
            exeVB_CONTROL(i).Offset = OffsetStart
            Get #FilePointer, OffsetStart + 4, bCol
            Get #FilePointer, OffsetStart + 5, bNum
            Get #FilePointer, OffsetStart + 7, NameLen
            exeVB_CONTROL(i).sName = ScanString(FilePointer, OffsetStart + 9) & " (" & bCol & "-" & bNum & ")"
            Get #FilePointer, OffsetStart + 10 + NameLen, exeVB_CONTROL(i).id
                If (exeVB_CONTROL(i).id And 255) = 255 Then
                    'il s'agit d'un contrôle "externe" a vb (ocx...)
                    exeVB_CONTROL(i).sType = ScanString(FilePointer, OffsetStart + 13 + NameLen)
                Else
                    exeVB_CONTROL(i).sType = Get_VBCTRL(exeVB_CONTROL(i).id)
                End If

            OffsetStart = OffsetStart + ObjBlock + 1
            i = i + 1: k = k + 1
            
        ElseIf ObjBlock = -1 Then
            'aaakkk :( bug impossible si on ouvre un exe vb6!!!
            Stop
        Else
            
            ReDim Preserve exeVB_CONTROL(1 To i)
            exeVB_CONTROL(i).Offset = OffsetStart
            Get #FilePointer, OffsetStart + 5, NameLen
            exeVB_CONTROL(i).sName = ScanString(FilePointer, OffsetStart + 7)
            Get #FilePointer, OffsetStart + 8 + NameLen, exeVB_CONTROL(i).id
                If (exeVB_CONTROL(i).id And 255) = 255 Then
                    'il s'agit d'un contrôle "externe" a vb (ocx...)
                    exeVB_CONTROL(i).sType = ScanString(FilePointer, OffsetStart + 11 + NameLen)
                Else
                    'contrôle vb interne
                    exeVB_CONTROL(i).sType = Get_VBCTRL(exeVB_CONTROL(i).id)
                End If
    
            OffsetStart = OffsetStart + ObjBlock + 1
            i = i + 1: k = k + 1
        End If
        
        If OffsetStart >= BlkEnd Then Exit Do

    Loop


    FormsDef(j).DefLen = i - FormsDef(j).DefPtr

Next j

End Sub

Sub ParseControlParams(FilePointer As Integer, ByRef tblVBCTRL() As CONTROL_DEF)
'récupère les paramètres des contrôles vb
'il y a un peu de redondance par rapport au code parser de bloc... tant pis
Dim i, j, l
Dim Offs As Long, iSize As Long, rvaEnd As Long
Dim bArray() As Byte
Dim gpLng As Long, gpInt As Integer, gpByte As Byte
Dim useStr As String

    l = UBound(tblVBCTRL())
    ReDim exeVB_CTRL_PRP(1 To l)
    
    For i = 1 To l
    
        Offs = tblVBCTRL(i).Offset
        Get #FilePointer, Offs, iSize
        If iSize < -1 Then
            'bit 31 = 1 : collection d'objet
            iSize = iSize And &H7FFFFFFF
            rvaEnd = Offs + iSize
            Offs = Offs + 7
        Else
            rvaEnd = Offs + iSize
            Offs = Offs + 5
        End If
        
        
        'saute l'attribut .Name
        Get #FilePointer, Offs, gpInt
        Offs = Offs + gpInt + 3
        
        'identifiant
        Get #FilePointer, Offs, gpInt
        'saute l'information sur un contrôle externe (ocx)
        If (gpInt And 255) = 255 Then
            Get #FilePointer, Offs + 4, gpInt
            Offs = Offs + 6 + gpInt
        Else
            Offs = Offs + 2
        End If
        
        'ici commence les attribut intrinsèque de l'objet/contrôle en cours d'étude :)
        '(bon y'en a un paquet, il faudra du temps avant de tous les récupérer...)
        Select Case (gpInt And 255)
            Case 255 'contrôle ocx (externe a vb)
            
            Case 1  'VB.Label
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                
                'backcolor?
                Offs = Offs + 2 + gpInt
                Get #FilePointer, Offs, gpLng
                Get #FilePointer, Offs + 5, gpByte
                If gpByte = 0 Then
                    Offs = Offs + 5
                End If
                'forecolor?
                Get #FilePointer, Offs, gpLng
                Get #FilePointer, Offs + 5, gpByte
                If gpByte = 0 Then
                    Offs = Offs + 5
                End If

                'left
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pLeft = gpInt
                'top
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pTop = gpInt
                'width
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pWidth = gpInt
                'height
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pHeight = gpInt
                
                Get #FilePointer, Offs + 2, gpInt
                If gpInt = 275 Then
                    'borderstyle = 3D
                    Offs = Offs + 2
                End If
                
                Get #FilePointer, Offs + 2, gpInt
                If gpInt = 276 Or gpInt = 532 Then
                    'alignement = right or center
                    Offs = Offs + 2
                End If
                
                Get #FilePointer, Offs + 2, gpInt
                If gpInt = &HFF1E Then
                    'wordwrap = true
                    Offs = Offs + 2
                End If

                Get #FilePointer, Offs + 2, gpInt
                If gpInt = 31 Then
                    'Backstyle = transparent
                    Offs = Offs + 2
                End If

                Get #FilePointer, Offs + 2, gpInt
                If gpInt = 36 Then
                    'usemnemonic = True
                    Offs = Offs + 2
                End If

                Get #FilePointer, Offs + 2, gpInt
                If gpInt = 42 Then
                    'tooltiptext = ...
                    Offs = Offs + 2
                End If
                
            Case 2 'VB.TextBox
            
                'backcolor?
                Get #FilePointer, Offs - 1, gpByte
                Get #FilePointer, Offs, gpLng
                If gpByte = 2 Then
                    Offs = Offs + 5
                End If
                'forecolor?
                Get #FilePointer, Offs - 1, gpByte
                Get #FilePointer, Offs, gpLng
                If gpByte = 3 Then
                    Offs = Offs + 5
                End If

                'left
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pLeft = gpInt
                'top
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pTop = gpInt
                'width
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pWidth = gpInt
                'height
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pHeight = gpInt

            
            Case 3  'VB.Frame
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)

            Case 4  'VB.CommandButton
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                
                Offs = Offs + 4 + gpInt
                
                'forecolor?
                Get #FilePointer, Offs - 1, gpByte
                Get #FilePointer, Offs, gpLng
                If gpByte = 2 Then
                    Offs = Offs + 5
                End If
                'backcolor?
                Get #FilePointer, Offs - 1, gpByte
                Get #FilePointer, Offs, gpLng
                If gpByte = 3 Then
                    Offs = Offs + 5
                End If

                'left
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pLeft = gpInt
                'top
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pTop = gpInt
                'width
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pWidth = gpInt
                'height
                Offs = Offs + 2
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pHeight = gpInt

                
            Case 5  'VB.CheckBox
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
            
            Case 6  'VB.OptionButton
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
            
            Case 13 'VB.Form
                'caption
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                '?
                'fillcolor
                '?
                'picture
                
                ''linktopic
                'Offs = Offs + gpInt + 14
                'Get #FilePointer, Offs, gpInt
                ''(linktopic string ici)
                
                'dépendance
                Offs = rvaEnd
                Do
                    Offs = Offs - 1
                    Get #FilePointer, Offs, gpByte
                Loop Until gpByte = 255
                
                Get #FilePointer, Offs - 4, gpInt
                If gpInt = 68 Then
                    'ShowInTaskBar = False
                    Offs = Offs - 2
                End If
                
                Offs = Offs - 18
                'left
                'Offs = Offs + gpInt + 4
                Get #FilePointer, Offs, gpLng
                exeVB_CTRL_PRP(i).pLeft = gpLng
                'top
                Offs = Offs + 4
                Get #FilePointer, Offs, gpLng
                exeVB_CTRL_PRP(i).pTop = gpLng
                'width
                Offs = Offs + 4
                Get #FilePointer, Offs, gpLng
                exeVB_CTRL_PRP(i).pWidth = gpLng
                'height
                Offs = Offs + 4
                Get #FilePointer, Offs, gpLng
                exeVB_CTRL_PRP(i).pHeight = gpLng
                
            Case 19 'VB.Menu
                'tweak de Urgo (vbfrance)
                Get #FilePointer, Offs, gpInt
                If gpInt = 768 Then
                   exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 4)
                Else
                   exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                End If
                '/tweak
                
                Offs = rvaEnd - 3
                'Offs = Offs + gpInt + 3
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).pRank = gpInt



            Case 20 'VB.MDIForm
                Get #FilePointer, Offs, gpInt
                exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)

            Case Else
            
        End Select

    Next i

Exit Sub



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

Private Sub ParseVBFunc(fp As Integer)
'recherche les points d'entrée dans le code compilé pour les différents sub -expérimental-
Dim i, j, cn, sBe
Dim t1, t2, t3, t4
t1 = 1

    i = 0
    ReDim exeVB_SUBS(1)
    cn = UBound(exeIMPORT_APINAME())    'pointeur fin des import d'api
    cn = 4096 + (4 * cn) + 1
    
    Do
DebLoop:
        cn = cn + 4
        Get #fp, cn, j
        If (j And &H25FF&) = &H25FF& Then Exit Do   'fin de la liste (repéré par un 25ffh ...)
        
        'récupère la liste. Bon ya plein de truc bizarre, dur dur de distingué si c'est un sub ou une func
        If j > 0 Then
            
            If (j And &HFF000000) > 0 Then
                t1 = j And &HFF000000
                
            ElseIf (j And &HFF0000) > 0 Then
                t2 = j And &HFF0000
                If t2 = &H80000 Then
                    If ((j And 65535) = 7) Or ((j And 65535) = 12) Then
                        cn = cn + 12
                    Else
                        cn = cn + 4
                    End If
                    Get #fp, cn, t4
                    GoSub AddNSub
                    exeVB_SUBS(i).SubType = 1
                    exeVB_SUBS(i).SubFrom = 1
                ElseIf t2 = &H40000 Then
                    If (j And 65535) = 4 Then
                        cn = cn + 12
                    Else
                        cn = cn + 8
                    End If
                    Get #fp, cn, t4
                    GoSub AddNSub
                    exeVB_SUBS(i).SubType = 1
                    exeVB_SUBS(i).SubFrom = 2
                End If
                'If t4 = 0 Then Stop
            End If
        End If
        
    Loop
    
Exit Sub

AddNSub:
    i = i + 1
    ReDim Preserve exeVB_SUBS(i)
    exeVB_SUBS(i).rvaCode = t4 - exeOPHEAD.ImageBase
Return
    
End Sub

Private Function ScanString(fp As Integer, ByVal Offset As Long) As String
'scanne une chaine de caractère ANSI, se termine par un 8-bits NULL
Dim b As Byte
Get #fp, Offset, b
Do
    ScanString = ScanString & Chr$(b)
    Offset = Offset + 1
    Get #fp, Offset, b
Loop Until b = 0

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
                ReDim Preserve exeVB6_APIDEF(1 To i)
                exeVB6_APIDEF(i).Rva = Val("&H" & sAdr)
                exeVB6_APIDEF(i).Ordinal = CLng(sOrd)
                exeVB6_APIDEF(i).uName = sName
                exeVB6_APIDEF(i).uDescr = sDef
            Else
                Exit Do
            End If
        Loop Until EOF(1)
    
    Close #lfp

End Sub

Private Function VBfunc_Description(ByVal inOrdinal As Long, ByVal inAPIname As String, ByRef outRName As String) As String
'renvoi la fonction vb associé à l'api vb appelé
Dim i As Long


If inOrdinal > 0 And inAPIname = "" Then
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
Dim l As Long, i As Long, s As Long

    
    i = UBound(outADRarray()) - 1
    Get #fp, OffsetADR, l
    Do
        i = i + 1
        ReDim Preserve outADRarray(1 To i)
        outADRarray(i).Address = l
        Get #fp, OffsetSTR, s
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
                .AddItem "  - offset " & exeVB_CONTROL(i).Offset
                
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
    
    .AddItem exePEHEAD.Signature & "  (" & PE_Check(exePEHEAD.Signature) & ")"
    .AddItem exePEHEAD.CpuType & " : pour CPU " & PE_CpuType(exePEHEAD.CpuType)
    .AddItem exePEHEAD.Objects & " objets (alias sections)"
    
    .AddItem exePEHEAD.TimeDate
    .AddItem exePEHEAD.PointerToSymbolTable
    
    .AddItem exePEHEAD.NumberOfSymbols
    .AddItem exePEHEAD.NThdrSize
    .AddItem exePEHEAD.Flags & " : " & PE_Flags(exePEHEAD.Flags)
    
    .AddItem "===== optional header ====="
    
    .AddItem exeOPHEAD.SizeOfOptionalHeader
    .AddItem exeOPHEAD.LinkMajor
    .AddItem exeOPHEAD.LinkMinor
    .AddItem exeOPHEAD.Reserved1
    
    .AddItem exeOPHEAD.Reserved2
    .AddItem exeOPHEAD.Reserved3
    
    .AddItem "Point d'entrée code : " & exeOPHEAD.EntryPointRVA
    .AddItem exeOPHEAD.Reserved4
    
    .AddItem exeOPHEAD.Reserved5
    .AddItem "ImageBase : " & exeOPHEAD.ImageBase & " (" & Hex(exeOPHEAD.ImageBase) & ")"
    
    .AddItem exeOPHEAD.ObjectAlign
    .AddItem exeOPHEAD.FileAlign
    
    .AddItem exeOPHEAD.OsMajor
    .AddItem exeOPHEAD.OsMinor
    .AddItem exeOPHEAD.UserMajor
    .AddItem exeOPHEAD.UserMinor
    
    .AddItem exeOPHEAD.SubSysMajor
    .AddItem exeOPHEAD.SubSysMinor
    .AddItem exeOPHEAD.Reserved6
    
    .AddItem "Taille image : " & exeOPHEAD.ImageSize
    .AddItem "Taille header : " & exeOPHEAD.HeaderSize
    
    .AddItem "Checksum : " & Hex(exeOPHEAD.FileCheckSum)
    .AddItem exeOPHEAD.SubSystemNT
    .AddItem "Flag de dll : " & exeOPHEAD.DLLflags
    
    .AddItem exeOPHEAD.StackReserveSize
    .AddItem exeOPHEAD.StackCommitSize
    
    .AddItem exeOPHEAD.HeapReserveSize
    .AddItem exeOPHEAD.HeapCommitSize
    
    .AddItem exeOPHEAD.LoaderFlags
    .AddItem exeOPHEAD.NumberOfRvaAndSizes

    .AddItem "===== data directories ====="

    .AddItem exeHEADIR.RvaEXPORT_TABLE
    .AddItem exeHEADIR.RvaTOTAL_EXPORT_DATA_SIZE
    
    .AddItem exeHEADIR.RvaIMPORT_TABLE
    .AddItem exeHEADIR.RvaTOTAL_IMPORT_DATA_SIZE
    
    .AddItem exeHEADIR.RvaRESOURCE_TABLE
    .AddItem exeHEADIR.RvaTOTAL_RESOURCE_DATA_SIZE
    
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
        .AddItem exeOTABLE(i).VirtualSize
        .AddItem exeOTABLE(i).VirtualAddress
    
        .AddItem exeOTABLE(i).SizeOfRawData
        .AddItem exeOTABLE(i).PointerToRawData
    
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
                If exeVB_API(l).sFrom <> BGS Then
                    BGS = exeVB_API(l).sFrom
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

With ATree
    .Nodes.Clear

    .Nodes.Add , 0, "root", exeVB_PROJECTNAME, 3
    
    
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
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Subs contenu : " & exeVB_MODULES(i).NumSub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16

        Case 98435  'form
            cf = cf + 1
            If cf = 1 Then
                .Nodes.Add "root", 4, "form", "Feuilles", 1
                .Nodes.Item("form").ExpandedImage = 2
            End If
            TNs = LCase$(exeVB_MODULES(i).sName)
            .Nodes.Add "form", 4, TNs, exeVB_MODULES(i).sName, 6
            .Nodes.Add TNs, 4, , "Subs contenu : " & exeVB_MODULES(i).NumSub, 19
            .Nodes.Add TNs, 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            For j = 1 To UBound(exeVB_FORMS())
                If TNs = LCase$(exeVB_CONTROL(exeVB_FORMS(j).DefPtr).sName) Then Exit For
            Next j
            
            .Nodes.Add TNs, 4, , "Offset objets : " & Hex$(exeVB_FORMS(j).rvaPtr) & "h", 16
            'rajoute les objets dans la form
            .Nodes.Add TNs, 4, "form_" & TNs, exeVB_FORMS(j).DefLen & " contrôles trouvés.", 18
            For k = exeVB_FORMS(j).DefPtr To exeVB_FORMS(j).DefPtr + exeVB_FORMS(j).DefLen - 1
                If exeVB_CONTROL(k).id = 276 Then .Nodes.Item(TNs).Image = 7    'icone MDI
                
                .Nodes.Add "form_" & TNs, 4, TNs & "_" & k, exeVB_CONTROL(k).sName & VBCTRL_CollectionSpare(exeVB_CONTROL(k).sName), 5
                .Nodes.Add TNs & "_" & k, 4, , "Offset " & Hex$(exeVB_CONTROL(k).Offset) & "h", 16
                .Nodes.Add TNs & "_" & k, 4, , "Type : " & exeVB_CONTROL(k).id & "  (" & exeVB_CONTROL(k).sType & ")", 18
                
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
            Next k
            
        Case 1146883  'classe
            cs = cs + 1
            If cs = 1 Then
                .Nodes.Add "root", 4, "clas", "Modules de classe", 1
                .Nodes.Item("clas").ExpandedImage = 2
            End If
            .Nodes.Add "clas", 4, LCase$(exeVB_MODULES(i).sName), exeVB_MODULES(i).sName, 9
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Subs contenu : " & exeVB_MODULES(i).NumSub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            
        Case 1941507  'contrôles utilisateur
            cu = cu + 1
            If cu = 1 Then
                .Nodes.Add "root", 4, "uctl", "Contrôles utilisateur", 1
                .Nodes.Item("uctl").ExpandedImage = 2
            End If
            .Nodes.Add "uctl", 4, LCase$(exeVB_MODULES(i).sName), exeVB_MODULES(i).sName, 11
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Subs contenu : " & exeVB_MODULES(i).NumSub, 19
            .Nodes.Add LCase$(exeVB_MODULES(i).sName), 4, , "Offset structure : " & Hex$(exeVB_MODULES(i).RvaOffset) & "h", 16
            
        Case Else
            'Stop
            '.Nodes.Add "root", 4, "prop", "Pages de propriétés", 1, 2
        End Select
    Next i

    'affiche plus d'info :
    .Nodes.Add "root", 4, "info", "Informations supplémentaires", 1
    .Nodes.Item("info").ExpandedImage = 2
    
        'DLL importés
        Dim TDs As String, ouR As String
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
                
                    If exeIMPORT_APINAME(j).ApiName = "" Then Exit Do
                
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
        If exeVB_CODEMAIN > 0 Then
            .Nodes.Add "codeep", 4, "sub_main", "Sub Main", 19
            .Nodes.Add "sub_main", 4, , "Offset " & Hex$(exeVB_CODEMAIN) & "h", 16
        End If
        'affiche la liste des subs trouvés
        k = UBound(exeVB_SUBS)
        For i = 1 To k
            'EN DEV (voir ParseVBfunc())
            Select Case exeVB_SUBS(i).SubType
            Case 1
                TDs = "Sub "
            Case 2
                TDs = "Sub "
            Case Else
                TDs = "Sub "
            End Select
            Select Case exeVB_SUBS(i).SubFrom
            Case 1
                TNs = " "
            Case 3
                TNs = " "
            Case 5
                TNs = " "
            Case Else
                TNs = " "
            End Select
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
Dim i, j, k, l, m
Dim bArray() As Byte
Dim Lp As Integer
Dim Tstr As String
With oPCB

    ReDim bArray(1 To (16 * Ligne) + 1)
    Lp = FreeFile
    
    Open Fichier For Binary Access Read As #Lp
        Seek #Lp, OffsetMod16 * 16 + 1
        Get #Lp, , bArray()
    Close #Lp
    
    .AutoRedraw = True
    .Cls
    .FontSize = 8
    .FontName = "Courier"
    
    j = OffsetMod16 * 16
    
    k = 0
    For i = 1 To Ligne
        .CurrentY = i * 150
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

Private Sub Init_VBCTRL()
'les noms indiqués sont ceux des fichiers .frm aux lignes "BEGIN"
    ReDim vbDEFCTRL(24)
    vbDEFCTRL(1).inID = 269
    vbDEFCTRL(1).cType = "VB.Form"
    vbDEFCTRL(2).inID = 549
    vbDEFCTRL(2).cType = "VB.Data"
    vbDEFCTRL(3).inID = 1042
    vbDEFCTRL(3).cType = "VB.FileListBox"
    vbDEFCTRL(4).inID = 523
    vbDEFCTRL(4).cType = "VB.Timer"
    vbDEFCTRL(5).inID = 1041
    vbDEFCTRL(5).cType = "VB.DirListBox"
    vbDEFCTRL(6).inID = 1040
    vbDEFCTRL(6).cType = "VB.DriveListBox"
    vbDEFCTRL(7).inID = 522
    vbDEFCTRL(7).cType = "VB.VScrollBar"
    vbDEFCTRL(8).inID = 521
    vbDEFCTRL(8).cType = "VB.HScrollBar"
    vbDEFCTRL(9).inID = 1032
    vbDEFCTRL(9).cType = "VB.ListBox"
    vbDEFCTRL(10).inID = 1287
    vbDEFCTRL(10).cType = "VB.ComboBox"
    vbDEFCTRL(11).inID = 262
    vbDEFCTRL(11).cType = "VB.OptionButton"
    vbDEFCTRL(12).inID = 261
    vbDEFCTRL(12).cType = "VB.CheckBox"
    vbDEFCTRL(13).inID = 260
    vbDEFCTRL(13).cType = "VB.CommandButton"
    vbDEFCTRL(14).inID = 259
    vbDEFCTRL(14).cType = "VB.Frame"
    vbDEFCTRL(15).inID = 1026
    vbDEFCTRL(15).cType = "VB.TextBox"
    vbDEFCTRL(16).inID = 1280
    vbDEFCTRL(16).cType = "VB.PictureBox"
    vbDEFCTRL(17).inID = 792
    vbDEFCTRL(17).cType = "VB.Image"
    vbDEFCTRL(18).inID = 791
    vbDEFCTRL(18).cType = "VB.Line"
    vbDEFCTRL(19).inID = 1046
    vbDEFCTRL(19).cType = "VB.Shape"
    vbDEFCTRL(20).inID = 257
    vbDEFCTRL(20).cType = "VB.Label"
    vbDEFCTRL(21).inID = 803
    vbDEFCTRL(21).cType = "VB.OLE"
    vbDEFCTRL(22).inID = 787
    vbDEFCTRL(22).cType = "VB.Menu"
    vbDEFCTRL(23).inID = 276
    vbDEFCTRL(23).cType = "VB.MDIForm"
    vbDEFCTRL(24).inID = 6440
    vbDEFCTRL(24).cType = "Objet-classe locale ?"

End Sub

Private Function Get_VBCTRL(inIDent As Integer) As String
Dim i
    For i = 1 To 24
        'seul l'octet de poids fort indique le type de contrôle (VB pur)
        If (vbDEFCTRL(i).inID And 255) = (inIDent And 255) Then
            Get_VBCTRL = vbDEFCTRL(i).cType
            Exit Function
        End If
    Next i
    
    Get_VBCTRL = "inconnu" 'ça devrai jamais arrivé ;p
    
End Function

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

Function Utils_EXEfilename(ByRef fullpath As String) As String
'renvoi que le nom du fichier à la fin d'un nom complet (dossier + fichier)
Dim p

    p = InStrRev(fullpath, "\")
    If p > 0 Then
        Utils_EXEfilename = Mid$(fullpath, p + 1)
    Else
        Utils_EXEfilename = fullpath
    End If

End Function
