VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================
' MANIPULATEUR PROCESSUS
'************************
'par Proger - 2002
'Permet de connaitre les processus actif, obtenir leur nom, obtenir les modules (dépendances dll...)
'Permet de copier (dump) la mémoire du processus, de changer localement la mémoire du processus
'
'Sources :
'basé sur quelques sources vbfrance
'basé sur le principe de procdump, rameditor de winhex, et cheat-o-matic
'
'Modifications, simplifications et ajout de nouvelles fonctions grâçe à :
'MSDN juillet 2002 (performance monitoring, memory managment, dll process and threads)
'API-Guide de allapi network
'
'
'Toutes les fonctions renvoi 0 si l'opération est un succès, ou -1 en cas d'erreur
Option Compare Binary
Option Explicit

'type pour connaitre l'utilisation de la mémoire par un processus
Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

'type pour connaitre le type de page mémoire actuellement lu
Private Type MEMORY_BASIC_INFORMATION
   BaseAddress As Long
   AllocationBase As Long
   AllocationProtect As Long
   RegionSize As Long
   State As Long
   Protect As Long
   lType As Long
End Type

'type pour connaitre l'OS, l'organisation de la mémoire et le(s) cpu(s)
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

'type pour connaitre l'occupation mémoire d'un module
Private Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type
    
'listage des processus actif :
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'ouverture/fermeture de processus :
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
'information sur les modules d'un processus :
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByRef lpModInfo As MODULEINFO, ByVal CbSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
'autres informations
Private Declare Function GetProcessHandleCount Lib "kernel32" (ByVal hProcess As Long, pdwHandleCount As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetProcessShutdownParameters Lib "kernel32" (lpdwLevel As Long, lpdwFlags As Long) As Long
'manipulation de la mémoire d'un processus :
Private Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'fermeture d'un processus :
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'"suicide" :
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
'infos système :
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

'copie de mémoire (utile pour copier rapidement un tableau dans un autre)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal lpDest As Any, ByVal lpSource As Any, ByVal cbCopy As Long)

'mode d'ouverture possible d'un processus (api openprocess())
Private Const PROCESS_TERMINATE = 1
Private Const PROCESS_VM_OPERATION = 8
Private Const PROCESS_VM_READ = 16
Private Const PROCESS_VM_WRITE = 32
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

'Etat de la page
Private Const MEM_FREE = &H10000        '65536  mémoire libre innaccessible
Private Const MEM_RESERVE = &H2000      '8192   mémoire protégé
Private Const MEM_COMMIT = &H1000       '4096   mémoire de travail
'Type de page
Private Const MEM_IMAGE = &H1000000     '16777216 mémoire miroir des exe PE (image des modules)
Private Const MEM_MAPPED = &H40000      '262144   mémoire utilisé pour mapper (?)
Private Const MEM_PRIVATE = &H20000     '131072   mémoire propre au processus (valeurs,...)


Public Function GetActiveProcess(ByRef outTbl() As Long) As Long
'retourne les pID des processus actif dans le tableau
Dim TblSize As Long
Dim TblReturn As Long
Dim exeAPI As Long
Erase outTbl()
ReDim outTbl(1 To 300) As Long 'prépare le tableau à reçevoir 300 processus
TblSize = 300 * 4
exeAPI = EnumProcesses(outTbl(1), TblSize, TblReturn)
If exeAPI <> 0 Then
    'redimensionne le tableau à la longueur exacte du nombre de processus actif
    ReDim Preserve outTbl(1 To TblReturn / 4) As Long
Else
    ReDim outTbl(1 To 1) As Long
    GetActiveProcess = -1
End If
End Function

Public Function KillPID(pIDval As Long) As Long
'force la fermeture d'un processus actif
Dim exeAPI As Long, hPrc As Long
Dim KillDWord As Long
hPrc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, 0, pIDval)
    exeAPI = GetExitCodeProcess(hPrc, KillDWord)
    exeAPI = TerminateProcess(hPrc, KillDWord)
    exeAPI = CloseHandle(hPrc)
'et voui, en 4 ligne on peux killer un process!
If exeAPI = 0 Then KillPID = -1
End Function

Public Function GetPIDName(pIDval As Long) As String
'retourne le chemin d'accès complet à l'exe du processus
Dim OpProcess As Long
Dim exeAPI As Long
Dim ModName As String
Dim BufDW As Long
Dim lngModules(1 To 200) As Long
'le nom du processus est contenu dans le premier module
OpProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDval)
If OpProcess <> 0 Then
    'récupère sous forme de table tout les modules du processus. On a besoin que du premier.
    exeAPI = EnumProcessModules(OpProcess, lngModules(1), 200, BufDW)
    If exeAPI <> 0 Then
        ModName = Space(260) 'valeur standard
        BufDW = 500
        'Recupère le nom du processus
        exeAPI = GetModuleFileNameExA(OpProcess, lngModules(1), ModName, BufDW)
        GetPIDName = Left$(ModName, exeAPI)
    Else
        GetPIDName = "-1"
    End If
End If
exeAPI = CloseHandle(OpProcess) 'toujours fermer un handle processus
End Function

Public Function GetPIDMemInfo(pIDval As Long, WorkSize As Long, PeakSize As Long, PageUsed As Long) As Long
'récupère quelques infos sur le process : la taille occupé en mémoire en cours (work) et maxi (peak), nombre de page mémoire
Dim BlkMem As PROCESS_MEMORY_COUNTERS
Dim CbSize As Long
Dim LhWnd As Long
Dim exeAPI As Long
CbSize = Len(BlkMem)
LhWnd = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDval)
If LhWnd <> 0 Then
    exeAPI = GetProcessMemoryInfo(LhWnd, BlkMem, CbSize)
    exeAPI = CloseHandle(LhWnd)
    WorkSize = BlkMem.WorkingSetSize
    PeakSize = BlkMem.PeakWorkingSetSize
    PageUsed = BlkMem.PagefileUsage
Else
    GetPIDMemInfo = -1
End If
End Function

Public Function GetPIDInfoOthers(pIDval As Long, nHandle As Long, Priority As Long)
'autres infos...
Dim LhWnd As Long
Dim exeAPI As Long

    LhWnd = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDval)
    
    If LhWnd <> 0 Then
        exeAPI = GetProcessHandleCount(LhWnd, nHandle)
        Priority = GetPriorityClass(LhWnd)
        exeAPI = CloseHandle(LhWnd)
    Else
        GetPIDInfoOthers = -1
    End If
    
End Function

Public Function GetPIDModule(pIDval As Long, outModuleName() As String, outModuleID() As Long, outModuleSize() As Long, outModuleAddr() As Long) As Long
'récupère la liste des modules d'un processus, ainsi que leur taille et emplacement mémoire
Dim LhWnd As Long
Dim exeAPI As Long
Dim ModName As String, ModPar As MODULEINFO, ModParLen As Long
Dim BufDW As Long, ic As Long, BufNM As Long
Dim lngModules(1 To 200) As Long
Erase outModuleName(): Erase outModuleID(): Erase outModuleSize(): Erase outModuleAddr() 'reset à zéro les tableaux
ModParLen = Len(ModPar)
LhWnd = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDval)
If LhWnd <> 0 Then
    exeAPI = EnumProcessModules(LhWnd, lngModules(1), 200, BufDW) 'récupère la liste des modules
    If exeAPI <> 0 Then
        BufNM = 500
        ModName = Space(260)
        For ic = 1 To (BufDW / 4)
            ReDim Preserve outModuleName(1 To ic) As String
            ReDim Preserve outModuleAddr(1 To ic) As Long
            ReDim Preserve outModuleID(1 To ic) As Long
            ReDim Preserve outModuleSize(1 To ic) As Long
            'récupère le nom du module
            exeAPI = GetModuleFileNameExA(LhWnd, lngModules(ic), ModName, BufNM)
            outModuleName(ic) = Left(ModName, exeAPI)
            'récupère les infos mémoire du module
            exeAPI = GetModuleInformation(LhWnd, lngModules(ic), ModPar, ModParLen)
            outModuleID(ic) = ModPar.EntryPoint
            outModuleAddr(ic) = ModPar.lpBaseOfDll
            outModuleSize(ic) = ModPar.SizeOfImage
        Next ic
    Else
        GetPIDModule = -1
    End If
End If
exeAPI = CloseHandle(LhWnd)
End Function

Public Function DumpModulePID(pIDval As Long, outByteTbl() As Byte, EntryPoint As Long, ModuleSize As Long) As Long
'copie un module d'un processus vers un tableau de bytes (dump)
Dim HandlePID As Long
Dim MBIInfo As MEMORY_BASIC_INFORMATION, MBILen As Long
Dim exeAPI As Long, TotByteWritten As Long, CumulByte As Long
MBILen = Len(MBIInfo)
Erase outByteTbl(): CumulByte = 1
ReDim outByteTbl(1 To ModuleSize) 'prépare le tableau de dump
HandlePID = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_VM_OPERATION, 0, pIDval)
If HandlePID <> 0 Then
Do
    exeAPI = VirtualQueryEx(HandlePID, ByVal EntryPoint, MBIInfo, MBILen)
    exeAPI = ReadProcessMemory(HandlePID, ByVal EntryPoint, VarPtr(outByteTbl(CumulByte)), MBIInfo.RegionSize, TotByteWritten)
    CumulByte = CumulByte + TotByteWritten
    EntryPoint = MBIInfo.BaseAddress + MBIInfo.RegionSize
Loop Until CumulByte >= ModuleSize Or exeAPI = 0
exeAPI = CloseHandle(HandlePID)
Else
    DumpModulePID = -1
End If
End Function

Public Function DumpValPID(pIDval As Long, outByteTbl() As Byte) As Long
'copie la mémoire de travail d'un processus (valeurs, tableaux,...) vers un tableau de bytes
Dim SysInfo As SYSTEM_INFO
Dim MBIInfo As MEMORY_BASIC_INFORMATION
Dim MBILen As Long
Dim HandlePID As Long, MemPtr As Long
Dim exeAPI As Long, TotByteWritten As Long
Dim ByteBuffer() As Byte, TblPtr As Long
MBILen = Len(MBIInfo)
HandlePID = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_VM_OPERATION, 0, pIDval)
If HandlePID <> 0 Then
    Erase outByteTbl(): TblPtr = 1 'prépare le tableau de dump
    Call GetSystemInfo(SysInfo) 'pour connaitre l'addresse mémoire de base pour applications
    MemPtr = SysInfo.lpMinimumApplicationAddress
    Do Until MemPtr >= SysInfo.lpMaximumApplicationAddress
        'boucle scannant toute la mémoire pour retrouver les pages mémoire du processus
        exeAPI = VirtualQueryEx(HandlePID, ByVal MemPtr, MBIInfo, MBILen)
        If MBIInfo.State = MEM_COMMIT And MBIInfo.lType = MEM_PRIVATE Then
            'c'est une page mémoire du processus!
            If MBIInfo.RegionSize > 0 Then
                ReDim ByteBuffer(1 To MBIInfo.RegionSize) As Byte
                'récupère toute la page
                exeAPI = ReadProcessMemory(HandlePID, ByVal MBIInfo.BaseAddress, VarPtr(ByteBuffer(1)), MBIInfo.RegionSize, TotByteWritten)
                ReDim Preserve outByteTbl(1 To TblPtr + TotByteWritten) As Byte
                'concatène cette page sous forme de tableau au tableau globale de dump
                RtlMoveMemory VarPtr(outByteTbl(TblPtr)), VarPtr(ByteBuffer(1)), TotByteWritten
                TblPtr = TblPtr + TotByteWritten
            End If
        End If
        MemPtr = MBIInfo.BaseAddress + MBIInfo.RegionSize
    Loop
    exeAPI = CloseHandle(HandlePID)
    If TblPtr = 1 Then DumpValPID = -1
Else
    DumpValPID = -1
End If
End Function

Public Function SaveDumpToFile(inDump() As Byte, outFile As String) As Long
'sauvegarde un tableau de bytes vers un fichier.
Dim FF As Integer
If Dir(outFile, vbNormal) = "" Then
    FF = FreeFile
    Open outFile For Binary As #FF
        Put #FF, 1, inDump()
    Close #FF
Else
    SaveDumpToFile = -1
End If
End Function

Public Function LastFileName(inName As String) As String
If Len(inName) > 2 Then
    LastFileName = Mid$(inName, InStrRev(inName, "\") + 1)
End If
End Function

Public Sub CreateTblFromString(inString As String, outTbl() As Byte)
'crée un tableau de bytes à partir d'une chaine de caractères
Dim i As Long
ReDim outTbl(1 To Len(inString)) As Byte
For i = 1 To Len(inString)
    outTbl(i) = Asc(Mid$(inString, i, 1))
Next i
End Sub

Public Function CheatOmatic_Cherche(pIDval As Long, Valeur() As Byte, outResulTbl() As Long) As Long
'recherche les positions de Valeur dans un dump du process pIDval
Dim MBIInfo As MEMORY_BASIC_INFORMATION, MBILen As Long
Dim SysInfo As SYSTEM_INFO, BasePtr As Long
Dim hdProcess As Long, exeAPI As Long

Call GetSystemInfo(SysInfo)
BasePtr = SysInfo.lpMinimumApplicationAddress
MBILen = Len(MBIInfo)
hdProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION, 0, pIDval)
If hdProcess = 0 Then
    CheatOmatic_Cherche = -1
    Exit Function
End If


Dim i As Long, TBW As Long, Lv As Long, n As Long, TempComp As Long
Dim TempSTRvaleur As String
Dim PetitTbl() As Byte

Lv = UBound(Valeur())
TempSTRvaleur = StrConv(Valeur(), vbUnicode)
ReDim PetitTbl(1 To Lv)
ReDim outResulTbl(1 To 1) As Long: n = 1: TempComp = 0
Do
    exeAPI = VirtualQueryEx(hdProcess, ByVal BasePtr, MBIInfo, MBILen)
    If MBIInfo.State = MEM_COMMIT And MBIInfo.lType = MEM_PRIVATE Then
        For i = 0 To MBIInfo.RegionSize - Lv
            exeAPI = ReadProcessMemory(hdProcess, VarPtr(MBIInfo.BaseAddress) + i, VarPtr(PetitTbl(1)), Lv, TBW)
            If TempSTRvaleur = StrConv(PetitTbl(), vbUnicode) Then
                ReDim Preserve outResulTbl(1 To n) As Long
                outResulTbl(n) = MBIInfo.BaseAddress + i
                n = n + 1
            End If
        Next i
    End If
    BasePtr = MBIInfo.BaseAddress + MBIInfo.RegionSize
Loop Until BasePtr >= SysInfo.lpMaximumApplicationAddress
Call CloseHandle(hdProcess)
End Function

Public Function CheatOmatic_ChercheIci(pIDval As Long, NewValeur() As Byte, SeekOffset() As Long, ResultTable() As Long)
'test si les valeurs pointé par SeekOffset() sont égale à NewValeur
Dim BasePtr As Long
Dim hdProcess As Long, exeAPI As Long
Dim MBIInfo As MEMORY_BASIC_INFORMATION, MBILen As Long
MBILen = Len(MBIInfo)
hdProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION, 0, pIDval)
If hdProcess = 0 Then
    CheatOmatic_ChercheIci = -1
    Exit Function
End If
Dim i As Long, TBW As Long, Lv As Long, n As Long, TempComp As Long
Dim PetitTbla() As Byte
Lv = UBound(NewValeur())
ReDim PetitTbla(1 To Lv)

ReDim ResultTable(1 To 1) As Long: n = 1: TempComp = 0
For i = 1 To UBound(SeekOffset())
    exeAPI = VirtualQueryEx(hdProcess, ByVal SeekOffset(i), MBIInfo, MBILen)
    If MBIInfo.State = MEM_COMMIT And MBIInfo.lType = MEM_PRIVATE Then
        exeAPI = ReadProcessMemory(hdProcess, SeekOffset(i), VarPtr(PetitTbla(1)), Lv, TBW)
        If StrConv(NewValeur(), vbUnicode) = StrConv(PetitTbla(), vbUnicode) Then
            ReDim Preserve ResultTable(1 To n) As Long
            ResultTable(n) = SeekOffset(i)
            n = n + 1
        End If
    End If
Next i
Call CloseHandle(hdProcess)
End Function

Public Function CheatOmatic_EcritValeur(pIDval As Long, NewValeur() As Byte, inPosition() As Long) As Long
'écrit la valeur NewValeur dans le process pIDval dans les positions (offset) retenu dans inPosition().
Dim exeAPI As Long, TBW As Long
Dim hdProcess As Long
'ouvre le process
hdProcess = OpenProcess(PROCESS_VM_WRITE Or PROCESS_VM_OPERATION, 0, pIDval)
If hdProcess = 0 Then
    'echec à l'ouverture
    CheatOmatic_EcritValeur = -1
    Exit Function
End If

'remplace les valeurs contenu aux emplacements inRelativePosition par NewValeur
Dim i As Long, ltbl As Long, vallen As Long
vallen = UBound(NewValeur)
ltbl = UBound(inPosition())
For i = 1 To ltbl
    exeAPI = WriteProcessMemory(hdProcess, ByVal inPosition(i), VarPtr(NewValeur(1)), vallen, TBW)
Next i
If TBW = 0 Then CheatOmatic_EcritValeur = -1
Call CloseHandle(hdProcess)
End Function

Public Function GetHexViewTbl(inByteTbl() As Byte, StartOffset As Long, StopOffset As Long, ByRef outHexStr As String) As Long
'renvoi une chaine contenant la vue sous forme "éditeur hexa" d'un tableau de bytes, a partir de StartOffset
Dim FullLn As Long, RestLn As Long
Dim StrSz As Long, TbRsz As Long
Dim inc As Long, Offs As Long, SOffs As Long, minc As Integer, AdV As Long
Dim InternalStr() As Byte

'définition du point de départ
If (StartOffset - 1) Mod 16 = 0 Then
    Offs = StartOffset
Else
    Offs = Int(StartOffset / 16)
End If
'définition du point d'arrivée
If StopOffset <= StartOffset Then StopOffset = StartOffset + 16
TbRsz = UBound(inByteTbl())
Do
    If StopOffset > TbRsz Then
        StopOffset = StopOffset - 16
    End If
Loop Until StopOffset <= TbRsz

If (StopOffset - 1) Mod 16 = 0 Then
    SOffs = StopOffset
Else
    SOffs = (StopOffset / 16) * 16
End If
If SOffs > StopOffset Then SOffs = SOffs - 16
'détermination du nombre de bytes à "hexview-er" et préparations
FullLn = SOffs - Offs + 1
StrSz = FullLn / 16 * 75
If Offs < 1 Then Offs = 1
AdV = 1
ReDim InternalStr(1 To StrSz + 1) As Byte
Dim TempInTbl(1 To 16) As Byte, TempOutTbl(1 To 75) As Byte
For inc = Offs To SOffs Step 16
    RtlMoveMemory VarPtr(TempInTbl(1)), VarPtr(inByteTbl(inc)), 16
    HexLigne TempInTbl(), TempOutTbl(), inc - 1
    RtlMoveMemory VarPtr(InternalStr(AdV)), VarPtr(TempOutTbl(1)), 75
    AdV = AdV + 75
Next inc

outHexStr = StrConv(InternalStr(), vbUnicode)
End Function

Private Sub HexLigne(inTbl() As Byte, outStr() As Byte, OffsetVal As Long)
'transforme une ligne de 16 octets en une ligne formatté hexview
'ReDim outStr(1 To 75) As Byte
Dim inc As Integer
'ligne d'un tableau hexa :
'00000000_00_00_00_00_00_00_00_00_00_00_00_00_00_00_00_00_################ & crlf  = 75 octets

'offset
Z_QuickLongHex OffsetVal, outStr(1), outStr(2), outStr(3), outStr(4), outStr(5), outStr(6), outStr(7), outStr(8)
For inc = 1 To 16
    outStr(6 + inc * 3) = 32 'espace
    Z_QuickByteHex inTbl(inc), outStr(7 + inc * 3), outStr(8 + inc * 3) 'octet en hex
    outStr(inc + 57) = Z_QuickChr(inTbl(inc)) 'char ansi
Next inc
outStr(57) = 32 'sp
outStr(74) = 13 'cr
outStr(75) = 10 'lf
End Sub


Private Function Z_QuickChr(ByteVal As Byte) As Byte
Select Case ByteVal
    Case 0 To 31 'non visible a l'écran => .
        Z_QuickChr = 46
        Exit Function
    Case Else
        Z_QuickChr = ByteVal
        Exit Function
End Select
End Function
Private Sub Z_QuickByteHex(ByteVal As Byte, HiBy As Byte, LoBy As Byte)
Dim HB As Byte, LB As Byte
HB = Int(ByteVal / 16)
LB = ByteVal Mod 16
Select Case HB
    Case 0 To 9
        HiBy = 48 + HB '0 - 9
    Case 10 To 16
        HiBy = 55 + HB 'A - F
End Select
Select Case LB
    Case 0 To 9
        LoBy = 48 + LB
    Case 10 To 16
        LoBy = 55 + LB
End Select
End Sub
Private Sub Z_QuickLongHex(LongVal As Long, B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte, B5 As Byte, B6 As Byte, B7 As Byte, B8 As Byte)
'répartition d'une variable long (dword, 4 octets) en 4 fois 1 octet (byte), façon langage c .
Dim FB1 As Byte, FB2 As Byte, FB3 As Byte, FB4 As Byte
RtlMoveMemory VarPtr(FB1), VarPtr(LongVal), 1
RtlMoveMemory VarPtr(FB2), VarPtr(LongVal) + 1, 1
RtlMoveMemory VarPtr(FB3), VarPtr(LongVal) + 2, 1
RtlMoveMemory VarPtr(FB4), VarPtr(LongVal) + 3, 1
Z_QuickByteHex FB1, B7, B8
Z_QuickByteHex FB2, B5, B6
Z_QuickByteHex FB3, B3, B4
Z_QuickByteHex FB4, B1, B2
End Sub
