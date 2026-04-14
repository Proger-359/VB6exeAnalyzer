Attribute VB_Name = "PeReaderExt"
'========================
' PEREADEREXT
'*************
'
'Extension pour l'intégration des résultats fourni par perdr.exe
'Copyright : Proger
'Création : janvier 2006

DefLng A-Z
Option Explicit
Option Compare Binary

Private PERDR_location As String

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal Millisecs As Long)
Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Function ShortName(fname As String) As String
Dim sln As Long, strn As String

    strn = String$(197, 0)
    sln = GetShortPathNameA(fname, strn, 196)
    ShortName = Left$(strn, sln)
    
End Function
Function CheckPERDR() As Boolean
Dim sd As String

    sd = Dir(App.Path & "\perdr.exe", vbNormal)
    If sd = "" Then
        CheckPERDR = False
    Else
        PERDR_location = ShortName(App.Path & "\perdr.exe")
        CheckPERDR = True
    End If

End Function

Sub ExecPERDR(exeToUnAsm As String)
'execute perdr et attend la fin de l'exécution
Dim ph As Long, fh As Long, lfp As Integer
Dim sh As String

    sh = PERDR_location & " " & ShortName(exeToUnAsm) & " > tmp-unasm.txt"
    lfp = FreeFile
    Open App.Path & "\exec.bat" For Output As #lfp
        Print #lfp, "cd " & App.Path
        Print #lfp, sh
    Close #lfp
    
    ph = Shell(App.Path & "\exec.bat", vbNormalFocus)
    
    Do
        fh = OpenProcess(1024&, 0, ph)
        Sleep 5&
        If fh = 0 Then Exit Do
        Call CloseHandle(fh)
        DoEvents
    Loop
    
End Sub

Function ParsePERDR() As Boolean
'récupere un désassemblage de PERDR
Dim fp As Integer
Dim gline As String
Dim rvh As Long, ep As Long, et As Long
Dim c As Long, p As Long
Dim gsline() As String

    If Dir(App.Path & "\tmp-unasm.txt") = "" Then
        ParsePERDR = False
        Exit Function
    End If
    
    Erase unASM.ASM_LIST
    
    ep = PEexe.exeVB_CODEENTRY + PEexe.exeOPHEAD.ImageBase
    et = PEexe.exeVB_CODEENTRY + PEexe.exeVB_CODELEN + PEexe.exeOPHEAD.ImageBase
    fp = FreeFile
    Open App.Path & "\tmp-unasm.txt" For Input Access Read As #fp
        c = 0
        Do Until EOF(fp)
            Line Input #fp, gline
            If Len(gline) > 8 Then
                rvh = CLng(Val("&h" & Left$(gline, 8)))
                If (rvh >= ep) And (rvh <= et) Then
                    c = c + 1
                    
                    ReDim Preserve ASM_LIST(1 To c) As VBDEASM
                    gsline = Split(gline, Chr$(9), , vbBinaryCompare)
                    ASM_LIST(c).rvaCode = rvh
                    'ASM_LIST(c).sHexDump = bArrayHexStr(bArray(), AdV)
                    If UBound(gsline) = 2 Then
                        ASM_LIST(c).sUnAsm = gsline(1) & " " & gsline(2)
                    Else
                        ASM_LIST(c).sUnAsm = gsline(1)
                    End If
                End If
                rvh = 0
            End If
            
        Loop

    Close #fp

End Function


Sub PeRDR_UnAsmVB(ByVal EntryPoint As Long, ByVal fp As Integer, ByVal codelen As Long, ByVal EntryRva As Long)
Dim i, c, o, rva, AdV   'variables d'itération et de progression dans le code
Dim ta, op              'variables de progression d'index
Dim p, jrva, cl, cs, cp 'variables de l'interpréteur/analyseur
Dim bDump As Byte, iDump As Integer, lDump As Long
Dim bArray(1 To 14) As Byte
Dim sBuffer As String, sB2 As String


    o = EntryPoint
    rva = EntryRva
    For ta = 1 To UBound(ASM_LIST) - 1
        'ini
        jrva = 0
    
        'lit
        Get #fp, o, bDump
        Get #fp, o, iDump
        Get #fp, o, bArray()

        AdV = ASM_LIST(ta + 1).rvaCode - ASM_LIST(ta).rvaCode
        ASM_LIST(ta).sHexDump = unASM.bArrayHexStr(bArray(), AdV)
        sBuffer = ASM_LIST(ta).sUnAsm
        ASM_LIST(ta).imDump = iDump
            
            'PERDR : récupération de l'adresse jump RVA (super lent, mais bon...)
            If bDump = &H68 Or bDump = &HE8 Or bDump = &HBA Or bDump = &H8B Or bDump = &HFF Or bDump = &HC7 Then
                p = InStr(3, sBuffer, "004", vbBinaryCompare)
                If p > 0 Then
                    jrva = Val("&h" & Mid$(sBuffer, p, 8))

                    If jrva < &H400000 Then
                        jrva = 0
                    Else
                        If (jrva > PEexe.exeOPHEAD.ImageBase) And (jrva - PEexe.exeOPHEAD.ImageBase) < PEexe.exeFILENAMEsize Then
                            ASM_LIST(ta).rvaJump = jrva
                        Else
                            jrva = 0
                        End If
                    End If
                End If
            End If
            
NAsm:
    If o Mod 12000 = 0 Then
        'indicateur de progression non bloquant
        frmPeExe.AddInfo "Désassemblage à " & Int((o - EntryPoint) / (cl - EntryPoint) * 100) & "%..."
    End If
    
    o = o + AdV
    rva = rva + AdV

    Next ta
            

End Sub


