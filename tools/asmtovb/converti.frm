VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversion doc Sang Cho vers Init_UnASM()"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Converti !"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Sang Cho to Init_UnASM : par proger pour vb6 analyseur.

DefLng A-Z
Option Explicit

Private Type FLAG_OPC
    sTxt As String
    bFlag As Byte
End Type
Private FlagOPC(0 To 255) As FLAG_OPC
Private PtrOPC(0 To 255) As Long


Private Cont As Long

Private Sub InitFlag()

    'description des flags  (les blancs sont en prévision pour le 64bits
    '/0         1
    FlagOPC(1).sTxt = "/0"
    FlagOPC(1).bFlag = 1
    '/1         2
    FlagOPC(2).sTxt = "/1"
    FlagOPC(2).bFlag = 2
    '/2         3
    FlagOPC(3).sTxt = "/2"
    FlagOPC(3).bFlag = 3
    '/3         4
    FlagOPC(4).sTxt = "/3"
    FlagOPC(4).bFlag = 4
    '/4         5
    FlagOPC(5).sTxt = "/4"
    FlagOPC(5).bFlag = 5
    '/5         6
    FlagOPC(6).sTxt = "/5"
    FlagOPC(6).bFlag = 6
    '/6         7
    FlagOPC(7).sTxt = "/6"
    FlagOPC(7).bFlag = 7
    '/7         8
    FlagOPC(8).sTxt = "/7"
    FlagOPC(8).bFlag = 8
    '           9...
    '/r         17
    FlagOPC(9).sTxt = "/r"
    FlagOPC(9).bFlag = 17
    'r/m8       18
    FlagOPC(10).sTxt = "r/m8"
    FlagOPC(10).bFlag = 18
    'r/m16      19
    FlagOPC(11).sTxt = "r/m16"
    FlagOPC(11).bFlag = 19
    'r/m32      20
    FlagOPC(12).sTxt = "r/m32"
    FlagOPC(12).bFlag = 20
    '           21
    'cb         22
    FlagOPC(13).sTxt = "cb"
    FlagOPC(13).bFlag = 22
    'cw         23
    FlagOPC(14).sTxt = "cw"
    FlagOPC(14).bFlag = 23
    'cd         24
    FlagOPC(15).sTxt = "cd"
    FlagOPC(15).bFlag = 24
    '           25
    'ib         26  cp
    FlagOPC(16).sTxt = "ib"
    FlagOPC(16).bFlag = 26
    'iw         27  cp
    FlagOPC(17).sTxt = "iw"
    FlagOPC(17).bFlag = 27
    'id         28  cp
    FlagOPC(18).sTxt = "id"
    FlagOPC(18).bFlag = 28
    '           29
    '+rb        30
    FlagOPC(19).sTxt = "+rb"
    FlagOPC(19).bFlag = 30
    '+rw        31
    FlagOPC(20).sTxt = "+rw"
    FlagOPC(20).bFlag = 31
    '+rd        32
    FlagOPC(21).sTxt = "+rd"
    FlagOPC(21).bFlag = 32
    '           33
    'rel8       34
    FlagOPC(22).sTxt = "rel8"
    FlagOPC(22).bFlag = 34
    'rel16      35
    FlagOPC(23).sTxt = "rel16"
    FlagOPC(23).bFlag = 35
    'rel32      36
    FlagOPC(24).sTxt = "rel32"
    FlagOPC(24).bFlag = 36
    '           37
    'r8         38
    FlagOPC(25).sTxt = "r8"
    FlagOPC(25).bFlag = 38
    'r16        39
    FlagOPC(26).sTxt = "r16"
    FlagOPC(26).bFlag = 39
    'r32        40
    FlagOPC(27).sTxt = "r32"
    FlagOPC(27).bFlag = 40
    '           41
    'imm8       42
    FlagOPC(28).sTxt = "imm8"
    FlagOPC(28).bFlag = 42
    'imm16      43
    FlagOPC(29).sTxt = "imm16"
    FlagOPC(29).bFlag = 43
    'imm32      44
    FlagOPC(30).sTxt = "imm32"
    FlagOPC(30).bFlag = 44
    '           45
    'ptr16:16   46
    FlagOPC(31).sTxt = "ptr16:16"
    FlagOPC(31).bFlag = 46
    'ptr16:32   47
    FlagOPC(1).sTxt = "ptr16:32"
    FlagOPC(1).bFlag = 47
    '           48
    '           49
    'm          50
    FlagOPC(32).sTxt = "m"
    FlagOPC(32).bFlag = 50
    'm8         51
    FlagOPC(33).sTxt = "m8"
    FlagOPC(33).bFlag = 51
    'm16        52
    FlagOPC(34).sTxt = "m16"
    FlagOPC(34).bFlag = 52
    'm32        53
    FlagOPC(35).sTxt = "m32"
    FlagOPC(35).bFlag = 53
    'm64        54
    FlagOPC(36).sTxt = "m64"
    FlagOPC(36).bFlag = 54
    '           55
    '           56
    'm16:16     60
    FlagOPC(37).sTxt = "m16:16"
    FlagOPC(37).bFlag = 60
    'm16:32     61
    FlagOPC(38).sTxt = "m16:32"
    FlagOPC(38).bFlag = 61
    '           62
    '           63
    'm16&32     64
    FlagOPC(39).sTxt = "m16&32"
    FlagOPC(39).bFlag = 64
    'm16&16     65
    FlagOPC(40).sTxt = "m16&32"
    FlagOPC(40).bFlag = 65
    'm32&32     66
    FlagOPC(41).sTxt = "m32&32"
    FlagOPC(41).bFlag = 66
    '           67
    '           68
    '           69
    'moffs8     70
    FlagOPC(42).sTxt = "moffs8"
    FlagOPC(42).bFlag = 70
    'moffs16    71
    FlagOPC(43).sTxt = "moffs16"
    FlagOPC(43).bFlag = 71
    'moffs32    72
    FlagOPC(44).sTxt = "moffs32"
    FlagOPC(44).bFlag = 72
    '           73
    '           74
    'Sreg
    FlagOPC(45).sTxt = "Sreg"
    FlagOPC(45).bFlag = 80
    
    
    'm32real    128  'fpu
    FlagOPC(46).sTxt = "m32real"
    FlagOPC(46).bFlag = 128
    'm64real    129  'fpu
    FlagOPC(47).sTxt = "m64real"
    FlagOPC(47).bFlag = 129
    'm80real    130  'fpu
    FlagOPC(48).sTxt = "m80real"
    FlagOPC(48).bFlag = 130
    '           131
    'm16int     132  'fpu
    FlagOPC(49).sTxt = "m16int"
    FlagOPC(49).bFlag = 132
    'm32int     133  'fpu
    FlagOPC(50).sTxt = "m32int"
    FlagOPC(50).bFlag = 133
    'm64int     134  'fpu
    FlagOPC(51).sTxt = "m64int"
    FlagOPC(51).bFlag = 134
    '           135
    'ST         159  'fpu
    FlagOPC(52).sTxt = "ST"
    FlagOPC(52).bFlag = 159
    'ST(0)      159  'fpu
    FlagOPC(53).sTxt = "ST(0)"
    FlagOPC(53).bFlag = 159
    'ST(i)      160  'fpu
    FlagOPC(54).sTxt = "ST(i)"
    FlagOPC(54).bFlag = 160
    '+i         160  'fpu
    FlagOPC(55).sTxt = "+i"
    FlagOPC(55).bFlag = 160
    
    'mm         192  'mmx
    FlagOPC(56).sTxt = "mm"
    FlagOPC(56).bFlag = 192
    'mm/m32     200  'mmx
    FlagOPC(57).sTxt = "mm/m32"
    FlagOPC(57).bFlag = 200
    'mm/m64     201  'mmx
    FlagOPC(58).sTxt = "mm/m64"
    FlagOPC(58).bFlag = 201

End Sub

Private Sub Command1_Click()
Dim t As String, S As String
Dim iv As String, InsC As String
Dim i As Long, tTv As Long

'men fout si sa vautre
Open "asm_clean.txt" For Input As #1
Open "asm_vb.txt" For Output As #2
    i = 1
Do
    Line Input #1, t$
    Cont = i
    iv = StrToIv(t$)
    Print #2, iv
    
    tTv = Val("&H" & Left$(t$, 2))
    If PtrOPC(tTv) = 0 Then
        PtrOPC(tTv) = Cont
        S = "TblPtrASM(" & tTv & ")=" & Cont
        Print #2, S
    End If
    
    i = i + 1
Loop Until EOF(1)
Close

Form1.Caption = "Conversion OK"

End Sub


Private Function StrToIv(iNs As String) As String
'analyse une ligne de la doc de Sang Cho
Dim SP As String
Dim FOC As String, OpLen As Long
Dim f1 As Byte, f2 As Byte, f3 As Byte, f4 As Byte, f5 As Byte, f6 As Byte, f7 As Byte, f8 As Byte
Dim sDB As String, sEN As String
Dim St As Long, Ss As Long, Sl As Long, i As Long, j As Long
Dim V As String * 1
Dim G As String * 1
V = ",": G = Chr$(34)

SP = Left$(iNs, 13)

If Mid$(SP, 3, 1) = "+" Then
    f1 = RetFlag(Mid$(SP, 3, 3))
    FOC = "&h" & Left$(SP, 2)
    OpLen = 1
Else
    sDB = Trim$(Mid$(SP, 4, 2))
    FOC = "&h" & sDB & Left$(SP, 2)
    If sDB <> "" Then
        OpLen = 2
    Else
        OpLen = 1
    End If
End If

If Mid$(SP, 6, 1) = "+" Then
    f2 = RetFlag(Trim$(Mid$(SP, 6, 3)))
End If

If Mid$(SP, 7, 1) = "/" Then
    f3 = RetFlag(Mid$(SP, 7, 2))
End If

If Mid$(SP, 8, 1) = "i" Then
    f3 = RetFlag(Mid$(SP, 8, 2))
    If Mid$(SP, 11, 1) <> " " Then
        f4 = RetFlag(Mid$(SP, 11, 2))
    End If
End If

If Mid$(SP, 10, 1) <> " " Then
    f4 = RetFlag(Trim$(Mid$(SP, 10, 3)))
End If


SP = Mid$(iNs, 14, 24)
sDB = "": sEN = "": St = 0: j = 1

For i = 1 To 24
    Sl = Asc(Mid$(SP, i, 1))
    If (Sl = 114 Or Sl = 109 Or Sl = 105) Then
        If sDB = "" Then sDB = Left$(SP, i - 1)
        If St = 0 Then St = i
    End If
    If (Sl = 32) Or (Sl = 44) Then
        If St > 0 Then
            If f5 = 0 Then
                f5 = RetFlag(Mid$(SP, St, i - St))
            ElseIf f6 = 0 Then
                f6 = RetFlag(Mid$(SP, St, i - St))
            ElseIf f7 = 0 Then
                f7 = RetFlag(Mid$(SP, St, i - St))
            ElseIf f8 = 0 Then
                f8 = RetFlag(Mid$(SP, St, i - St))
            End If
            St = 0
            j = i
        End If
    End If
Next i
sEN = Trim$(Mid$(SP, j, 24))
If sDB = "" Then
    sDB = sEN: sEN = ""
End If

StrToIv = "AddAOC TblASM_OPCODE(" & Cont & ")" & V & FOC & V & OpLen & V & f1 & V & f2 & V & f3 & V & f4 & V & f5 & V & f6 & V & f7 & V & f8 & V & G & sDB & G & V & G & sEN & G
End Function

Private Function RetFlag(inString As String) As Byte
'renvoi le numéro-flag a partir du texte trouvé
Dim i
For i = 0 To 58
    If FlagOPC(i).sTxt = inString Then
        RetFlag = FlagOPC(i).bFlag
        Exit Function
    End If
Next i
End Function

Private Sub Form_Load()
'ne débouche pas les siphons de douches
    InitFlag
End Sub
