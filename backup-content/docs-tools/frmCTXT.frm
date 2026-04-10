VERSION 5.00
Begin VB.Form frmCTXT 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'MSVBM60_API.TXT convertion : par proger pour vb6 analyseur

Private Sub Form_Load()
Dim InTpStr As String
Dim OutStr As String
Dim Adr As String, Ord As String, Nm As String, Rl As String
Dim V As String * 1
V = ","
Rl = "undef"

Open "MSVBM60_API.txt" For Input Access Read As #1
Open "VB60_APIDEF.txt" For Output As #2
Do Until EOF(1)
Line Input #1, InTpStr
If Trim$(InTpStr) <> "" Then
    Adr = Mid$(InTpStr, 7, 8)
    Ord = Trim$(Mid$(InTpStr, 20, 4))
    Nm = Trim$(Mid$(InTpStr, 38))
    OutStr = Adr & V & Ord & V & Nm & V & Rl
    Print #2, OutStr
End If
Loop
Close

End Sub
