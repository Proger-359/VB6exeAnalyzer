VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualiseur de processus"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6390
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Infos module"
      Height          =   1935
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
      Begin VB.CommandButton Command4 
         Caption         =   "Dumper"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Adresse mémoire module :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Espace mémoire occupé :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Module IDentifier :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "Liste"
      Top             =   65
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Processus :"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Cheat-O-matic!"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dumper"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tuer le processus"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Modules :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ProcessIDs() As Long
Private CurrentPID As Long
Private CurModuleName() As String
Private CurModuleBase() As Long
Private CurModuleSize() As Long
Private CurModuleID() As Long

Dim ProcWork As New MemWork

Private Sub Combo1_Click()
If Combo1.ListIndex = -1 Then Exit Sub
CurrentPID = Combo1.ListIndex + 1
r = ProcWork.GetPIDModule(ProcessIDs(CurrentPID), CurModuleName(), CurModuleID(), CurModuleSize(), CurModuleBase())
List1.Clear
Dim Muz As Long, Mpk As Long, Pus As Long
s = ProcWork.GetPIDMemInfo(ProcessIDs(CurrentPID), Muz, Mpk, Pus)
Label7.Caption = "Mémoire occupé : " & Muz
Label8.Caption = "Pic d'occupation : " & Mpk
Label10.Caption = "PID : " & ProcessIDs(CurrentPID)
If (r Or s) = -1 Then Exit Sub  'controle d'erreur
For i = 1 To UBound(CurModuleName())
    List1.AddItem ProcWork.LastFileName(CurModuleName(i))
Next i
Label1.Caption = "Modules (" & UBound(CurModuleName()) & ") :"
End Sub

Private Sub Command1_Click()
If CurrentPID <= 0 Then Exit Sub
r = ProcWork.KillPID(ProcessIDs(CurrentPID))
If r = 0 Then
    Call Form_Activate
    List1.Clear
    Text1.Text = ""
    Label10.Caption = "PID : effacé"
Else
    MsgBox "Impossible de tuer le processus " & Combo1.Text & " . Ce programme est peut-être un service.", vbExclamation + vbOKCancel, "Kill Process"
End If
End Sub

Private Sub Command2_Click()
Dim TempDump() As Byte
If CurrentPID = 0 Then Exit Sub
d = ProcWork.DumpValPID(ProcessIDs(CurrentPID), TempDump())
If d = -1 Or ProcessIDs(CurrentPID) = 0 Then
    MsgBox "Dump impossible! - le programme visé est peut-être un service.", vbExclamation + vbOKCancel, "Dumper"
    Exit Sub
End If
Dim HDV As New FormHex
HDV.StartMe TempDump()
End Sub

Private Sub Command3_Click()
Dim UrC As New FormOMat
If CurrentPID = 0 Then Exit Sub
FormOMat.IniCheating ProcessIDs(CurrentPID)
FormOMat.Caption = "Modifications de " & Combo1.Text
End Sub

Private Sub Command4_Click()
If List1.ListIndex < 0 Then Exit Sub
Dim TempDump() As Byte
d = ProcWork.DumpModulePID(ProcessIDs(CurrentPID), TempDump(), CurModuleBase(List1.ListIndex + 1), CurModuleSize(List1.ListIndex + 1))
If d = -1 Then
    MsgBox "Impossible de dumper ce module", vbOKOnly
Else
    Dim HDV As New FormHex
    HDV.StartMe TempDump()
    HDV.Caption = "Hex : dump de " & CurModuleName(List1.ListIndex + 1)
End If
End Sub

Private Sub Form_Activate()
Call ProcWork.GetActiveProcess(ProcessIDs())
Combo1.Clear
For i = 1 To UBound(ProcessIDs())
    t$ = ProcWork.LastFileName(ProcWork.GetPIDName(ProcessIDs(i)))
    Combo1.AddItem t$
Next i
End Sub

Private Sub List1_Click()
If List1.ListIndex = -1 Then Exit Sub
Text1.Text = CurModuleName(List1.ListIndex + 1)
Label5.Caption = Right("00000000" & Hex(CurModuleSize(List1.ListIndex + 1)), 8) & "h"
Label6.Caption = Right("00000000" & Hex(CurModuleBase(List1.ListIndex + 1)), 8) & "h"
Label9.Caption = Right("00000000" & Hex(CurModuleID(List1.ListIndex + 1)), 8) & "h"
End Sub

