Attribute VB_Name = "VbObjAnalyse"
DefLng A-Z
Option Explicit
Option Compare Binary

Private Sub Init_VBCTRL()

'les noms indiqués sont ceux des fichiers .frm aux lignes "BEGIN"
    ReDim vbDEFCTRL(24)
    vbDEFCTRL(1).inID = 269  '01 0D
    vbDEFCTRL(1).cType = "VB.Form"
    vbDEFCTRL(2).inID = 549  '02 25
    vbDEFCTRL(2).cType = "VB.Data"
    vbDEFCTRL(3).inID = 1042 '04 12
    vbDEFCTRL(3).cType = "VB.FileListBox"
    vbDEFCTRL(4).inID = 523  '02 0B
    vbDEFCTRL(4).cType = "VB.Timer"
    vbDEFCTRL(5).inID = 1041 '04 11
    vbDEFCTRL(5).cType = "VB.DirListBox"
    vbDEFCTRL(6).inID = 1040 '04 10
    vbDEFCTRL(6).cType = "VB.DriveListBox"
    vbDEFCTRL(7).inID = 522  '02 0A
    vbDEFCTRL(7).cType = "VB.VScrollBar"
    vbDEFCTRL(8).inID = 521  '02 09
    vbDEFCTRL(8).cType = "VB.HScrollBar"
    vbDEFCTRL(9).inID = 1032 '04 08
    vbDEFCTRL(9).cType = "VB.ListBox"
    vbDEFCTRL(10).inID = 1287 '05 07
    vbDEFCTRL(10).cType = "VB.ComboBox"
    vbDEFCTRL(11).inID = 262  '01 06
    vbDEFCTRL(11).cType = "VB.OptionButton"
    vbDEFCTRL(12).inID = 261  '01 05
    vbDEFCTRL(12).cType = "VB.CheckBox"
    vbDEFCTRL(13).inID = 260  '01 04
    vbDEFCTRL(13).cType = "VB.CommandButton"
    vbDEFCTRL(14).inID = 259  '01 03
    vbDEFCTRL(14).cType = "VB.Frame"
    vbDEFCTRL(15).inID = 1026 '04 02
    vbDEFCTRL(15).cType = "VB.TextBox"
    vbDEFCTRL(16).inID = 1280 '05 00
    vbDEFCTRL(16).cType = "VB.PictureBox"
    vbDEFCTRL(17).inID = 792  '03 18
    vbDEFCTRL(17).cType = "VB.Image"
    vbDEFCTRL(18).inID = 791  '03 17
    vbDEFCTRL(18).cType = "VB.Line"
    vbDEFCTRL(19).inID = 1046 '04 16
    vbDEFCTRL(19).cType = "VB.Shape"
    vbDEFCTRL(20).inID = 257  '01 01
    vbDEFCTRL(20).cType = "VB.Label"
    vbDEFCTRL(21).inID = 803  '03 23
    vbDEFCTRL(21).cType = "VB.OLE"
    vbDEFCTRL(22).inID = 787  '03 13
    vbDEFCTRL(22).cType = "VB.Menu"
    vbDEFCTRL(23).inID = 276  '01 14
    vbDEFCTRL(23).cType = "VB.MDIForm"
    vbDEFCTRL(24).inID = 6440 '19 28
    vbDEFCTRL(24).cType = "Objet-classe locale ?"

End Sub

Sub ParseControl(ByVal FilePointer As Integer, ByRef FormsDef() As CONTROL_FORM)
'récupère les différents contrôles utilisés dans les forms de l'exe vb
Dim OffsetStart As Long 'début du bloc info contrôle
Dim SizeBlock As Long   'taille du bloc contenant l'info sur les conrôles
Dim NumCtrl As Long     'nombre de contrôles dans le bloc
Dim ObjBlock As Long    'taille d'un segment (définition d'1 contrôle)
Dim NameLen As Integer  'longueur du nom d'un contrôle (sur 2 octets)
Dim bCol As Byte, bNum As Integer  'collection : identifiant, numéro
Dim fID As Integer      'compteur d'ordre d'apparition dans le form
Dim i As Long, j As Long, BlkEnd As Long, k As Long
Dim BugCheck As Byte '(... ben oui ya des truc que je pige pas encore)
Dim BugCheck2 As Integer


    Call Init_VBCTRL    'charge l'identificateur de contrôles

i = 1
For j = 1 To UBound(FormsDef())
    'BOUCLE 1 : form par form
    
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
    fID = &H2F8

    Do
        'BOUCLE 2 : objet par objet
        
Cscan:
        Get #FilePointer, OffsetStart, ObjBlock
        
        If OffsetStart + 4 > BlkEnd Then Exit Do
        If k > NumCtrl Then
            'on dépasse le nombre d'objets théorique... dans la pratique on dépasse souvent.
        End If
                
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
            exeVB_CONTROL(i).offset = OffsetStart
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
                
            exeVB_CONTROL(i).frmID = fID + bCol * 4

            OffsetStart = OffsetStart + ObjBlock + 1
            i = i + 1: k = k + 1
            
        ElseIf ObjBlock = -1 Then
            'aaakkk :( bug impossible si on ouvre un exe vb6!!!
            Stop
        Else
            
            ReDim Preserve exeVB_CONTROL(1 To i)
            exeVB_CONTROL(i).offset = OffsetStart
            Get #FilePointer, OffsetStart + 4, bCol
            Get #FilePointer, OffsetStart + 5, NameLen
            exeVB_CONTROL(i).sName = ScanString(FilePointer, OffsetStart + 7)
            If exeVB_CONTROL(i).sName = "menu_copy" Then Stop
            Get #FilePointer, OffsetStart + 8 + NameLen, exeVB_CONTROL(i).id
                If (exeVB_CONTROL(i).id And 255) = 255 Then
                    'il s'agit d'un contrôle "externe" a vb (ocx...)
                    exeVB_CONTROL(i).sType = ScanString(FilePointer, OffsetStart + 11 + NameLen)
                Else
                    'contrôle vb interne
                    exeVB_CONTROL(i).sType = Get_VBCTRL(exeVB_CONTROL(i).id)
                End If

            exeVB_CONTROL(i).frmID = fID + bCol * 4
            
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
    
        Offs = tblVBCTRL(i).offset
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
        tblVBCTRL(i).LenTr = iSize
        
        
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
                Get #FilePointer, Offs, gpInt
                'Menu invisible : astuce de urgo (VBFrance : auteur 19905)
                    If gpInt = 768 Then
                       exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 4)
                    Else
                       exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                    End If
                    'code original
                    'exeVB_CTRL_PRP(i).sCaption = ScanString(FilePointer, Offs + 2)
                
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


Function Get_VBCTRL(inIDent As Integer) As String
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
