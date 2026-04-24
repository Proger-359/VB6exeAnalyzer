Attribute VB_Name = "VbObjAnalyse"
DefLng A-Z
Option Explicit
Option Compare Binary


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
