Attribute VB_Name = "Globais"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public TLIa As TLIApplication
Public Interface As InterfaceInfo
Public vtInfo As VarTypeInfo
Public mObjectTarget As Object

Public Type UDT_CORD
    X As Single
    Y As Single
End Type

Public Type UDT_LINE_CORD
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
End Type

Public Enum sSENTIDO
    Esquerda = 0
    Acima = 1
    Direita = 2
    Abaixo = 3
End Enum

Public Enum VarEnum
    vtInteger = 2
    vtDouble = 5
    vtSingle = 4
    vtLong = 3
    vtVariant = 12  'Unsuported
    vtString = 8
    vtVoid = 24     'Unsuported "Metodo"
    vtByte = 17
    vtBoolean = 11
    vtObject = 9    'Unsuported
    vtEnum = 0
End Enum

Public Const m_def_Layer = 1

Public Enum TipoSave
    Normal = 0      'Access Database
    Componente = 1  '
    HPLG2 = 2       'PCL languange
    Imagem = 3      '
    NovoNome = 4    '
    Copia = 5       '
    Impressão = 6   'Temp File to Print
End Enum

Public Enum TipoImprimir
    WinDefPrinter = 0
    Ploterdevice = 1
End Enum
'

Public Function CheckValidDataType(vtTeste As Integer) As Boolean
        
    
        'filtrando o s tipos válidos
        Select Case vtTeste
            Case Is = VarEnum.vtBoolean
                CheckValidDataType = True
            Case Is = VarEnum.vtByte
                CheckValidDataType = True
            Case Is = VarEnum.vtInteger
                CheckValidDataType = True
            Case Is = VarEnum.vtLong
                CheckValidDataType = True
            Case Is = VarEnum.vtSingle
                CheckValidDataType = True
            Case Is = VarEnum.vtDouble
                CheckValidDataType = True
            Case Is = VarEnum.vtString
                CheckValidDataType = True
            Case Is = VarEnum.vtEnum
                CheckValidDataType = True
            Case Else
                ' unsuported
                CheckValidDataType = False
        End Select
    

End Function

Public Function Arredonda(sngValue As Single) As Single

Dim sngMax As Single
Dim sngMin As Single
Dim sngTemp As Single
    
    sngMin = sngValue \ 1
    sngTemp = sngValue - sngMin
    
    If sngTemp >= 0.5 Then
        Arredonda = sngMin + 1
    Else
        Arredonda = sngMin
    End If

End Function

Public Function VerificaAlvoRetangulo(X As Single, Y As Single, Eixo As UDT_LINE_CORD) As Boolean
'
' Verificando se uma cordenada esta dentro de um retangulo
'
    With Eixo
        If X > .X1 Then
            If X < .X2 Then
                If Y > .Y1 Then
                    If Y < .Y2 Then
                        VerificaAlvoRetangulo = True
                        Exit Function
                    End If
                End If
            End If
        End If
        
    End With
    
End Function

Public Function VerificaAlvoTraço(X As Single, Y As Single, Eixo As UDT_LINE_CORD) As Boolean
'
'   Verifica se encontra uma cordenada interpola um traço
'

    'Verificando retas
    With Eixo
        If .X1 = .X2 Then
            'Vertical
            If X = .X1 Then
                If .Y1 < .Y2 Then
                    'Desce
                    If Y > .Y1 Then
                        If Y < .Y2 Then
                            VerificaAlvoTraço = True
                            Exit Function
                        End If
                    End If
                Else
                    'Sobe
                    If Y < .Y1 Then
                        If Y > .Y2 Then
                            VerificaAlvoTraço = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf .Y1 = .Y2 Then
            'Horizontal
            If Y = .Y1 Then
                If .X1 < .X2 Then
                    'Direita
                    If X > .X1 Then
                        If X < .X2 Then
                            VerificaAlvoTraço = True
                            Exit Function
                        End If
                    End If
                Else
                    'Esquerda
                    If X < .X1 Then
                        If X > .X2 Then
                            VerificaAlvoTraço = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Else
            'Diagonal
            If .X1 < .X2 Then
                'Direita
                If .Y1 < .Y2 Then
                    'Desce
                    If X > .X1 Then
                        If X < .X2 Then
                            If Y > .Y1 Then
                                If Y < .Y2 Then
                                    'Dentro da região
                                    If Arredonda((Abs(X - .X1) / Abs(.X2 - .X1)) * 10) = Arredonda((Abs(Y - .Y1) / Abs(.Y2 - .Y1)) * 10) Then
                                        VerificaAlvoTraço = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    'Sobe
                    If X > .X1 Then
                        If X < .X2 Then
                            If Y < .Y1 Then
                                If Y > .Y2 Then
                                    'Dentro da região
                                    If Arredonda((Abs(X - .X1) / Abs(.X2 - .X1)) * 10) = Arredonda((Abs(Y - .Y1) / Abs(.Y2 - .Y1)) * 10) Then
                                        VerificaAlvoTraço = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'Esquerda
                If .Y1 < .Y2 Then
                    'Desce
                    If X < .X1 Then
                        If X > .X2 Then
                            If Y > .Y1 Then
                                If Y < .Y2 Then
                                    'Dentro da região
                                    If Arredonda((Abs(X - .X1) / Abs(.X2 - .X1)) * 10) = Arredonda((Abs(Y - .Y1) / Abs(.Y2 - .Y1)) * 10) Then
                                        VerificaAlvoTraço = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    'Sobe
                    If X < .X1 Then
                        If X > .X2 Then
                            If Y < .Y1 Then
                                If Y > .Y2 Then
                                    'Dentro da região
                                    If Arredonda((Abs(X - .X1) / Abs(.X2 - .X1)) * 10) = Arredonda((Abs(Y - .Y1) / Abs(.Y2 - .Y1)) * 10) Then
                                        VerificaAlvoTraço = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With

End Function

Public Function ChkNull(inptstr) As Variant
    
    If IsNull(inptstr) Then
        ChkNull = ""
    Else
        ChkNull = inptstr
    End If

End Function

Public Sub PrintLog(strNewLine As String)

Dim ff As Integer

    ff = FreeFile
    Open App.Path & "/MilanPCB.log" For Append As #ff
        Print #ff, strNewLine
    Close #ff
    
End Sub

Sub Main()

    frmPCB.Show

End Sub
