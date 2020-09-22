VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl Board 
   BackColor       =   &H00404040&
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   ScaleHeight     =   6900
   ScaleWidth      =   7845
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2160
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line lnSelec 
      BorderColor     =   &H00C0C0FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   720
      X2              =   720
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Line lnSelec 
      BorderColor     =   &H00C0C0FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   600
      X2              =   600
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Line lnSelec 
      BorderColor     =   &H00C0C0FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Line lnSelec 
      BorderColor     =   &H00C0C0FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Shape ShConector 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Index           =   0
      Left            =   960
      Shape           =   1  'Square
      Top             =   960
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Line ElLine 
      Visible         =   0   'False
      X1              =   1080
      X2              =   2100
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "Principal"
      Begin VB.Menu mnuPrincipalExcluir 
         Caption         =   "Excluir"
      End
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtElastic
    X As Single
    Y As Single
    Enabled As Boolean
End Type
Private Elastic As udtElastic

Private Type udtIgnoreFoco
    intCount As Integer
    X As Single
    Y As Single
End Type
Private IgnoreFoco As udtIgnoreFoco

Private mHabDesagrupar      As Boolean
Private mHabAgrupar         As Boolean
Private mHabilitaGirar      As Boolean

Private SelectedArea                        As UDT_LINE_CORD

Private WithEvents mHS                       As HScrollBar
Attribute mHS.VB_VarHelpID = -1
Private WithEvents mVS                       As VScrollBar
Attribute mVS.VB_VarHelpID = -1
Private WithEvents LoadCompoForm            As frmLoadComponentes
Attribute LoadCompoForm.VB_VarHelpID = -1
Private WithEvents SelectCompoForm          As frmSelectComponente
Attribute SelectCompoForm.VB_VarHelpID = -1
Private WithEvents mToolStandart            As Toolbar
Attribute mToolStandart.VB_VarHelpID = -1
Private WithEvents mToolStandartObjects     As Toolbar
Attribute mToolStandartObjects.VB_VarHelpID = -1
Private mPropInterface                      As ObjectPropIII
Private mBarraStatus                        As StatusBar

Private colPCBCopy                          As New Collection
Private colComponenteCopy                   As New Collection
Private colGrupoCopy                        As New Collection
Private colPCB                              As New Collection       'objetos do PCB
Private colSelec                            As New Collection       'objetos selecionados
Private colGrupo                            As New Collection
Private colComponente                       As New Collection
Private colCompDisponiveis                  As New Collection
Private SaveOpen                            As New clsSaveOpen
Private InputBoxMkCom                       As New clsImputMakeComponente

Private bntTraço                            As Boolean
Private Grupo                               As clsGrupo
Private Componente                          As clsComponente
Private Objetos                             As clsObjetos

Public WithEvents Prop                      As clsBoardProp
Attribute Prop.VB_VarHelpID = -1
Private WithEvents Ilha                     As clsIlha
Attribute Ilha.VB_VarHelpID = -1
Private WithEvents Traço                    As clsTraço
Attribute Traço.VB_VarHelpID = -1
Private WithEvents Corner                   As clsCorner
Attribute Corner.VB_VarHelpID = -1



Public Event CaptionChange(NovoCaption As String)
'
Private Sub Corner_GotFocus()

    mPropInterface.ObjectTarget = Corner
    colSelec.Add Corner.ID, "C" & Corner.ID
    
End Sub

Private Sub Corner_LostFocus()

    colSelec.Remove ("C" & Corner.ID)
    
End Sub


Private Sub Corner_Paint()
'
'   Desenha o Corner
'
Dim X As Single
Dim Y As Single

    UserControl.ForeColor = vbWhite
    If VerificaQuadro(Corner.X, Corner.Y) = True Then
        
        X = Corner.X - Prop.X1
        Y = Corner.Y - Prop.Y1
   
        UserControl.DrawWidth = 1
        UserControl.PSet (X, Y)
        
        Corner.iShape.Left = X - 25
        Corner.iShape.Top = Y - 25
   
'        picBuffer.DrawWidth = 1
'        picBuffer.PSet (X, Y)
'        UserControl.PaintPicture picBuffer.Image, 0, 0
    
    End If
    
End Sub


Private Sub Ilha_GotFocus()
    
    mPropInterface.ObjectTarget = Ilha
    colSelec.Add Ilha.Corner, "I" & Ilha.Corner
    
End Sub

Private Sub Ilha_LostFocus()

    colSelec.Remove ("I" & Ilha.Corner)
    
End Sub

Private Sub Ilha_Paint()
'
'   Desenha Ilha
'
Dim X As Single
Dim Y As Single
Dim Rad As Integer
Dim St As Integer
Dim Et As Integer
Dim lngCor As Long

    Set Corner = colPCB.Item("Corner-" & Ilha.Corner)

    If VerificaQuadro(Corner.X, Corner.Y) = True Then
    
        X = Corner.X - Prop.X1
        Y = Corner.Y - Prop.Y1
        
        St = Ilha.IlhaFuro / 2
        Et = Ilha.IlhaLargura / 2
        
        If Ilha.HasFocus = True Then
            lngCor = QBColor(Prop.GetLayerColor(Ilha.Layer) + 8)
        Else
            lngCor = QBColor(Prop.GetLayerColor(Ilha.Layer))
        End If
        
        UserControl.DrawWidth = 1
        UserControl.ForeColor = lngCor

        For Rad = St To Et
            UserControl.Circle (X, Y), Rad, lngCor
        Next
        
        'desenha o conector
        Corner.Refresh
    
'        picBuffer.DrawWidth = 1
'        picBuffer.ForeColor = lngCor
'        For Rad = St To Et
'            picBuffer.Circle (X, Y), Rad, lngCor
'        Next
'        UserControl.PaintPicture picBuffer.Image, 0, 0
        
    End If
    
End Sub

Private Sub LoadCompoForm_Carregados(colCarregados As Collection)
    
    Set colCompDisponiveis = colCarregados
    
End Sub

Private Sub mHS_Change()

    With Prop
        .X1 = mHS.Value
        .X2 = .X1 + UserControl.ScaleWidth
    End With
    Call RedrawAll
    
End Sub

Private Sub mnuPrincipalExcluir_Click()

    Call Excluir
    
End Sub

Private Sub mToolStandart_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
    
        Case Is = "Abrir"
            Call Abrir
        
        Case Is = "Salvar"
            Call Salvar(Normal)
        
        Case Is = "Imprimir"
            
        
        Case Is = "Agrupar"
            Call Agrupar
        
        Case Is = "Desagrupar"
            Call Desagrupar
        
        Case Is = "Componentes"
            Call ShowLoadComponentes
        
        Case Is = "GirarDireita"
            Call Rotate(Direita)
        
        Case Is = "GirarEsquerda"
            Call Rotate(Esquerda)
        
        Case Is = "Copiar"
            Call Copiar
        
        Case Is = "Recortar"
            Call Copiar
            Call Excluir
        
        Case Is = "Colar"
            Call Colar
    
    End Select
            
End Sub

Private Sub mToolStandartObjects_ButtonClick(ByVal Button As MSComctlLib.Button)
'
' Captura o precionamento dos botões standart objects
'

Dim dsButton As MSComctlLib.Button

    'cancelamento de operações
    bntTraço = False

    For Each dsButton In mToolStandartObjects.Buttons
        dsButton.Value = tbrUnpressed
    Next
    
    mToolStandartObjects.Buttons.Item(Button.Key).Value = tbrPressed
    mToolStandartObjects.Refresh
    
    Select Case Button.Key
        Case Is = "Arco"
            Prop.Ferramenta = fArco
            
        Case Is = "Componente"
            Prop.Ferramenta = fComponente
            
        Case Is = "Corte"
            Prop.Ferramenta = fCorte
            
        Case Is = "Elipse"
            Prop.Ferramenta = fElipse
            
        Case Is = "Ilha"
            Prop.Ferramenta = fIlha
            
        Case Is = "Retangulo"
            Prop.Ferramenta = fRetangulo
            
        Case Is = "Seleção"
            Prop.Ferramenta = fSeleção
            
        Case Is = "Traço"
            Prop.Ferramenta = fTraço
            
    End Select
    
    UserControl.SetFocus
    
End Sub

Private Sub mVS_Change()

    With Prop
        .Y1 = mVS.Value
        .Y2 = .Y1 + UserControl.ScaleHeight
    End With
    Call RedrawAll
    
End Sub

Private Sub Prop_Nome(NewName As String)
'
'   Apos o Salvamento
'
    RaiseEvent CaptionChange(NewName & " - " & LoadResString(151))
    
End Sub

Private Sub SelectCompoForm_Selecionado(mComponente As clsComponente)
    
    Set Componente = New clsComponente
    With Componente
        .ID = mComponente.ID
        .DBID = mComponente.DBID
        .Nome = mComponente.Nome
    End With
    
End Sub

Private Sub Traço_GotFocus()

    mPropInterface.ObjectTarget = Traço
    colSelec.Add Traço.StartCorner, "Ts" & Traço.StartCorner
    colSelec.Add Traço.EndCorner, "Te" & Traço.EndCorner
    
End Sub


Private Sub Traço_LostFocus()

    colSelec.Remove ("Ts" & Traço.StartCorner)
    colSelec.Remove ("Te" & Traço.EndCorner)
    
End Sub


Private Sub Traço_Paint()
'
'   Desenha um Traço em uma coordenada
'

Dim X1 As Single
Dim X2 As Single
Dim Y1 As Single
Dim Y2 As Single
Dim lngCor As Long
Dim bntInArea As Boolean

    bntInArea = False
    ' desenha o traço
    Set Corner = colPCB.Item("Corner-" & Traço.StartCorner)
    
    If VerificaQuadro(Corner.X, Corner.Y) = True Then
        bntInArea = True
    End If
    
    X1 = Corner.X - Prop.X1
    Y1 = Corner.Y - Prop.Y1
    
    Set Corner = colPCB.Item("Corner-" & Traço.EndCorner)
    
    If VerificaQuadro(Corner.X, Corner.Y) = True Then
        bntInArea = True
    End If
    
    X2 = Corner.X - Prop.X1
    Y2 = Corner.Y - Prop.Y1
    
    If bntInArea = True Then
    
        If Traço.HasFocus = True Then
            lngCor = QBColor(Prop.GetLayerColor(Traço.Layer) + 8)
        Else
            lngCor = QBColor(Prop.GetLayerColor(Traço.Layer))
        End If
        
        UserControl.DrawWidth = Traço.Largura
        UserControl.ForeColor = lngCor
        UserControl.Line (X1, Y1)-(X2, Y2)

        ' desenha os corner do traço
        UserControl.DrawWidth = 1
        UserControl.ForeColor = vbWhite
        UserControl.PSet (X1, Y1)
        UserControl.PSet (X2, Y2)
        
    End If
        
End Sub

Private Sub UserControl_Initialize()
    
    bntTraço = False
    Set Objetos = New clsObjetos
    Set Prop = New clsBoardProp
    Prop.Ferramenta = fSeleção
    lnSelec(0).Visible = False
    lnSelec(1).Visible = False
    lnSelec(2).Visible = False
    lnSelec(3).Visible = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim dsButton As MSComctlLib.Button

    ' corrigindo os eixos
    X = (Arredonda(X / 64)) * 64
    Y = (Arredonda(Y / 64)) * 64
    
    X = X + Prop.X1
    Y = Y + Prop.Y1
    
    ' se houver uma seleção
    SelectedArea.X1 = X
    SelectedArea.Y1 = Y
    
    HabAgrupar = False
    HabDesagrupar = False
    HabilitaGirar = False
    
    If Button = vbRightButton Then
        
            ' Cancela as operações
            Elastic.Enabled = False
            ElLine.Visible = False
            bntTraço = False
            
            'ativa a ferramenta de seleção
            For Each dsButton In mToolStandartObjects.Buttons
                dsButton.Value = tbrUnpressed
            Next
            
            ' altera o status do botão
            mToolStandartObjects.Buttons("Seleção").Value = tbrPressed
            Prop.Ferramenta = fSeleção
    
            ' ativa o foco
            Call AdicionarFocoObjArea(X, Y)
                        
    ElseIf Button = vbLeftButton Then
    
        ' sera aplicada uma ferramernta
        
        Select Case Prop.Ferramenta
        
            Case Is = fSeleção
                ' ativa o foco no objeto
                Call AdicionarFocoObjArea(X, Y)
                
            Case Is = fTraço
                                
                If bntTraço = True Then
                    Call AddTraço(X, Y)
                    mPropInterface.ObjectTarget = Traço
                End If
                ' inicia uma linha elastica
                With Elastic
                    .Enabled = True
                    .X = X
                    .Y = Y
                End With
                ' prepara a linha
                With ElLine
                    .BorderWidth = Prop.LarguraTraço
                    .BorderColor = QBColor(Prop.GetActiveLayerColor)
                End With
                bntTraço = True
                
            Case Is = fIlha
            
                Call AddIlha(X, Y)
                mPropInterface.ObjectTarget = Ilha
                                
            Case Is = fComponente
                
                'Mostra o formulario de escolha de componente
                Set SelectCompoForm = New frmSelectComponente
                SelectCompoForm.ComponentesC = colCompDisponiveis
                Set Componente = Nothing
                SelectCompoForm.Show 1
                
                If Not (Componente Is Nothing) Then
                    Call AddComponente(Componente, X, Y)
                End If
                mPropInterface.ObjectTarget = Componente
                
            Case Is = fCorte
                '
            Case Is = fElipse
                '
            Case Is = fArco
                '
            Case Is = fRetangulo
                '
        End Select
        
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
'

On Error GoTo errTrat

Dim vntCorner As Variant
Dim bntMoveu As Boolean

    ' corrigindo os eixos
    X = (Arredonda(X / 64)) * 64
    Y = (Arredonda(Y / 64)) * 64
    
    X = X + Prop.X1
    Y = Y + Prop.Y1
    
    SelectedArea.X2 = X
    SelectedArea.Y2 = Y
    
    If Button = vbLeftButton Then
    
        With SelectedArea
        
            If TypeOf mPropInterface.ObjectTarget Is clsBoardProp Then
                Call GerarLinhasSeleção
            Else
                
                bntMoveu = False
                
                For Each vntCorner In colSelec
                    bntMoveu = True
                    Set Corner = colPCB("Corner-" & CStr(vntCorner))
                    Corner.X = Corner.X + (.X2 - .X1)
                    Corner.Y = Corner.Y + (.Y2 - .Y1)
                Next
                
                If bntMoveu = True Then
                    Call RedrawAll
                    .X1 = .X2
                    .Y1 = .Y2
                End If
                
            End If
            
        End With
        
    End If
    
    With Elastic
        
        If .Enabled = True Then
            ElLine.Visible = True
            ElLine.X1 = .X - Prop.X1
            ElLine.Y1 = .Y - Prop.Y1
            ElLine.X2 = X - Prop.X1
            ElLine.Y2 = Y - Prop.Y1
        End If
        
    End With
    
errTrat:
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' corrigindo os eixos
    X = (Arredonda(X / 64)) * 64
    Y = (Arredonda(Y / 64)) * 64
    
    X = X + Prop.X1
    Y = Y + Prop.Y1
    
    SelectedArea.X2 = X
    SelectedArea.Y2 = Y
    
    lnSelec(0).Visible = False
    lnSelec(1).Visible = False
    lnSelec(2).Visible = False
    lnSelec(3).Visible = False
    
    If Prop.Ferramenta = fSeleção Then
    
        If Button = vbRightButton Then
            
            If Not (TypeOf mPropInterface.ObjectTarget Is clsBoardProp) Then
                PopupMenu mnuPrincipal
            End If
            
        Else
            
            If TypeOf mPropInterface.ObjectTarget Is clsBoardProp Then
                ' O usuario pretente selecionar um grupo de objetos?
                With SelectedArea
                    If .X1 <> .X2 Or .Y1 <> .Y2 Then    'sim
                        'faz a busca e cantos dos objetos dentro da coleção
                        HabilitaGirar = False
                        HabAgrupar = False
                        HabDesagrupar = False
                        Call VerificaCornersArea(SelectedArea)
                    End If
                End With
            End If
            
        End If
    
    End If

End Sub

Private Sub UserControl_Resize()

    If Not (mHS Is Nothing) Then
        With mHS
            .Max = 1000 - (UserControl.ScaleWidth / 256)
            Prop.X1 = .Value
            Prop.X2 = Prop.X1 + UserControl.ScaleWidth
        End With
    End If
    
    If Not (mVS Is Nothing) Then
        With mVS
            .Max = 1000 - (UserControl.ScaleHeight / 256)
            Prop.Y1 = .Value
            Prop.Y2 = Prop.Y1 + UserControl.ScaleHeight
        End With
    End If
    
'    picBuffer.Width = UserControl.ScaleWidth
'    picBuffer.Height = UserControl.ScaleHeight
    
    Call RedrawAll

End Sub

Private Sub UserControl_Terminate()
    Set Prop = Nothing
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get PropInterface() As ObjectPropIII
    Set PropInterface = mPropInterface
End Property

Public Property Let PropInterface(NewValue As ObjectPropIII)
    
    Set mPropInterface = NewValue
    Prop.Enabled = False
    mPropInterface.ObjectTarget = Prop
    Prop.Enabled = True
    
End Property

Public Property Get VS() As VScrollBar
    Set VS = mVS
End Property

Public Property Let VS(NewValue As VScrollBar)
    
    Set mVS = NewValue
    
    With mVS
        .Max = 1000 - (UserControl.ScaleHeight / 256)
        .Min = 0
        .Value = .Max / 2
        .LargeChange = 10
        .SmallChange = 1
    End With
    
End Property

Public Property Get HS() As HScrollBar
    Set HS = mHS
End Property

Public Property Let HS(NewValue As HScrollBar)
    
    Set mHS = NewValue
    
    With mHS
        .Max = 1000 - (UserControl.ScaleWidth / 256)
        .Min = 0
        .Value = .Max / 2
        .LargeChange = 10
        .SmallChange = 1
    End With
    
End Property

Public Property Get ToolStandart() As Toolbar
    Set ToolStandart = mToolStandart
End Property

Public Property Let ToolStandart(NewValue As Toolbar)
    
    Set mToolStandart = NewValue

End Property

Public Property Get ToolStandartObjects() As Toolbar
    Set ToolStandartObjects = mToolStandart
End Property

Public Property Let ToolStandartObjects(NewValue As Toolbar)
    
    Set mToolStandartObjects = NewValue
    mToolStandartObjects.Buttons("Seleção").Value = tbrPressed
    Prop.Ferramenta = fSeleção
    
End Property

Public Property Get BarraStatus() As StatusBar

    Set BarraStatus = mBarraStatus
    
End Property

Public Property Let BarraStatus(NewValue As StatusBar)
    
    Set mBarraStatus = NewValue
    
End Property

Private Sub DrawGrid()

On Local Error GoTo errTrat

Dim X       As Single
Dim Y       As Single
Dim X1      As Single
Dim Y1      As Single
Dim X2      As Single
Dim Y2      As Single
Dim I       As Integer
Dim stepX   As Single
Dim stepY   As Single

    UserControl.AutoRedraw = True

    If Prop.Grid = True Then
    
        Screen.MousePointer = vbHourglass
        If UserControl.Width <> 0 And UserControl.Height <> 0 Then
            
            UserControl.Cls
            UserControl.DrawWidth = 1
            UserControl.ForeColor = QBColor(Prop.GridColor)
            UserControl.BackColor = QBColor(Prop.BackColor)
            
            stepX = 256 * Prop.Zoom
            stepY = 256 * Prop.Zoom
            X2 = UserControl.ScaleWidth
            Y2 = UserControl.ScaleHeight
            
            If stepX <> 0 And stepY <> 0 Then
                For X1 = stepX To X2 Step stepX
                    For Y1 = stepY To Y2 Step stepY
                        UserControl.PSet (X1, Y1)
                    Next
                    I = I + 1
                Next
            End If
            
        End If

    End If
    
    
errTrat:
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Prop_BackColor(BkColor As Long)
    UserControl.BackColor = BkColor
    Call RedrawAll
End Sub

Private Sub Prop_DrawGrid()
    Call RedrawAll
End Sub

Private Sub Prop_ClearGrid()
    Call RedrawAll
End Sub

Private Function GetCornerInCoordnates(X As Single, Y As Single) As clsCorner
'
'   Procura por um corner
'
Dim vntCorner As Variant

    Set GetCornerInCoordnates = Nothing
    
    For Each vntCorner In colPCB
    
        If TypeOf vntCorner Is clsCorner Then
            If X = vntCorner.X Then
                If Y = vntCorner.Y Then
                    Set GetCornerInCoordnates = vntCorner
                    Exit For
                End If
            End If
        End If
    Next
    
End Function

Private Sub AddIlha(X As Single, Y As Single)
'
'   Adiciona uma Ilha em uma Coordenada
'

    Set Corner = GetCornerInCoordnates(X, Y)
    If Corner Is Nothing Then
        Call AddCorner(X, Y)
    End If
        
    Set Ilha = New clsIlha
        
    With Ilha
        .Corner = Corner.ID
        .ID = Objetos.Ilha
        .IlhaLargura = Prop.IlhaLargura
        .IlhaFuro = Prop.IlhaFuro
        .Layer = Prop.Layer
    End With

    colPCB.Add Ilha, "Ilha-" & Ilha.ID
    
    Ilha.Refresh
    
End Sub

Private Sub AddTraço(X As Single, Y As Single)

    Set Traço = New clsTraço
    
    Traço.ID = Objetos.Traço
    
    ' criando o canto inicial
    Set Corner = GetCornerInCoordnates(Elastic.X, Elastic.Y)
    If Corner Is Nothing Then
        Call AddCorner(Elastic.X, Elastic.Y)
    End If
    '
    Traço.StartCorner = Corner.ID
    '
    ' criando o canto final
    Set Corner = GetCornerInCoordnates(X, Y)
    If Corner Is Nothing Then
        Call AddCorner(X, Y)
    End If
    '
    Traço.EndCorner = Corner.ID
    '
    Traço.Largura = Prop.LarguraTraço
    Traço.Layer = Prop.Layer
    
    ' adicionando a coleção de objetos
    colPCB.Add Traço, "Traço-" & Traço.ID
    
    Traço.Refresh
    
End Sub

Private Sub VerificaCornersArea(Area As UDT_LINE_CORD)

Dim I As Long
Dim G As Long
Dim XY As Single
Dim vntObject As Variant
Dim colFocus As New Collection
Dim colCompo As New Collection
Dim vntCorner As Variant
Dim cNum As Long
Dim cGrp As Long

    'corrigindo
    With Area
    
        If .X1 > .X2 Then
            XY = .X1
            .X1 = .X2
            .X2 = XY
        End If
        
        If .Y1 > .Y2 Then
            XY = .Y1
            .Y1 = .Y2
            .Y2 = XY
        End If
        
        ' Adicionando o foco aos cantos
        For Each vntObject In colPCB
            
            If TypeOf vntObject Is clsCorner Then
            
                Set Corner = vntObject
                
                If Corner.X > .X1 Then
                    If Corner.X < .X2 Then
                        If Corner.Y > .Y1 Then
                            If Corner.Y < .Y2 Then
                            
                                I = I + 1
                                
                                Corner.HasFocus = True
                                
                                If Corner.Grupo <> -1 Then
                                    colFocus.Add Corner
                                End If
                                
                                If Corner.Componente <> -1 Then
                                    colCompo.Add Corner
                                End If
                                                                
                            End If
                        End If
                    End If
                End If
                
            End If
            
        Next
        
    End With
        
    For Each vntCorner In colFocus
        For Each vntObject In colPCB
            If TypeOf vntObject Is clsCorner Then
                If vntCorner.Grupo = vntObject.Grupo Then
                    G = G + 1
                    Set Corner = vntObject
                    Corner.HasFocus = True
                End If
            End If
        Next
    Next
    
    G = 0
    For Each vntCorner In colCompo
        For Each vntObject In colPCB
            If TypeOf vntObject Is clsCorner Then
                If vntCorner.Componente = vntObject.Componente Then
                    G = G + 1
                    Set Corner = vntObject
                    Corner.HasFocus = True
                End If
            End If
        Next
    Next
        
    If I > 0 Then
        HabAgrupar = True
        HabilitaGirar = True
    Else
        HabAgrupar = False
        HabilitaGirar = False
    End If
    
End Sub

Private Function VerificaObjeto(X As Single, Y As Single) As Object
'
' Verifica se existe algum objeto na cordenada X,Y
'
Dim FoundArea As UDT_LINE_CORD
Dim vntObjeto As Variant
Dim intIgnore As Integer

    With IgnoreFoco
        ' para pesquisar na camada de baixo
        If X = .X Then
            If Y = .Y Then
                .intCount = .intCount + 1
            Else
                .intCount = 0
            End If
        Else
            .intCount = 0
        End If
        intIgnore = .intCount
        .X = X
        .Y = Y
    End With
    
    For Each vntObjeto In colPCB
    
        If TypeOf vntObjeto Is clsIlha Then
                                    
            Set Corner = colPCB.Item("Corner-" & vntObjeto.Corner)
            With FoundArea
                ' Retangulo de busca
                .X1 = Corner.X - (vntObjeto.IlhaLargura / 2)
                .Y1 = Corner.Y - (vntObjeto.IlhaLargura / 2)
                .X2 = Corner.X + (vntObjeto.IlhaLargura / 2)
                .Y2 = Corner.Y + (vntObjeto.IlhaLargura / 2)
            End With
            
            ' Verificando
            If VerificaAlvoRetangulo(X, Y, FoundArea) = True Then
                
                If intIgnore = 0 Then
                    Set VerificaObjeto = vntObjeto
                    Exit Function
                Else
                    intIgnore = intIgnore - 1
                End If
                
            End If
        
        ElseIf TypeOf vntObjeto Is clsTraço Then
            
            ' Início do Traço
            Set Corner = colPCB.Item("Corner-" & vntObjeto.StartCorner)
            With FoundArea
                .X1 = Corner.X
                .Y1 = Corner.Y
            End With
            
            ' Término do Traço
            Set Corner = colPCB.Item("Corner-" & vntObjeto.EndCorner)
            With FoundArea
                .X2 = Corner.X
                .Y2 = Corner.Y
            End With
            
            ' Verificando
            If VerificaAlvoTraço(X, Y, FoundArea) = True Then
                
                If intIgnore = 0 Then
                    Set VerificaObjeto = vntObjeto
                    Exit Function
                Else
                    intIgnore = intIgnore - 1
                End If
                
            End If
            
        ElseIf TypeOf vntObjeto Is clsCorner Then
            
            If vntObjeto.X = X Then
                If vntObjeto.Y = Y Then
                    
                    If intIgnore = 0 Then
                        Set VerificaObjeto = vntObjeto
                        Exit Function
                    Else
                        intIgnore = intIgnore - 1
                    End If
                    
                End If
            End If
            
        End If
    Next

End Function

Private Sub RedrawAll()
'
' Redesenha os objetos na tela
'
Dim vntObject As Variant

    UserControl.Cls
    Call DrawGrid
    
    For Each vntObject In colPCB
    
        If TypeOf vntObject Is clsIlha Then
            Set Ilha = vntObject
            Ilha.Refresh
        ElseIf TypeOf vntObject Is clsTraço Then
            Set Traço = vntObject
            Traço.Refresh
        ElseIf TypeOf vntObject Is clsCorner Then
            Set Corner = vntObject
            Corner.Refresh
        End If
    
    Next

End Sub

Private Sub RemoverFoco(objExeto As String, objID As Long)

Dim vntObjeto As Variant
           
    For Each vntObjeto In colPCB
    
        If TypeName(vntObjeto) = "clsIlha" Then
        
            Set Ilha = vntObjeto
            If TypeName(Ilha) = objExeto Then
                If Ilha.ID <> objID Then
                    Ilha.HasFocus = False
                End If
            Else
                Ilha.HasFocus = False
            End If
            
        ElseIf TypeName(vntObjeto) = "clsTraço" Then
        
            Set Traço = vntObjeto
            If TypeName(Traço) = objExeto Then
                If Traço.ID <> objID Then
                    Traço.HasFocus = False
                End If
            Else
                Traço.HasFocus = False
            End If
            
        ElseIf TypeName(vntObjeto) = "clsCorner" Then
        
            Set Corner = vntObjeto
            If TypeName(Corner) = objExeto Then
                If Corner.ID <> objID Then
                    Corner.HasFocus = False
                End If
            Else
                Corner.HasFocus = False
            End If
            
        End If
        
    Next
        
End Sub

Private Sub AdicionarFocoObjArea(X As Single, Y As Single)
'
'   Verifica se existe algum Objeto
'
Dim objBusca As Object
Dim vntObject As Variant

    ' devolve o objeto na cordenada
    ' se for a mesma cama, pesquisa na camada inferior
    Set objBusca = VerificaObjeto(X, Y)
    
    If Not (objBusca Is Nothing) Then
    
        If objBusca.HasFocus = False Then
            
            If objBusca.Grupo = -1 Then
            
                If objBusca.Componente = -1 Then
                        
                    If TypeOf objBusca Is clsIlha Then
                    
                        Set Ilha = objBusca
                        Ilha.HasFocus = True
                        Call RemoverFoco(TypeName(Ilha), objBusca.ID)
                        
                    ElseIf TypeOf objBusca Is clsTraço Then
                    
                        Set Traço = objBusca
                        Traço.HasFocus = True
                        Call RemoverFoco(TypeName(Traço), objBusca.ID)
                        
                    ElseIf TypeOf objBusca Is clsCorner Then
                    
                        Set Corner = objBusca
                        Corner.HasFocus = True
                        Call RemoverFoco(TypeName(Corner), objBusca.ID)
                        
                    End If
                
                Else
                
                    RemoverFoco "", 0
                    
                End If
            
            Else
                
                RemoverFoco "", 0
            
            End If
            
        End If
        
        If objBusca.Componente <> -1 Then
            For Each vntObject In colPCB
                If vntObject.Componente = objBusca.Componente Then
                    If TypeName(vntObject) = "clsCorner" Then
                        Set Corner = vntObject
                        Corner.HasFocus = True
                    End If
                End If
            Next
            
            Set Componente = colComponente("Componente-" & objBusca.Componente)
            
            mPropInterface.ObjectTarget = Componente
            
            HabilitaGirar = True
            
        End If
        
        'se o objeto buscado pertence a um grupo seleciona todo o grupo
        If objBusca.Grupo <> -1 Then
            For Each vntObject In colPCB
                If vntObject.Grupo = objBusca.Grupo Then
                    If TypeName(vntObject) = "clsCorner" Then
                        Set Corner = vntObject
                        Corner.HasFocus = True
                    End If
                End If
            Next
            
            Set Grupo = colGrupo("Grupo-" & objBusca.Grupo)
            
            HabDesagrupar = True
            HabilitaGirar = True
            
            mPropInterface.ObjectTarget = Grupo
            
        End If
    
    Else
    
        Prop.Enabled = False
        mPropInterface.ObjectTarget = Prop
        Prop.Enabled = True
        Prop.HasFocus = True
        Call RemoverFoco("", 0)
        
    End If
    
End Sub

Private Sub AddCorner(X As Single, Y As Single)
'
'   Adiciona um Conector em uma Coordenada
'
Dim ub As Integer

    Set Corner = New clsCorner
    With Corner
        .ID = Objetos.Corner
        .X = X
        .Y = Y
        .Layer = Prop.Layer
    End With
    
    ub = ShConector.Count
    Load ShConector(ub)
    Corner.iShape = ShConector(ub)
    
    ' adicionando a coleção de objetos
    colPCB.Add Corner, "Corner-" & Corner.ID
    Corner.Refresh
    
    RaiseEvent CaptionChange(Prop.Nome)
    
End Sub

Private Sub GerarLinhasSeleção()

    With SelectedArea
        lnSelec(0).Visible = True
        lnSelec(0).X1 = .X1 - Prop.X1
        lnSelec(0).X2 = .X1 - Prop.X1   'para baixo
        lnSelec(0).Y1 = .Y1 - Prop.Y1
        lnSelec(0).Y2 = .Y2 - Prop.Y1
        '
        lnSelec(1).Visible = True
        lnSelec(1).X1 = .X1 - Prop.X1
        lnSelec(1).X2 = .X2 - Prop.X1   'para direita
        lnSelec(1).Y1 = .Y2 - Prop.Y1
        lnSelec(1).Y2 = .Y2 - Prop.Y1
        '
        lnSelec(2).Visible = True
        lnSelec(2).X1 = .X2 - Prop.X1
        lnSelec(2).X2 = .X2 - Prop.X1   'para cima
        lnSelec(2).Y1 = .Y2 - Prop.Y1
        lnSelec(2).Y2 = .Y1 - Prop.Y1
        '
        lnSelec(3).Visible = True
        lnSelec(3).X1 = .X2 - Prop.X1
        lnSelec(3).X2 = .X1 - Prop.X1   'para esquerda
        lnSelec(3).Y1 = .Y1 - Prop.Y1
        lnSelec(3).Y2 = .Y1 - Prop.Y1
    End With
        
End Sub

Private Function VerObjCorSelec() As Collection
'
'   Seleciona Objetos cujos Cornes Estejam Selecionados
'

Dim vntObject As Variant
Dim tmpColection As New Collection
        
    For Each vntObject In colPCB

        If TypeOf vntObject Is clsIlha Then

            ' Se o corner da ilha estiver marcado
            Set Corner = colPCB("Corner-" & vntObject.Corner)
            If Corner.HasFocus = True Then
                Set Ilha = vntObject
                tmpColection.Add Ilha
            End If

        ElseIf TypeOf vntObject Is clsTraço Then

            ' Se ambos os corners do traço estão marcados
            Set Corner = colPCB("Corner-" & vntObject.StartCorner)
            If Corner.HasFocus = True Then
                Set Corner = colPCB("Corner-" & vntObject.EndCorner)
                If Corner.HasFocus = True Then
                    Set Traço = vntObject
                    tmpColection.Add Traço
                End If
            End If

        End If

    Next

    Set VerObjCorSelec = tmpColection
    
End Function

Private Function VerObjCorNotSelec() As Collection
'
'   Seleciona Objetos cujos Cornes Estejam Selecionados
'

Dim vntObject As Variant
Dim stCorner As clsCorner
Dim edCorner As clsCorner
Dim tmpColection As New Collection
        
    For Each vntObject In colPCB
        
        If vntObject.HasFocus = False Then
            
            If TypeOf vntObject Is clsIlha Then
    
                Set Corner = colPCB("Corner-" & vntObject.Corner)
                If Corner.HasFocus = False Then
                    Set Ilha = vntObject
                    tmpColection.Add Ilha
                End If
    
            ElseIf TypeOf vntObject Is clsTraço Then
    
                Set stCorner = colPCB("Corner-" & vntObject.StartCorner)
                Set edCorner = colPCB("Corner-" & vntObject.EndCorner)
                'se algum dos corner não tiver foco
                If stCorner.HasFocus = False Then
                    Set Traço = vntObject
                    tmpColection.Add Traço
                Else
                    If edCorner.HasFocus = False Then
                        Set Traço = vntObject
                        tmpColection.Add Traço
                    End If
                End If
    
            End If
        
        End If
        
    Next

    Set VerObjCorNotSelec = tmpColection
    
End Function

Public Sub Abrir(Optional strFile As String)

On Error GoTo ErrHandler

Dim vntObject   As Variant
Dim vntGrupo    As Variant
Dim vntCompo    As Variant
Dim ub          As Long
Dim ubC         As Long
Dim ubI         As Long
Dim ubT         As Long
Dim ubG         As Long
Dim ubP         As Long
Dim tmpCol      As Collection

    If strFile = "" Then
    
        CD1.CancelError = True
        CD1.Flags = cdlOFNHideReadOnly
        CD1.Filter = "All Files (*.*)|*.*|Data Files(*.mpcb)|*.mpcb"
        CD1.FilterIndex = 2
        CD1.ShowOpen
        
        SaveOpen.FileName = CD1.FileName
    
    Else
        
        SaveOpen.FileName = strFile
    
    End If
    
    Set tmpCol = SaveOpen.Abrir
    
    Set colPCB = tmpCol("PCB")
    Set colGrupo = tmpCol("GRUPO")
    Set colComponente = tmpCol("COMPO")
    
    ' Adicionando os shapes
    For Each vntObject In colPCB
        If TypeOf vntObject Is clsCorner Then
            Set Corner = vntObject
            ub = ShConector.Count
            Load ShConector(ub)
            Corner.AddShape ShConector(ub)
        End If
    Next
    
    'Ajustando os coantadores
    ubC = 0
    ubI = 0
    ubT = 0
    ubG = 0
    ubP = 0
    For Each vntObject In colPCB
        If TypeOf vntObject Is clsCorner Then
            If vntObject.ID >= ubC Then
                ubC = vntObject.ID + 1
            End If
        ElseIf TypeOf vntObject Is clsIlha Then
            If vntObject.ID >= ubI Then
                ubI = vntObject.ID + 1
            End If
        ElseIf TypeOf vntObject Is clsTraço Then
            If vntObject.ID >= ubT Then
                ubT = vntObject.ID + 1
            End If
        End If
    Next
    
    For Each vntGrupo In colGrupo
        If vntGrupo.ID >= ubG Then
            ubG = vntGrupo.ID + 1
        End If
    Next
    
    For Each vntCompo In colComponente
        If vntCompo.ID >= ubP Then
            ubP = vntCompo.ID + 1
        End If
    Next
    
    Objetos.Componente = ubP
    Objetos.Grupo = ubG
    Objetos.Corner = ubC
    Objetos.Ilha = ubI
    Objetos.Traço = ubT
    
    Call RedrawAll
    
Exit Sub
ErrHandler:
  'User pressed the Cancel button
  Exit Sub

End Sub

Public Sub Salvar(Optional eModo As TipoSave)
      
Dim strModo As String

    Select Case eModo
        
        Case Is = TipoSave.Normal
            
            If Prop.Nome = CD1.FileTitle Then
                SaveOpen.FileName = CD1.FileName
                SaveOpen.SalvarMDB colPCB, colGrupo, colComponente
            Else
                strModo = "Milano-PCB Board File (*.mpcb)|*.mpcb"
                If ShowCD(strModo) = True Then
                    SaveOpen.FileName = CD1.FileName
                    SaveOpen.SalvarMDB colPCB, colGrupo, colComponente
                    Prop.Nome = CD1.FileTitle
                End If
            End If
            
        Case Is = TipoSave.NovoNome
            
            strModo = "Milano-PCB Board File (*.mpcb)|*.mpcb"
            If ShowCD(strModo) = True Then
                SaveOpen.FileName = CD1.FileName
                SaveOpen.SalvarMDB colPCB, colGrupo, colComponente
                Prop.Nome = CD1.FileTitle
            End If
        
        Case Is = TipoSave.Componente
            
            If Not (Grupo Is Nothing) Then
                
                Set Componente = New clsComponente
                
                Componente.ID = Grupo.ID
                Componente.Nome = Grupo.Nome
                
                Set Componente = InputBoxMkCom(Componente)
                
                If Not (Componente Is Nothing) Then
                    Call SaveOpen.SalvarComponente(colPCB, Componente)
                End If
                
            Else
                MsgBox "Nenhum grupo foi definido"
            End If
            
            Set Componente = Nothing
            Exit Sub
            
        Case Is = TipoSave.HPLG2
        
            strModo = "Ploter Common Languange (*.pcl)|*.pcl"
            If ShowCD(strModo) = True Then
                SaveOpen.FileName = CD1.FileName
                SaveOpen.SalvarHPLG2 colPCB
            End If
            
        Case Is = TipoSave.Imagem
        
            strModo = "Bitmap (*.bmp)|*.bmp"
            If ShowCD(strModo) = True Then
                SaveOpen.FileName = CD1.FileName
                'SaveOpen
            End If
            
'        Case Is = TipoSave.Copia

    End Select
    
End Sub

Private Function ShowCD(strModo As String) As Boolean
    
On Error GoTo ErrHandler

    CD1.CancelError = True
    CD1.Flags = cdlOFNHideReadOnly
    CD1.Filter = strModo
    CD1.FilterIndex = 1
    CD1.ShowSave
    
    ShowCD = True

Exit Function
ErrHandler:
    
    'Botão cancelar foi precionado
    ShowCD = False
    
End Function

Private Sub Agrupar()
'
'   Agrupa os Objetos selecionados
'

Dim vntObject  As Variant
Dim ColObjAgrupar As New Collection
Dim ColAllCorners As New Collection
Dim intObjCount As Integer
Dim lngID As Long

    intObjCount = 0
    ' algum objeto selecionado
    For Each vntObject In colPCB
        If Not (TypeOf vntObject Is clsCorner) Then
            If vntObject.HasFocus = True Then
                ' completando a coleção de objetos selecionados
                ColObjAgrupar.Add vntObject
            End If
        Else
            ' completando a coleção de todos os corners
            ColAllCorners.Add vntObject, "C" & vntObject.ID
        End If
    Next
    
    ' completando a coleção de objetos selecionados pelos corner
    For Each vntObject In VerObjCorSelec
        ColObjAgrupar.Add vntObject
        intObjCount = intObjCount + 1
    Next
    
    If intObjCount < 2 Then
        ' não ha objetos para agrupar
        Beep
        Exit Sub
    End If

    ' Lista de objetos não selecionados
    For Each vntObject In VerObjCorNotSelec
        If TypeOf vntObject Is clsIlha Then
            ' removendo da lista de todos os corners
            On Error Resume Next
            ColAllCorners.Remove ("C" & vntObject.Corner)
            Err.Clear
        ElseIf TypeOf vntObject Is clsTraço Then
            On Error Resume Next
            ColAllCorners.Remove ("C" & vntObject.StartCorner)
            Err.Clear
            On Error Resume Next
            ColAllCorners.Remove ("C" & vntObject.EndCorner)
            Err.Clear
        End If
    Next
    
    '   ID disponivel
    lngID = Objetos.Grupo
    
    '   Aplicando o novo ID aos objetos
    For Each vntObject In ColObjAgrupar
        If TypeOf vntObject Is clsIlha Then
            Set Ilha = colPCB("Ilha-" & vntObject.ID)
            Ilha.Grupo = lngID
        ElseIf TypeOf vntObject Is clsTraço Then
            Set Traço = colPCB("Traço-" & vntObject.ID)
            Traço.Grupo = lngID
        End If
    Next
    
    '   Aplicando o novo ID aos corners
    For Each vntObject In ColAllCorners
        Set Corner = colPCB("Corner-" & vntObject.ID)
        Corner.Grupo = lngID
    Next
    
    '   Criando o Grupo
    Set Grupo = New clsGrupo
    With Grupo
        .ID = lngID
        .Nome = "Grupo-" & .ID
    End With
    
    colGrupo.Add Grupo, "Grupo-" & Grupo.ID
    
    mPropInterface.ObjectTarget = Grupo
    
End Sub

Private Sub Excluir()
'
'   Exclui os objetos selecionados
'
    
Dim vntCompo            As Variant
Dim vntGrupo            As Variant
Dim vntCorner           As Variant
Dim vntObject           As Variant
Dim ColObjExcluir       As New Collection
Dim ColObjNotExcluir    As New Collection
Dim ColAllCorners       As New Collection
Dim ColCornersExcluir   As New Collection
Dim I                   As Integer
Dim bntFound            As Boolean

    ' Coleção de Corners a excluir
    For Each vntObject In colPCB
        If Not (TypeOf vntObject Is clsCorner) Then
            If vntObject.HasFocus = True Then
                ColObjExcluir.Add vntObject
            End If
        End If
    Next
    
    ' Coleção de Todos os Corners
    For Each vntObject In colPCB
        If TypeOf vntObject Is clsCorner Then
            ColAllCorners.Add vntObject, "C" & vntObject.ID
        End If
    Next
    
    ' Coleção de Objetos Selecionados pelos Corners
    For Each vntObject In VerObjCorSelec
        ColObjExcluir.Add vntObject
    Next
    
    ' Se algum corner pertence a um objeto não selecionado
    ' não remova ele
    For Each vntObject In VerObjCorNotSelec
        If TypeOf vntObject Is clsIlha Then
            For Each vntCorner In ColAllCorners
                If vntCorner.ID = vntObject.Corner Then
                    ColAllCorners.Remove ("C" & vntObject.Corner)
                End If
            Next
        ElseIf TypeOf vntObject Is clsTraço Then
            For Each vntCorner In ColAllCorners
                If vntCorner.ID = vntObject.StartCorner Then
                    ColAllCorners.Remove ("C" & vntObject.StartCorner)
                End If
            Next
            For Each vntCorner In ColAllCorners
                If vntCorner.ID = vntObject.EndCorner Then
                    ColAllCorners.Remove ("C" & vntObject.EndCorner)
                End If
            Next
        End If
    Next
    
    ' Removendo os Objetos
    For Each vntObject In ColObjExcluir
        If TypeOf vntObject Is clsIlha Then
            Set Ilha = vntObject
            Ilha.HasFocus = False
            colPCB.Remove ("Ilha-" & vntObject.ID)
        ElseIf TypeOf vntObject Is clsTraço Then
            Set Traço = vntObject
            Traço.HasFocus = False
            colPCB.Remove ("Traço-" & vntObject.ID)
        End If
    Next
    
    ' Tirando o Foco e removendo o Corner
    For Each vntObject In ColAllCorners
        Set Corner = colPCB("Corner-" & vntObject.ID)
        Corner.HasFocus = False
        colPCB.Remove ("Corner-" & vntObject.ID)
    Next
    
    ' Limpando a lista de Componentes
    For Each vntCompo In colComponente
        bntFound = False
        For Each vntObject In colPCB
            If vntObject.Componente = vntCompo.ID Then
                bntFound = True
                Exit For
            End If
        Next
        If bntFound = False Then
            colComponente.Remove ("Componente-" & vntCompo.ID)
        End If
    Next
    
    ' Limpando a lista de Grupos
    For Each vntGrupo In colGrupo
        bntFound = False
        For Each vntObject In colPCB
            If vntObject.Grupo = vntGrupo.ID Then
                bntFound = True
                Exit For
            End If
        Next
        If bntFound = False Then
            colGrupo.Remove ("Grupo-" & vntGrupo.ID)
        End If
    Next
    
    For I = 1 To ShConector.Count - 1
        Unload ShConector(I)
    Next
    
    For Each vntObject In colPCB
        If TypeOf vntObject Is clsCorner Then
            Set Corner = vntObject
            I = ShConector.Count
            Load ShConector(I)
            Corner.AddShape ShConector(I)
        End If
    Next
    ' Redesenhando
    Call RedrawAll
    
End Sub

Public Property Get HabDesagrupar() As Boolean
    HabDesagrupar = mHabDesagrupar
End Property

Public Property Let HabDesagrupar(NewValue As Boolean)
    mHabDesagrupar = NewValue
    mToolStandart.Buttons("Desagrupar").Enabled = NewValue
End Property

Public Property Get HabAgrupar() As Boolean
    HabAgrupar = mHabAgrupar
End Property

Public Property Let HabAgrupar(NewValue As Boolean)
    mHabAgrupar = NewValue
    mToolStandart.Buttons("Agrupar").Enabled = NewValue
End Property

Private Sub Desagrupar()
'
'   Desagrupa os Objetos Selecionados
'

Dim vntObject  As Variant
Dim vntGrupo As Variant
Dim ColGrupos As New Collection
Dim bntCopy As Boolean
    
    ' determinando quantos grupos estão selecionados
    For Each vntObject In colPCB
    
        If vntObject.HasFocus = True Then
        
            If vntObject.Grupo <> -1 Then
                
                bntCopy = True
                
                For Each vntGrupo In ColGrupos
                    
                    If vntGrupo = vntObject.Grupo Then
                        ' grupo já catalogado
                        bntCopy = False
                    
                    End If
                
                Next
                
                If bntCopy = True Then
                
                    'clase de grupos selecionados
                    ColGrupos.Add vntObject.Grupo
                    
                End If
                
            End If
            
        End If
        
    Next
    
    ' apagando a informação sobre o grupo
    For Each vntGrupo In ColGrupos
        For Each vntObject In colPCB
            If vntObject.Grupo = vntGrupo Then
                ' componentes pertecentes a este grupo
                vntObject.Grupo = -1
                mPropInterface.ObjectTarget = vntObject
            End If
        Next
        colGrupo.Remove ("Grupo-" & vntGrupo)
    Next
    
End Sub

Private Sub AddComponente(Comp As clsComponente, X As Single, Y As Single)
'
'   Adiciona o Componente e Seus Objetos
'

Dim tmpComp As clsComponente
Dim tmpComponentes As New Collection
Dim vntObject As Variant
Dim tmCorner As clsCorner
Dim tmpProp As clsBoardProp

    Set tmpComp = Comp
    
    Set tmpComponentes = SaveOpen.AbrirComponente(tmpComp.DBID)
    tmpComp.ID = Objetos.Componente
    
    Set tmpProp = New clsBoardProp
    
    ' armazenando as configurações
    tmpProp.Layer = Prop.Layer
    tmpProp.IlhaLargura = Prop.IlhaLargura
    tmpProp.IlhaFuro = Prop.IlhaFuro
    tmpProp.LarguraTraço = Prop.LarguraTraço
    
    'Usa as proprias rotinas de criação
    For Each vntObject In tmpComponentes
    
        If TypeOf vntObject Is clsIlha Then
            
            Prop.IlhaLargura = vntObject.IlhaLargura
            Prop.IlhaFuro = vntObject.IlhaFuro
            Prop.Layer = vntObject.Layer
            
            Set tmCorner = tmpComponentes("Corner-" & vntObject.Corner)
            Call AddIlha(X + tmCorner.X, Y + tmCorner.Y)
            
            Ilha.Componente = tmpComp.ID
            
            Set tmCorner = colPCB("Corner-" & Ilha.Corner)
            tmCorner.Componente = tmpComp.ID
        
        ElseIf TypeOf vntObject Is clsTraço Then
        
            Prop.LarguraTraço = vntObject.Largura
            Prop.Layer = vntObject.Layer
            
            Set tmCorner = tmpComponentes("Corner-" & vntObject.StartCorner)
            Elastic.X = X + tmCorner.X
            Elastic.Y = Y + tmCorner.Y
            
            Set tmCorner = tmpComponentes("Corner-" & vntObject.EndCorner)
            Call AddTraço(X + tmCorner.X, Y + tmCorner.Y)
            
            Traço.Componente = tmpComp.ID
            
            Set tmCorner = colPCB("Corner-" & Traço.StartCorner)
            tmCorner.Componente = tmpComp.ID
            
            Set tmCorner = colPCB("Corner-" & Traço.EndCorner)
            tmCorner.Componente = tmpComp.ID
            
        End If
        
    Next
    
    colComponente.Add tmpComp, "Componente-" & tmpComp.ID
    
    ' restaurando as configurações
    Prop.Layer = tmpProp.Layer
    Prop.IlhaLargura = tmpProp.IlhaLargura
    Prop.IlhaFuro = tmpProp.IlhaFuro
    Prop.LarguraTraço = tmpProp.LarguraTraço
    
End Sub

Private Sub ShowLoadComponentes()

    Set LoadCompoForm = New frmLoadComponentes
    LoadCompoForm.ComponentesC = colCompDisponiveis
    LoadCompoForm.ComponentesEmUso = colComponente
    LoadCompoForm.Show 1
    Set LoadCompoForm = Nothing
    
End Sub

Private Function VerificaQuadro(pX As Single, pY As Single) As Boolean

    VerificaQuadro = False
    
    With Prop
        If pX > .X1 Then
            If pY > .Y1 Then
                If pX < .X2 Then
                    If pY < .Y2 Then
                        VerificaQuadro = True
                    End If
                End If
            End If
        End If
    End With

End Function

Private Sub Rotate(Sentido As sSENTIDO)
'
'   Gira uma seleção em 90º
'

Dim vntObject   As Variant
Dim colGirar    As New Collection
Dim retArea     As UDT_LINE_CORD
Dim tmpCord     As UDT_CORD

    For Each vntObject In colSelec
        Set Corner = colPCB("Corner-" & vntObject)
        colGirar.Add Corner, ("Corner-" & vntObject)
    Next
    
    ' definindo o retangulo
    With retArea
        ' descobrindo o Maior X e Y
        For Each vntObject In colGirar
            If .X2 < vntObject.X Then
                .X2 = vntObject.X
            End If
            If .Y2 < vntObject.Y Then
                .Y2 = vntObject.Y
            End If
        Next
        .X1 = .X2
        .Y1 = .Y2
        ' descobrindo o Menor X e Y
        For Each vntObject In colGirar
            If .X1 > vntObject.X Then
                .X1 = vntObject.X
            End If
            If .Y1 > vntObject.Y Then
                .Y1 = vntObject.Y
            End If
        Next
    End With
    
    ' reposicionando
    With retArea
        
        If Sentido = Direita Then
            
            For Each vntObject In colGirar
                tmpCord.X = .X2 + (.Y2 - vntObject.Y)
                tmpCord.Y = .Y2 - (.X2 - vntObject.X)
                vntObject.X = tmpCord.X
                vntObject.Y = tmpCord.Y
            Next
        
        ElseIf Sentido = Esquerda Then
            
            For Each vntObject In colGirar
                tmpCord.X = .X1 - (.Y2 - vntObject.Y)
                tmpCord.Y = .Y1 + (.X2 - vntObject.X)
                vntObject.X = tmpCord.X
                vntObject.Y = tmpCord.Y
            Next
        
        End If
        
    End With
    
    Call RedrawAll
    
End Sub

Private Property Get HabilitaGirar() As Boolean
    
    HabilitaGirar = mHabilitaGirar

End Property

Private Property Let HabilitaGirar(NewValue As Boolean)
    
    mHabilitaGirar = NewValue
    mToolStandart.Buttons("GirarEsquerda").Enabled = NewValue
    mToolStandart.Buttons("GirarDireita").Enabled = NewValue

End Property

Private Sub Copiar()
'
'   Copia os objetos selecionados
'

Dim colPCBCopy                          As New Collection
Dim colComponenteCopy                   As New Collection
Dim colGrupoCopy                        As New Collection
Dim vntCompo            As Variant
Dim vntGrupo            As Variant
Dim vntCorner           As Variant
Dim vntObject           As Variant
Dim ColObjCopiar        As New Collection
Dim ColAllCorners       As New Collection
Dim I                   As Integer
Dim bntFound            As Boolean
Dim tmpComp As clsComponente

    ' Coleção de Objetos com Foco
    For Each vntObject In colPCB
    
        If Not (TypeOf vntObject Is clsCorner) Then
            
            If vntObject.HasFocus = True Then
            
                ColObjCopiar.Add vntObject
                
            End If
            
        End If
        
    Next
    
    ' Coleção de Todos os Corners
    For Each vntObject In colPCB
    
        If TypeOf vntObject Is clsCorner Then
            
            ColAllCorners.Add vntObject, "C" & vntObject.ID
        
        End If
        
    Next
    
    ' Coleção de Objetos Selecionados pelos Corners
    For Each vntObject In VerObjCorSelec
        
        ColObjCopiar.Add vntObject
        
    Next
    
    ' Se algum corner pertence a um objeto não selecionado
    ' não Copie ele
    For Each vntObject In VerObjCorNotSelec
    
        If TypeOf vntObject Is clsIlha Then
            
            For Each vntCorner In ColAllCorners
            
                If vntCorner.ID = vntObject.Corner Then
                
                    ColAllCorners.Remove ("C" & vntObject.Corner)
                    
                End If
                
            Next
        
        ElseIf TypeOf vntObject Is clsTraço Then
            
            For Each vntCorner In ColAllCorners
            
                If vntCorner.ID = vntObject.StartCorner Then
                    
                    ColAllCorners.Remove ("C" & vntObject.StartCorner)
                    
                End If
                
            Next
            
            For Each vntCorner In ColAllCorners
                
                If vntCorner.ID = vntObject.EndCorner Then
                
                    ColAllCorners.Remove ("C" & vntObject.EndCorner)
                
                End If
            
            Next
        
        End If
        
    Next
    
    'limpando colPCBCopy
    Set colPCBCopy = New Collection
    
    ' Copiando os Objetos
    For Each vntObject In ColObjCopiar
        
        If TypeOf vntObject Is clsIlha Then
            
            Set Ilha = vntObject
            colPCBCopy.Add Ilha, "Ilha-" & vntObject.ID
            
        ElseIf TypeOf vntObject Is clsTraço Then
            
            Set Traço = vntObject
            colPCBCopy.Add Traço, "Traço-" & vntObject.ID
        
        End If
        
    Next
    
    ' Copiando os Corner's
    For Each vntObject In ColAllCorners
        Set Corner = colPCB("Corner-" & vntObject.ID)
        colPCBCopy.Add Corner, ("Corner-" & vntObject.ID)
    Next
    
    ' criando a lista de grupo
    For Each vntObject In colPCBCopy
        
        If Not (vntObject.Grupo = -1) Then
            
            bntFound = False
            
            For Each vntGrupo In colGrupoCopy
                If vntGrupo.ID = vntObject.Grupo Then
                    ' grupo já catalogado
                    bntFound = True
                    Exit For
                End If
            Next
            
            If bntFound = False Then
                Set Grupo = colGrupo("Grupo-" & vntObject.Grupo)
                colGrupoCopy.Add Grupo, "Grupo-" & Grupo.ID
            End If
            
        End If
        
    Next
    
    ' criando a lista de componente
    For Each vntObject In colPCBCopy
        
        If Not (vntObject.Componente = -1) Then
            
            bntFound = False
            
            For Each vntCompo In colComponenteCopy
                If vntCompo.ID = vntObject.Componente Then
                    ' componente já catalogado
                    bntFound = True
                    Exit For
                End If
            Next
            
            If bntFound = False Then
                Set tmpComp = colComponente("Componente-" & vntObject.Componente)
                colComponenteCopy.Add tmpComp, "Componente-" & tmpComp.ID
            End If
            
        End If
        
    Next
    
    ' excluindo se existir o arquivo temporario
    If Dir(App.Path & "\Copy.dat") <> "" Then
        Kill App.Path & "\Copy.dat"
    End If
    
    ' incluindo novo arquivo
    SaveOpen.FileName = App.Path & "\Copy.dat"
    Call SaveOpen.SalvarMDB(colPCBCopy, colGrupoCopy, colComponenteCopy)
    
End Sub

Private Sub Colar()
'
'   Cola os objetos que estão no arquivo Copy.dat
'

On Error GoTo ErrHandler

Dim colPCBCopy                          As Collection
Dim colComponenteCopy                   As Collection
Dim colGrupoCopy                        As Collection
Dim tmpCol                              As Collection
Dim oldID                               As Long
Dim ub                                  As Long
Dim vntCorner                           As Variant
Dim tmpComp                             As clsComponente
Dim vntObject                           As Variant
Dim vntGrupo                            As Variant
Dim vntCompo                            As Variant

    SaveOpen.FileName = App.Path & "\Copy.dat"
    Set tmpCol = SaveOpen.Abrir
    
    Set colPCBCopy = tmpCol("PCB")
    Set colGrupoCopy = tmpCol("GRUPO")
    Set colComponenteCopy = tmpCol("COMPO")
    
    ' ajustando os grupos
    For Each vntGrupo In colGrupoCopy
        
        oldID = vntGrupo.ID
        vntGrupo.ID = Objetos.Grupo
        Set Grupo = vntGrupo
        
        For Each vntObject In colPCBCopy
            
            If vntObject.Grupo = oldID Then
                vntObject.Grupo = Grupo.ID
            End If
            
        Next
        
        colGrupo.Add Grupo, "Grupo-" & Grupo.ID
        
    Next
    
    ' ajustando os componente
    For Each vntCompo In colComponenteCopy
        
        oldID = vntCompo.ID
        vntCompo.ID = Objetos.Componente
        Set tmpComp = vntCompo
        
        For Each vntObject In colPCBCopy
            
            If vntObject.Componente = oldID Then
                vntObject.Componente = tmpComp.ID
            End If
            
        Next
        
        colComponente.Add tmpComp, "Componente-" & tmpComp.ID
    
    Next
    
    For Each vntCorner In colPCBCopy
        
        If TypeOf vntCorner Is clsCorner Then
            
            ' Adicionando os shapes
            Set Corner = vntCorner
            ub = ShConector.Count
            Load ShConector(ub)
            Corner.AddShape ShConector(ub)
            
            oldID = Corner.ID
            Corner.ID = Objetos.Corner
            
            ' Identificando os Corners
            For Each vntObject In colPCBCopy
                    
                If TypeOf vntObject Is clsIlha Then
                    
                    If vntObject.Corner = oldID Then
                        vntObject.Corner = Corner.ID
                    End If
                
                ElseIf TypeOf vntObject Is clsTraço Then
                    
                    If vntObject.StartCorner = oldID Then
                        vntObject.StartCorner = Corner.ID
                    ElseIf vntObject.EndCorner = oldID Then
                        vntObject.EndCorner = Corner.ID
                    End If
                
                End If
                
            Next
            
        End If
        
    Next
    
    'Ajustando os IDs
    For Each vntObject In colPCBCopy
        
        If TypeOf vntObject Is clsCorner Then
            
            Set Corner = vntObject
            colPCB.Add Corner, "Corner-" & Corner.ID
            
        ElseIf TypeOf vntObject Is clsIlha Then
            
            Set Ilha = vntObject
            Ilha.ID = Objetos.Ilha
            colPCB.Add Ilha, "Ilha-" & Ilha.ID
        
        ElseIf TypeOf vntObject Is clsTraço Then
            
            Set Traço = vntObject
            Traço.ID = Objetos.Traço
            colPCB.Add Traço, "Traço-" & Traço.ID
        
        End If
        
    Next
    
    Call RedrawAll
    
Exit Sub
ErrHandler:

  Exit Sub

End Sub

Public Sub Imprimir()
'
'   Imprime o arquivo no dispositivo Indicado
'

    
End Sub
