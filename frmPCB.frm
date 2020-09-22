VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPCB 
   BackColor       =   &H8000000A&
   Caption         =   "MilanPCB -"
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8355
   Icon            =   "frmPCB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   557
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmMousePinter 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   600
   End
   Begin MilanoPCB.ObjectPropIII ObjectPropIII1 
      Height          =   5055
      Left            =   5400
      TabIndex        =   6
      Top             =   600
      Width           =   2895
      _extentx        =   5106
      _extenty        =   8916
   End
   Begin MSComctlLib.ImageList imgList2 
      Left            =   1800
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":014A
            Key             =   "Seleção"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":06E4
            Key             =   "Traço"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":083E
            Key             =   "Ilha"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":0998
            Key             =   "CutTraço"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":0F32
            Key             =   "Componente"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolLeft 
      Align           =   3  'Align Left
      Height          =   5385
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   9499
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList1 
      Left            =   1200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":14CC
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1626
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1780
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":18DA
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1A34
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1B8E
            Key             =   "Recortar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1CE8
            Key             =   "Colar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1E42
            Key             =   "Propriedades"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":1F9C
            Key             =   "Ajuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":20F6
            Key             =   "SilkLayer"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":2250
            Key             =   "TopLayer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":27EA
            Key             =   "DownLayer"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":2944
            Key             =   "Agrupar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":2A56
            Key             =   "Desagrupar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":2B68
            Key             =   "Ferramentas"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":2FBA
            Key             =   "Componentes"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":3114
            Key             =   "GirarDireita"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPCB.frx":36AE
            Key             =   "GirarEsquerda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolTop 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Zoom"
         Top             =   0
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9102
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pctBoard 
      Height          =   3975
      Left            =   720
      ScaleHeight     =   3915
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
      Begin VB.HScrollBar HS 
         Height          =   255
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3795
      End
      Begin VB.VScrollBar VS 
         Height          =   3555
         Left            =   3960
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin MilanoPCB.Board Board1 
         Height          =   3495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3855
         _extentx        =   6800
         _extenty        =   6165
      End
   End
   Begin VB.Shape shMove 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   5280
      Top             =   720
      Width           =   135
   End
   Begin VB.Menu mnu_Arquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnu_Arquivo_Novo 
         Caption         =   "Novo"
      End
      Begin VB.Menu mnu_Arquivo_Abrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnu_Arquivo_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Arquivo_Salvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Arquivo_SalvarComo 
         Caption         =   "Salvar Como"
      End
      Begin VB.Menu mnu_Arquivo_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Arquivo_MakeComponente 
         Caption         =   "Make a Componente"
      End
      Begin VB.Menu mnu_Arquivo_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Arquivo_Imprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_Arquivo_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Arquivo_Sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnu_Editar 
      Caption         =   "Editar"
      Begin VB.Menu mnu_Editar_Copiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnu_Editar_Recortar 
         Caption         =   "Recortar"
      End
      Begin VB.Menu mnu_Editar_Colar 
         Caption         =   "Colar"
      End
      Begin VB.Menu mnu_Editar_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Editar_Propriedades 
         Caption         =   "Propriedades"
      End
   End
   Begin VB.Menu mnu_Exibir 
      Caption         =   "Exibir"
      Begin VB.Menu mnu_Exibir_Prop 
         Caption         =   "Janela de Propriedades"
      End
      Begin VB.Menu mnu_Exibir_BarraStatus 
         Caption         =   "Barra de Status"
      End
      Begin VB.Menu mnu_Exibir_FerramentasStandart 
         Caption         =   "Barra de Ferramentas Standart"
      End
      Begin VB.Menu mnu_Exibir_Ferramentas 
         Caption         =   "Barra de Ferramentas Edição"
      End
   End
   Begin VB.Menu mnu_Ferramentas 
      Caption         =   "Ferramentas"
      Begin VB.Menu mnu_Ferramentas_Agrupar 
         Caption         =   "Agrupar"
      End
      Begin VB.Menu mnu_Ferramentas_Desagrupar 
         Caption         =   "Desagrupar"
      End
      Begin VB.Menu mnu_Ferramentas_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Ferramentas_Opções 
         Caption         =   "Opções"
         Begin VB.Menu mnu_Ferramentas_Opções_Board 
            Caption         =   "Board"
         End
      End
   End
   Begin VB.Menu mnu_Ajuda 
      Caption         =   "Ajuda"
      Begin VB.Menu mnu_Ajuda_Sobre 
         Caption         =   "Sobre o Milano-PCB"
      End
      Begin VB.Menu mnu_Ajuda_OnLine 
         Caption         =   "Ajuda On Line"
      End
   End
End
Attribute VB_Name = "frmPCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_def_BackColor = vbBlack

Private mResizePropControl      As Boolean
Private movPos                  As Integer
'

Private Property Get ResizePropControl() As Boolean
    ResizePropControl = mResizePropControl
End Property

Private Property Let ResizePropControl(NewValue As Boolean)
    
    mResizePropControl = NewValue
    shMove.Visible = mResizePropControl
    shMove.ZOrder 0
    pctBoard.Enabled = Not (mResizePropControl)
    Board1.Enabled = Not (mResizePropControl)
    ObjectPropIII1.Enabled = Not (mResizePropControl)
    
End Property

Private Sub Board1_CaptionChange(NovoCaption As String)

    Me.Caption = "Milano PCB - " & NovoCaption
    
End Sub

Private Sub cmbZoom_Click()

    Board1.Prop.Zoom = Val(cmbZoom.List(cmbZoom.ListIndex)) / 100
    
End Sub

Private Sub Form_Load()

Dim strTmp As String

    ' Apresentando a Licença
        On Error GoTo Continua
        Open App.Path & "\Lc.lic" For Input As #1
            Line Input #1, strTmp
        Close #1
        If strTmp <> "1" Then
            Dialog.Show 1
        End If
        
Continua:
    
        If Err Then
            Dialog.Show 1
        End If
        
        If Dir(App.Path & "\MilanPCB.log") <> "" Then
            Kill App.Path & "\MilanPCB.log"
        End If
    '
    
    ' Para que os controles aparessam posicionados corretamente
    Me.ScaleHeight = Me.Height
    Me.ScaleWidth = Me.Width
    '
    
    Call LoadControls

    If Command <> "" Then
        Board1.Abrir (Command)
    End If
        
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim tmpCord As UDT_LINE_CORD

    ResizePropControl = False
    If ObjectPropIII1.Visible = True Then
        
        With tmpCord
        
            If VS.Visible = True Then
                .X1 = VS.Left + VS.Width
            Else
                .X1 = pctBoard.Left + pctBoard.Width
            End If
            
            .X2 = ObjectPropIII1.Left
            
            If ToolTop.Visible = True Then
                .Y1 = ToolTop.Height
            Else
                .Y1 = 0
            End If
            
            If StatusBar1.Visible = True Then
                .Y2 = Me.ScaleHeight - StatusBar1.Height
            Else
                .Y2 = Me.ScaleHeight
            End If
            
            If X > .X1 Then
                If X < .X2 Then
                    If Y > .Y1 Then
                        If Y < .Y2 Then
                            ResizePropControl = True
                            movPos = ObjectPropIII1.Left - X
                        End If
                    End If
                End If
            End If
        
        End With
    
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ResizePropControl = True Then
        ObjectPropIII1.Width = Me.ScaleWidth - (X + movPos)
        Call Form_Resize
    End If
    
    Dim tmpCord As UDT_LINE_CORD

    If ObjectPropIII1.Visible = True Then
        
        With tmpCord
        
            If VS.Visible = True Then
                .X1 = VS.Left + VS.Width
            Else
                .X1 = pctBoard.Left + pctBoard.Width
            End If
            
            .X2 = ObjectPropIII1.Left
            
            If ToolTop.Visible = True Then
                .Y1 = ToolTop.Height
            Else
                .Y1 = 0
            End If
            
            If StatusBar1.Visible = True Then
                .Y2 = Me.ScaleHeight - StatusBar1.Height
            Else
                .Y2 = Me.ScaleHeight
            End If
            
            If X > .X1 Then
                If X < .X2 Then
                    If Y > .Y1 Then
                        If Y < .Y2 Then
                            Me.MousePointer = 9
                            tmMousePinter.Enabled = True
                            shMove.Left = .X1 + 10
                            shMove.Top = .Y1
                            shMove.Width = .X2 - (.X1 + 20)
                            shMove.Height = .Y2 - .Y1
                        Else
                            Me.MousePointer = 0
                        End If
                    Else
                        Me.MousePointer = 0
                    End If
                Else
                    Me.MousePointer = 0
                End If
            Else
                Me.MousePointer = 0
            End If
        
        End With
        
    Else
    
        Me.MousePointer = 0
    
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ResizePropControl = False
    Me.MousePointer = 0
    
End Sub

Private Sub Form_Resize()

On Error GoTo errTrat
    
Dim intWidth As Integer
Dim intHeight As Integer

    With pctBoard
        
        '   Width
        intWidth = 0
        If ToolLeft.Visible = True Then
            intWidth = intWidth + ToolLeft.Width
        End If
        .Left = intWidth

        If ObjectPropIII1.Visible = True Then
            intWidth = intWidth + (ObjectPropIII1.Width + 100)
        End If
        .Width = Me.ScaleWidth - intWidth
        
        '   Height
        intHeight = 0
        If ToolTop.Visible = True Then
             intHeight = intHeight + ToolTop.Height
        End If
        .Top = intHeight

        If StatusBar1.Visible = True Then
            intHeight = intHeight + StatusBar1.Height
        End If
        .Height = Me.ScaleHeight - intHeight
        
    End With
    
    intHeight = 0
    With ObjectPropIII1
        
        '   Width
        intWidth = 0
        If .Visible = True Then
            intWidth = intWidth + .Width
        End If
        .Left = Me.ScaleWidth - intWidth
        
        '   Height
        intHeight = 0
        If ToolTop.Visible = True Then
            intHeight = intHeight + ToolTop.Height
        End If
        .Top = intHeight
        If StatusBar1.Visible = True Then
            intHeight = intHeight + StatusBar1.Height
        End If
        .Height = Me.ScaleHeight - intHeight
        
    End With

    
    With shMove
        
        .Visible = mResizePropControl
        
        '   Width
        .Left = ObjectPropIII1.Left - 80
        .Width = 60
        
        '   Height
        .Top = ObjectPropIII1.Top
        .Height = ObjectPropIII1.Height
        
    End With
    
Exit Sub
errTrat:
        
'        Debug.Print Err.Number & " " & Err.Description
        
End Sub

Private Sub LoadControls()

Dim I As Integer

    ToolTop.ImageList = imgList1
    With ToolTop.Buttons
        .Add , "Novo", , tbrDefault, "Novo"
        .Add , "Abrir", , tbrDefault, "Abrir"
        .Add , , , tbrPlaceholder
        .Add , "Salvar", , tbrDefault, "Salvar"
        .Add , "Imprimir", , tbrDefault, "Imprimir"
        .Add , , , tbrPlaceholder
        .Add , "Copiar", , tbrDefault, "Copiar"
        .Add , "Recortar", , tbrDefault, "Recortar"
        .Add , "Colar", , tbrDefault, "Colar"
        .Add , , , tbrPlaceholder
        .Add , "Propriedades", , tbrDefault, "Propriedades"
        .Add , "Ferramentas", , tbrDefault, "Ferramentas"
        .Add , "Componentes", , tbrDefault, "Componentes"
        .Add , , , tbrPlaceholder
        .Add , "GirarEsquerda", , tbrDefault, "GirarEsquerda"
        .Add , "GirarDireita", , tbrDefault, "GirarDireita"
        .Add , , , tbrPlaceholder
        .Add , "TopLayer", , tbrDefault, "TopLayer"
        .Add , "DownLayer", , tbrDefault, "DownLayer"
        .Add , "SilkLayer", , tbrDefault, "SilkLayer"
        .Add , , , tbrPlaceholder
        .Add , "Agrupar", , tbrDefault, "Agrupar"
        .Add , "Desagrupar", , tbrDefault, "Desagrupar"
        .Add , , , tbrPlaceholder
        .Add , "Ajuda", , tbrDefault, "Ajuda"
        ' *-*-*
        For I = 1 To .Count
            .Item(I).ToolTipText = .Item(I).Key
        Next
        ' *-*-*
    End With
    
    ToolLeft.ImageList = imgList2
    With ToolLeft.Buttons
        .Add 1, "Seleção", , tbrDefault, "Seleção"
        .Add , , , tbrPlaceholder
        .Add 2, "Traço", , tbrDefault, "Traço"
        .Add 3, "Ilha", , tbrDefault, "Ilha"
        .Add 5, "Componente", , tbrDefault, "Componente"
        .Add , , , tbrPlaceholder
        .Add 6, "CutTraço", , tbrDefault, "CutTraço"
        ' *-*-*
        For I = 1 To .Count
            .Item(I).ToolTipText = .Item(I).Key
        Next
        ' *-*-*
    End With
    
    With cmbZoom
        .AddItem "50 %"
        .AddItem "75 %"
        .AddItem "100 %"
        .AddItem "150 %"
        .AddItem "200 %"
        .AddItem "400 %"
        .ListIndex = 2
        .Visible = False
    End With

    With Board1
        .Width = pctBoard.ScaleWidth
        .Height = pctBoard.ScaleHeight
        .Left = 0
        .Top = 0
        .Prop.Enabled = True
        .Prop.Grid = True
        .PropInterface = ObjectPropIII1
        .ToolStandart = ToolTop
        .ToolStandartObjects = ToolLeft
        .BarraStatus = StatusBar1
        .HS = HS
        .VS = VS
    End With
    
    mnu_Exibir_FerramentasStandart.Checked = True
    mnu_Exibir_BarraStatus.Checked = True
    mnu_Exibir_Ferramentas.Checked = True
    mnu_Exibir_Prop.Checked = True
    
    ToolTop.Buttons.Item("Ferramentas").Value = tbrPressed
    ToolTop.Buttons.Item("Propriedades").Value = tbrPressed
    
    shMove.Visible = False
    
    Me.Caption = "Milano PCB - " & Board1.Prop.Nome
    
End Sub

Private Sub mnu_Ajuda_OnLine_Click()
'
'
'

Dim iret As Long

    ' open URL into the default internet browser
    iret = ShellExecute(Me.hwnd, vbNullString, "http://geocities.yahoo.com.br/dantemilanosouto/", vbNullString, "c:\", 1)
    
End Sub

Private Sub mnu_Ajuda_Sobre_Click()

    Load frmAbaut
    frmAbaut.imgIcon.Picture = Me.Icon
    frmAbaut.Show 1
    
End Sub

Private Sub mnu_Arquivo_Abrir_Click()

    Board1.Abrir
    
End Sub

Private Sub mnu_Arquivo_MakeComponente_Click()

    Call Board1.Salvar(Componente)
    
End Sub

Private Sub mnu_Arquivo_Salvar_Click()

    Call Board1.Salvar(Normal)
    
End Sub

Private Sub mnu_Arquivo_SalvarComo_Click()

    Call Board1.Salvar(NovoNome)

End Sub

Private Sub mnu_Exibir_BarraStatus_Click()
    
    mnu_Exibir_BarraStatus.Checked = Not (mnu_Exibir_BarraStatus.Checked)
    StatusBar1.Visible = mnu_Exibir_BarraStatus.Checked
    Call Form_Resize
    
End Sub

Private Sub mnu_Exibir_Ferramentas_Click()

    mnu_Exibir_Ferramentas.Checked = Not (mnu_Exibir_Ferramentas.Checked)
    ToolLeft.Visible = mnu_Exibir_Ferramentas.Checked
    If mnu_Exibir_Ferramentas.Checked = True Then
        ToolTop.Buttons.Item("Ferramentas").Value = tbrPressed
    Else
        ToolTop.Buttons.Item("Ferramentas").Value = tbrUnpressed
    End If
    Call Form_Resize
    
End Sub

Private Sub mnu_Exibir_FerramentasStandart_Click()
    
    mnu_Exibir_FerramentasStandart.Checked = Not (mnu_Exibir_FerramentasStandart.Checked)
    ToolTop.Visible = mnu_Exibir_FerramentasStandart.Checked
    Call Form_Resize
    
End Sub

Private Sub mnu_Exibir_Prop_Click()

    mnu_Exibir_Prop.Checked = Not (mnu_Exibir_Prop.Checked)
    ObjectPropIII1.Visible = mnu_Exibir_Prop.Checked
    If mnu_Exibir_Prop.Checked = True Then
        ToolTop.Buttons.Item("Propriedades").Value = tbrPressed
    Else
        ToolTop.Buttons.Item("Propriedades").Value = tbrUnpressed
    End If
    Call Form_Resize
    
End Sub

Private Sub tmMousePinter_Timer()
    Me.MousePointer = 0
    tmMousePinter.Enabled = False
End Sub

Private Sub ToolTop_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        Case Is = "Ferramentas"
            Call mnu_Exibir_Ferramentas_Click
        Case Is = "Propriedades"
            Call mnu_Exibir_Prop_Click
    End Select
            
End Sub

Private Sub VS_Change()

    'Board1.Top = VS.Value * -1

End Sub

Private Sub mnu_Ferramentas_Opções_Board_Click()
    
    'Board1.ShowSetings
    
End Sub

Private Sub pctBoard_Resize()
    
On Error GoTo errTrat
    
    With Board1
        .Left = 0
        .Top = 0
        .Width = pctBoard.ScaleWidth - VS.Width
        .Height = pctBoard.ScaleHeight - HS.Height
    End With

    With VS
        .Left = pctBoard.ScaleWidth - .Width
        .Top = 0
        .Height = pctBoard.ScaleHeight - HS.Height
    End With
    
    With HS
        .Left = 0
        .Width = pctBoard.ScaleWidth - VS.Width
        .Top = pctBoard.ScaleHeight - .Height
    End With
        
errTrat:

End Sub


