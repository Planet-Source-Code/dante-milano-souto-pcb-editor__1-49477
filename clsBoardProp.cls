VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoardProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
Const m_def_TopLayerColor = 4
Const m_def_DownLayerColor = 2
Const m_def_SilkLayerColor = 6
Const m_def_BackColor = 0
Const m_def_GridColor = 7
Const m_def_Zoom = 1
Const m_def_Enabled = False
Const m_def_LarguraTraço = 1
Const m_def_IlhaLargura = 100
Const m_def_IlhaFuro = 25
Const m_def_FocusColor = vbCyan

Public Enum ActiveTools
    fSeleção = 0
    fTraço = 1
    fIlha = 2
    'fCanto = 3
    fComponente = 4
    fCorte = 5
    fElipse = 6
    fArco = 7
    fRetangulo = 8
End Enum

Public Enum CurrentLayer
    DonwLayer = 0
    TopLayer = 1
    SilkLayer = 2
End Enum

Public Enum PCBColor
    Preto = 0
    Azul = 1
    Verde = 2
    Cyan = 3
    Vermelho = 4
    Magenta = 5
    Amarelo = 6
    Branco = 7
End Enum

Private mZoom               As Double
Private mTopLayerColor      As PCBColor
Private mDownLayerColor     As PCBColor
Private mSilkLayerColor     As PCBColor
Private mBackColor          As PCBColor
Private mGridColor          As PCBColor
Private mLayer              As CurrentLayer
Private mEnabled            As Boolean
Private mGrid               As Boolean
Private mX1                 As Long
Private mX2                 As Long
Private mY1                 As Long
Private mY2                 As Long
Private mLarguraTraço       As Integer
Private mIlhaFuro           As Integer
Private mIlhaLargura        As Integer
Private mFerramenta         As ActiveTools
Private mFocusColor         As Long
Private mHasFocus           As Boolean
Private mNome               As String
'
Public Event Change()
Public Event DrawGrid()
Public Event ClearGrid()
Public Event BackColor(BkColor As PCBColor)
Public Event Nome(NewName As String)
Public Event GotFocus()
Public Event LostFocus()

'
Private Sub Class_Initialize()

    Enabled = m_def_Enabled
    TopLayerColor = m_def_TopLayerColor
    DownLayerColor = m_def_DownLayerColor
    SilkLayerColor = m_def_SilkLayerColor
    BackColor = m_def_BackColor
    GridColor = m_def_GridColor
    Layer = m_def_Layer
    Zoom = m_def_Zoom
    LarguraTraço = m_def_LarguraTraço
    IlhaLargura = m_def_IlhaLargura
    IlhaFuro = m_def_IlhaFuro
    FocusColor = m_def_FocusColor
    mNome = LoadResString(114)
    
End Sub

Public Property Get X1() As Long
    
    X1 = mX1
    
End Property
Public Property Let X1(NewValue As Long)
    
    mX1 = NewValue * 256
    
End Property

Public Property Get X2() As Long
    
    X2 = mX2
    
End Property
Public Property Let X2(NewValue As Long)
    
    mX2 = NewValue
    
End Property

Public Property Get Y1() As Long
    
    Y1 = mY1
    
End Property
Public Property Let Y1(NewValue As Long)
    
    mY1 = NewValue * 256
    
End Property

Public Property Get Y2() As Long
    
    Y2 = mY2
    
End Property
Public Property Let Y2(NewValue As Long)
    
    mY2 = NewValue
    
End Property

Public Property Get Enabled() As Boolean
    
    Enabled = mEnabled
    
End Property
Public Property Let Enabled(NewValue As Boolean)
    
    mEnabled = NewValue
    
End Property

Public Property Get HasFocus() As Boolean
    
    HasFocus = mHasFocus
    
End Property

Public Property Let HasFocus(NewValue As Boolean)
    
    If NewValue = True Then
        If mHasFocus = False Then
            RaiseEvent GotFocus
        End If
    Else
        If mHasFocus = True Then
            RaiseEvent LostFocus
        End If
    End If
    
    mHasFocus = NewValue

End Property

Public Property Get BackColor() As PCBColor

    BackColor = mBackColor
    
End Property

Public Property Let BackColor(NewValue As PCBColor)

    mBackColor = NewValue
    If mEnabled = True Then
        RaiseEvent BackColor(NewValue)
    End If
    
End Property

Public Property Get GridColor() As PCBColor

    GridColor = mGridColor
    
End Property

Public Property Let GridColor(NewValue As PCBColor)

    mGridColor = NewValue
    If mEnabled = True Then
        RaiseEvent DrawGrid
    End If
    
End Property

Public Property Get Grid() As Boolean

    Grid = mGrid
    
End Property

Public Property Let Grid(New_Grid As Boolean)

    mGrid = New_Grid
    If mEnabled = True Then
        RaiseEvent DrawGrid
    End If
    
End Property

Public Property Get SilkLayerColor() As PCBColor

    SilkLayerColor = mSilkLayerColor
    
End Property

Public Property Let SilkLayerColor(NewValue As PCBColor)

    mSilkLayerColor = NewValue
    
End Property

Public Property Get DownLayerColor() As PCBColor

    DownLayerColor = mDownLayerColor
    
End Property

Public Property Let DownLayerColor(NewValue As PCBColor)

    mDownLayerColor = NewValue
    
End Property

Public Property Get TopLayerColor() As PCBColor

    TopLayerColor = mTopLayerColor
    
End Property

Public Property Let TopLayerColor(NewValue As PCBColor)

    mTopLayerColor = NewValue
    
End Property

Public Property Get FocusColor() As Long

    FocusColor = mFocusColor
    
End Property

Public Property Let FocusColor(NewValue As Long)

    mFocusColor = NewValue
    
End Property

Public Property Get Layer() As CurrentLayer

    Layer = mLayer
    
End Property

Public Property Let Layer(vNewValue As CurrentLayer)

    mLayer = vNewValue
    
End Property

Public Property Let Ferramenta(Newferr As ActiveTools)

    mFerramenta = Newferr
    
End Property

Public Property Get Ferramenta() As ActiveTools

    Ferramenta = mFerramenta
    
End Property

Public Property Get Zoom() As Double

    Zoom = mZoom
    
End Property
Public Property Let Zoom(NewValue As Double)

    mZoom = NewValue
    If mEnabled = True Then
        RaiseEvent DrawGrid
    End If
    
End Property

Public Property Get LarguraTraço() As Integer

    LarguraTraço = mLarguraTraço
    
End Property

Public Property Let LarguraTraço(NewValue As Integer)

    mLarguraTraço = NewValue
    
End Property

Public Property Get IlhaLargura() As Integer

    IlhaLargura = mIlhaLargura
    
End Property

Public Property Let IlhaLargura(ByVal NewValue As Integer)

    mIlhaLargura = NewValue
    
End Property

Public Property Get IlhaFuro() As Integer

    IlhaFuro = mIlhaFuro
    
End Property

Public Property Let IlhaFuro(ByVal NewValue As Integer)

    mIlhaFuro = NewValue
    
End Property

Public Sub SetLayerColor(intLayer As CurrentLayer, lngColor As PCBColor)

    Select Case intLayer
        Case Is = CurrentLayer.DonwLayer
            DownLayerColor = lngColor
        Case Is = CurrentLayer.TopLayer
            TopLayerColor = lngColor
        Case Is = CurrentLayer.SilkLayer
            SilkLayerColor = lngColor
    End Select

End Sub

Public Function GetLayerColor(intLayer As CurrentLayer) As PCBColor

    Select Case intLayer
        Case Is = CurrentLayer.DonwLayer
            GetLayerColor = mDownLayerColor
        Case Is = CurrentLayer.TopLayer
            GetLayerColor = mTopLayerColor
        Case Is = CurrentLayer.SilkLayer
            GetLayerColor = mSilkLayerColor
    End Select

End Function

Public Function GetActiveLayerColor() As PCBColor

    Select Case Layer
        Case Is = CurrentLayer.DonwLayer
            GetActiveLayerColor = mDownLayerColor
        Case Is = CurrentLayer.TopLayer
            GetActiveLayerColor = mTopLayerColor
        Case Is = CurrentLayer.SilkLayer
            GetActiveLayerColor = mSilkLayerColor
    End Select

End Function

Public Property Get Nome() As String
    
    Nome = mNome
    
End Property

Public Property Let Nome(NewValue As String)
    
    mNome = NewValue
    RaiseEvent Nome(mNome)
    
End Property


Public Function Aliases(strQuery As String) As String

    Select Case strQuery
        Case Is = "BackColor"
            Aliases = LoadResString(101)
        Case Is = "GridColor"
            Aliases = LoadResString(102)
        Case Is = "SilkLayerColor"
            Aliases = LoadResString(103)
        Case Is = "DownLayerColor"
            Aliases = LoadResString(104)
        Case Is = "TopLayerColor"
            Aliases = LoadResString(105)
        Case Is = "Layer"
            Aliases = LoadResString(106)
        Case Is = "Grid"
            Aliases = LoadResString(107)
        Case Is = "Zoom"
            Aliases = LoadResString(108)
        Case Is = "LarguraTraço"
            Aliases = LoadResString(109)
        Case Is = "IlhaLargura"
            Aliases = LoadResString(110)
        Case Is = "IlhaFuro"
            Aliases = LoadResString(111)
'        Case Is = "Ferramenta"
'            Aliases = LoadResString(112)
        Case Is = "&Nome"
            Aliases = mNome
    End Select
    
End Function

Public Function HelpContexto(strQuery As String) As String

    Select Case strQuery
        Case Is = "BackColor"
            HelpContexto = LoadResString(122)
        Case Is = "GridColor"
            HelpContexto = LoadResString(123)
        Case Is = "SilkLayerColor"
            HelpContexto = LoadResString(124)
        Case Is = "DownLayerColor"
            HelpContexto = LoadResString(124)
        Case Is = "TopLayerColor"
            HelpContexto = LoadResString(124)
        Case Is = "Layer"
            HelpContexto = LoadResString(125)
        Case Is = "Grid"
            HelpContexto = LoadResString(126)
        Case Is = "Zoom"
            HelpContexto = LoadResString(127)
        Case Is = "LarguraTraço"
            HelpContexto = LoadResString(128)
        Case Is = "IlhaLargura"
            HelpContexto = LoadResString(129)
        Case Is = "IlhaFuro"
            HelpContexto = LoadResString(130)
'        Case Is = "Ferramenta"
'            HelpContexto = LoadResString(112)
        Case Is = "&Nome"
            HelpContexto = mNome
    End Select
    
End Function


