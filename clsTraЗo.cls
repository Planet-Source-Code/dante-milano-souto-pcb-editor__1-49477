VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTraço"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TraçoLayer
    DonwLayer = 0
    TopLayer = 1
    SilkLayer = 2
End Enum

Private mStartCorner    As Integer
Private mEndCorner      As Integer
Private mLargura        As Integer
Private mLayer          As TraçoLayer
Private mID             As Long
Private mEnabled        As Boolean
Private mHasFocus       As Boolean
Private mGrupo          As Long
Private mComponente     As Long

Public Event Paint()
Public Event Change()
Public Event GotFocus()
Public Event LostFocus()
'

Public Property Get ID() As Long

    ID = mID
    
End Property

Public Property Let ID(NewValue As Long)

    mID = NewValue
    
End Property

Public Property Get Grupo() As Long

    Grupo = mGrupo
    
End Property

Public Property Let Grupo(NewValue As Long)

    mGrupo = NewValue
    
End Property

Public Property Get Componente() As Long

    Componente = mComponente
    
End Property

Public Property Let Componente(NewValue As Long)

    mComponente = NewValue
    
End Property

Public Property Get Largura() As Integer

    Largura = mLargura
    
End Property

Public Property Let Largura(NewValue As Integer)

    If NewValue <> 0 Then
        mLargura = NewValue
    End If
    
    If mEnabled = True Then
        RaiseEvent Change
    End If
    
End Property

Public Property Get StartCorner() As Integer

    StartCorner = mStartCorner
    
End Property

Public Property Let StartCorner(NewValue As Integer)

    mStartCorner = NewValue
    
End Property

Public Property Get EndCorner() As Integer

    EndCorner = mEndCorner
    
End Property

Public Property Let EndCorner(NewValue As Integer)

    mEndCorner = NewValue
    
End Property

Public Property Get Layer() As TraçoLayer

    Layer = mLayer
    
End Property

Public Property Let Layer(NewValue As TraçoLayer)

    mLayer = NewValue
    If mEnabled = True Then
        RaiseEvent Change
    End If
    
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
    RaiseEvent Paint
    
End Property

Public Sub Refresh()
    
    RaiseEvent Paint
    
End Sub

Public Function Aliases(strQuery As String) As String

    Select Case strQuery
        Case Is = "Layer"
            Aliases = LoadResString(106)
        Case Is = "Largura"
            Aliases = LoadResString(110)
        Case Is = "&Nome"
            Aliases = LoadResString(116) & "-" & CStr(mID)
    End Select
    
End Function

Public Function HelpContexto(strQuery As String) As String

    Select Case strQuery
        Case Is = "Layer"
            HelpContexto = LoadResString(106)
        Case Is = "Largura"
            HelpContexto = LoadResString(110)
        Case Is = "&Nome"
            HelpContexto = LoadResString(116) & "-" & CStr(mID)
    End Select
    
End Function

Private Sub Class_Initialize()

    Enabled = False
    Grupo = -1
    Componente = -1
    
End Sub
