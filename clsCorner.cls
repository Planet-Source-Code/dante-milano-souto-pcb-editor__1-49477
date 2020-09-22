VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CornerLayer
    DonwLayer = 0
    TopLayer = 1
    SilkLayer = 2
End Enum

Private mID             As Long
Private mX              As Single
Private mY              As Single
Private mLayer          As CornerLayer
Private mHasFocus       As Boolean
Private mGrupo          As Long
Private mComponente     As Long
Private mShape          As Shape

Public Event Paint()
Public Event Change()
Public Event GotFocus()
Public Event LostFocus()

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

Public Property Get X() As Single

    X = mX
    
End Property

Public Property Let X(NewValue As Single)

    mX = NewValue
    If Not (mShape Is Nothing) Then
        mShape.Left = mX - 25
    End If
    
End Property

Public Property Get Y() As Single

    Y = mY
    
End Property

Public Property Let Y(NewValue As Single)

    mY = NewValue
    If Not (mShape Is Nothing) Then
        mShape.Top = mY - 25
    End If
    
End Property

Public Property Get Layer() As CornerLayer

    Layer = mLayer
    
End Property

Public Property Let Layer(NewValue As CornerLayer)

    mLayer = NewValue
    
End Property

Public Property Get HasFocus() As Boolean
    
    HasFocus = mHasFocus
    
End Property

Public Property Let HasFocus(NewValue As Boolean)
    
    If NewValue = True Then
        If mHasFocus = False Then
            mShape.Left = mX - 25
            mShape.Top = mY - 25
            mShape.Visible = True
            RaiseEvent GotFocus
        End If
    Else
        If mHasFocus = True Then
            mShape.Visible = False
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
        Case Is = "&Nome"
            Aliases = LoadResString(117) & "-" & CStr(mID)
    End Select
    
End Function

Public Function HelpContexto(strQuery As String) As String

    Select Case strQuery
        Case Is = "Layer"
            HelpContexto = LoadResString(106)
        Case Is = "&Nome"
            HelpContexto = LoadResString(117) & "-" & CStr(mID)
    End Select
    
End Function

Public Sub AddShape(inControl As Variant)

    Set mShape = inControl
    With mShape
        .Width = 60
        .Height = 60
        .Visible = False
    End With
    
End Sub

Private Sub Class_Initialize()

    Grupo = -1
    Componente = -1
    
End Sub

Private Sub Class_Terminate()

    If Not (mShape Is Nothing) Then
        Set mShape = Nothing
    End If
    
End Sub

Public Property Get iShape() As Variant

    Set iShape = mShape
    
End Property

Public Property Let iShape(NewValue As Variant)
    
    Set mShape = NewValue
    With mShape
        .Width = 60
        .Height = 60
        .Visible = False
    End With
    
End Property
