VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'   Clase contadora de Objetos
'

Private mCorner     As Long
Private mTraço      As Long
Private mIlha       As Long
Private mGrupo      As Long
Private mComponente As Long
'

Property Get Corner() As Long
    Corner = mCorner
    mCorner = mCorner + 1
End Property

Property Let Corner(NewValue As Long)
    mCorner = NewValue
End Property

Property Get Ilha() As Long
    Ilha = mIlha
    mIlha = mIlha + 1
End Property

Property Let Ilha(NewValue As Long)
    mIlha = NewValue
End Property

Property Get Traço() As Long
    Traço = mTraço
    mTraço = mTraço + 1
End Property

Property Let Traço(NewValue As Long)
    mTraço = NewValue
End Property

Property Get Grupo() As Long
    Grupo = mGrupo
    mGrupo = mGrupo + 1
End Property

Property Let Grupo(NewValue As Long)
    mGrupo = NewValue
End Property

Property Get Componente() As Long
    Componente = mComponente
    mComponente = mComponente + 1
End Property

Property Let Componente(NewValue As Long)
    mComponente = NewValue
End Property

