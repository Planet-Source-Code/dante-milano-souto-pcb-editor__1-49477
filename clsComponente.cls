VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mID         As Long
Private mDBID         As Long
Private mNome       As String

Public Property Get ID() As Long
    ID = mID
End Property

Public Property Let ID(NewValue As Long)
    mID = NewValue
End Property

Public Property Get DBID() As Long
    DBID = mDBID
End Property

Public Property Let DBID(NewValue As Long)
    mDBID = NewValue
End Property

Public Property Get Nome() As String
    Nome = mNome
End Property

Public Property Let Nome(NewValue As String)
    mNome = NewValue
End Property

Public Function Aliases(strQuery As String) As String

    Select Case strQuery
        Case Is = "Nome"
            Aliases = LoadResString(119)
        Case Is = "&Nome"
            Aliases = LoadResString(120) & "-" & CStr(mID)
    End Select
    
End Function

Public Function HelpContexto(strQuery As String) As String

    Select Case strQuery
        Case Is = "Nome"
            HelpContexto = LoadResString(119)
        Case Is = "&Nome"
            HelpContexto = LoadResString(120) & "-" & CStr(mID)
    End Select
    
End Function
