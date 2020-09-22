VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImputMakeComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' clsImputMakeComponente

Private WithEvents mForm As frmMakeComponente
Attribute mForm.VB_VarHelpID = -1
Private tmpCompo As clsComponente
Private bntValido As Boolean

Public Function Valor(mCompo As clsComponente) As clsComponente
Attribute Valor.VB_UserMemId = 0

    Set mForm = New frmMakeComponente

    mForm.ID = mCompo.ID
    mForm.Nome = mCompo.Nome

    Set tmpCompo = Nothing
    mForm.Show 1

    Set Valor = tmpCompo

    Set mForm = Nothing
    
End Function

Private Sub mForm_OkPressed()
    
    Set tmpCompo = New clsComponente
    With tmpCompo
        .ID = mForm.ID
        .Nome = mForm.Nome
    End With
    
End Sub
