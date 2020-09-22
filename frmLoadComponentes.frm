VERSION 5.00
Begin VB.Form frmLoadComponentes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Componentes"
   ClientHeight    =   6210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   0
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
      Begin VB.CommandButton cmdRemAll 
         Caption         =   "<< All"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "All >>"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdRem 
         Caption         =   "<< Rem"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add >>"
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox lstLocal 
         Height          =   3960
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox lstDataBase 
         Height          =   3960
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Disponiveis"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoadComponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' trabalha com a colCompDisponiveis do Board
'

Private DBase           As Database
Private tbComponente    As Recordset
Private mColAllCompo    As New Collection
Private mColCompo       As New Collection
Private mColCompoEmUso  As New Collection
Private mCompo          As clsComponente
'

Public Event Carregados(colCarregados As Collection)

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub cmdAdd_Click()
'
'   Adiciona um Item da Lista Database para a lista local
'

    If lstDataBase.ListIndex <> -1 Then
        lstLocal.AddItem lstDataBase.List(lstDataBase.ListIndex)
        lstDataBase.RemoveItem (lstDataBase.ListIndex)
    End If
    
End Sub

Private Sub cmdAddAll_Click()
'
'   Adiciona todos os Itens da Lista Database para a lista local
'
Dim I As Long

    For I = lstDataBase.ListCount - 1 To 0 Step -1
        lstLocal.AddItem lstDataBase.List(I)
        lstDataBase.RemoveItem (I)
    Next

End Sub

Private Sub cmdRem_Click()
'
'   Remove um Item da Lista Local para a Lista Database
'
Dim vntObject As Variant
Dim strT() As String
Dim bntMove As Boolean

    If lstLocal.ListIndex <> -1 Then
        bntMove = True
        For Each vntObject In mColCompoEmUso
            strT = Split(lstLocal.List(lstLocal.ListIndex))
            If strT(1) = vntObject.Nome Then
                bntMove = False
            End If
        Next
        If bntMove = True Then
            lstDataBase.AddItem lstLocal.List(lstLocal.ListIndex)
            lstLocal.RemoveItem (lstLocal.ListIndex)
        End If
    End If
    
End Sub

Private Sub cmdRemAll_Click()
'
'   Remove Todos os Items da Lista Local para a Lista Database
'
Dim I As Long
Dim bntMove As Boolean
Dim vntObject As Variant
Dim strT() As String

    For I = lstLocal.ListCount - 1 To 0 Step -1
        bntMove = True
        For Each vntObject In mColCompoEmUso
            strT = Split(lstLocal.List(I))
            If strT(1) = vntObject.Nome Then
                bntMove = False
            End If
        Next
        If bntMove = True Then
            lstDataBase.AddItem lstLocal.List(I)
            lstLocal.RemoveItem (I)
        End If
    Next

End Sub

Private Sub Form_Load()

    Call Carregar
    
End Sub

Private Sub Carregar()

On Error GoTo errTrat
Dim vntObject As Object

    Set DBase = OpenDatabase(App.Path & "\componentes.dat", True)
    Set tbComponente = DBase.OpenRecordset("COMPONENTE")
    
    With tbComponente
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF = True
                Set mCompo = New clsComponente
                mCompo.DBID = !ID
                mCompo.Nome = ChkNull(!Nome)
                mColAllCompo.Add mCompo, "Componente-" & mCompo.DBID
                .MoveNext
            Loop
        End If
    End With
    
    tbComponente.Close
    DBase.Close
    Set DBase = Nothing
    Set tbComponente = Nothing
    
    For Each vntObject In mColAllCompo
        lstDataBase.AddItem vntObject.DBID & " " & vntObject.Nome
    Next
    
Exit Sub
errTrat:

    MsgBox "Impossível carregar"

End Sub

Public Property Get ComponentesC() As Collection
'
'   Componentes Disponiveis
'
    
    Set ComponentesC = mColCompo
    
End Property

Public Property Let ComponentesC(NewValue As Collection)
    
Dim vntObject As Variant

    Set mColCompo = NewValue
    
    For Each vntObject In mColCompo
        lstLocal.AddItem vntObject.ID & " " & vntObject.Nome
    Next
    
End Property

Public Property Get ComponentesEmUso() As Collection
'
'   Componentes 'Em Uso' não podes ser removidos da lista
'   (na verdade não afetaria em nada)
'
    
    Set ComponentesEmUso = mColCompoEmUso
    
End Property

Public Property Let ComponentesEmUso(NewValue As Collection)
    
    Set mColCompoEmUso = NewValue
    
End Property

Private Sub OKButton_Click()

Dim I As Integer
Dim tmpCompo As clsComponente
Dim strT() As String
    
    For I = I To mColCompo.Count - 1
        mColCompo.Remove (1)
    Next
    For I = 0 To lstLocal.ListCount - 1
        strT = Split(lstLocal.List(I), " ")
        Set tmpCompo = New clsComponente
        With tmpCompo
            .DBID = strT(0)
            .Nome = strT(1)
        End With
        mColCompo.Add tmpCompo
    Next
    RaiseEvent Carregados(mColCompo)
    Unload Me
    
End Sub

