VERSION 5.00
Begin VB.Form frmSelectComponente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione um Componente"
   ClientHeight    =   4770
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5775
      TabIndex        =   5
      Top             =   0
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
      Begin VB.ListBox lstLocal 
         Height          =   2790
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblNome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cdmOK 
      Caption         =   "Selecionar"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mColCompo       As New Collection

Public Event Selecionado(mComponente As clsComponente)


Public Property Get ComponentesC() As Collection
    Set ComponentesC = mColCompo
End Property

Public Property Let ComponentesC(NewValue As Collection)
    
Dim vntObject As Variant

    Set mColCompo = NewValue
    
    For Each vntObject In mColCompo
        lstLocal.AddItem vntObject.DBID & " " & vntObject.Nome
    Next
    
End Property

Private Sub cdmOK_Click()
    
Dim tmCompo As clsComponente
Dim vntObject As Variant
Dim strT() As String

    If lstLocal.ListIndex <> -1 Then
        For Each vntObject In mColCompo
            Set tmCompo = vntObject
            strT = Split(lstLocal.List(lstLocal.ListIndex), " ")
            If tmCompo.Nome = strT(1) Then
                RaiseEvent Selecionado(tmCompo)
            End If
        Next
    End If
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
End Sub

Private Sub lstLocal_Click()

Dim tmCompo As clsComponente
Dim vntObject As Variant
Dim strT() As String

    If lstLocal.ListIndex <> -1 Then
        For Each vntObject In mColCompo
            Set tmCompo = vntObject
            strT = Split(lstLocal.List(lstLocal.ListIndex), " ")
            If tmCompo.Nome = strT(1) Then
                lblNome.Caption = tmCompo.DBID & " " & tmCompo.Nome
            End If
        Next
    End If
    
End Sub
