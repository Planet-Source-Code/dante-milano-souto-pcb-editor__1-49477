VERSION 5.00
Begin VB.Form frmMakeComponente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Componente"
   ClientHeight    =   3120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
      Begin VB.TextBox texNome 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmMakeComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkPressed()
Public Event CancelPressed()

Private mID As Long

Public Property Get ID() As String
    ID = mID
End Property

Public Property Let ID(NewValue As String)
    mID = NewValue
End Property

Public Property Get Nome() As String
    Nome = texNome.Text
End Property

Public Property Let Nome(NewValue As String)
    texNome.Text = Trim(NewValue)
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
'
'
'
    
    If texNome.Text = "" Then
        Beep
        texNome.SelStart = 1
        texNome.SelLength = Len(texNome.Text)
        texNome.SetFocus
    Else
        RaiseEvent OkPressed
        Unload Me
    End If
    
End Sub
