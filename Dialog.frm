VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrato de Licença"
   ClientHeight    =   4875
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6345
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6345
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2460
      Top             =   4140
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1140
      Width           =   6315
   End
   Begin VB.OptionButton optNaoAceito 
      Caption         =   "Não Aceito"
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   4380
      Width           =   1515
   End
   Begin VB.OptionButton optAceito 
      Caption         =   "Aceito"
      Height          =   375
      Left            =   900
      TabIndex        =   2
      Top             =   3960
      Width           =   1395
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   435
      Left            =   5100
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "Continuar"
      Height          =   435
      Left            =   3720
      TabIndex        =   0
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contrato de Licença de Uso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   720
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   0
      Picture         =   "Dialog.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdContinuar_Click()

    Open App.Path & "\Lc.lic" For Output As #1
        Print #1, "1"
    Close
    Call Associar
    Unload Me
        
End Sub

Private Sub cmdSair_Click()
'
'
'
    
    End
    
End Sub

Private Sub Form_Load()
'
'
'

Dim strFile As String

    Open App.Path & "\Licença.txt" For Input As #1
        Text1.Text = Input(LOF(1), #1)
    Close #1

End Sub

Private Sub Timer1_Timer()

    If optAceito.Value = True Then
        cmdContinuar.Enabled = True
    Else
        cmdContinuar.Enabled = False
    End If
    
End Sub
