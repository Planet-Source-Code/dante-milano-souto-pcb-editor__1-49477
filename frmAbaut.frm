VERSION 5.00
Begin VB.Form frmAbaut 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSairbtn 
      Interval        =   2000
      Left            =   7320
      Top             =   5940
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "cmdContinuar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   3735
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Index           =   0
      X1              =   60
      X2              =   7980
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Image imgIcon 
      Height          =   615
      Left            =   420
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lb1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App.Copryrigth"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   300
      TabIndex        =   3
      Top             =   6600
      Width           =   4785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Index           =   3
      X1              =   7980
      X2              =   7980
      Y1              =   60
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Index           =   2
      X1              =   60
      X2              =   60
      Y1              =   60
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Index           =   1
      X1              =   60
      X2              =   7980
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   5100
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   300
      X2              =   7680
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Label lb1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App.Versão"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   6300
      Width           =   2565
   End
   Begin VB.Label lb1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App.FileDescription"
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   5820
      Width           =   6645
   End
   Begin VB.Label lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App.Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   540
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   2550
      Left            =   5580
      Picture         =   "frmAbaut.frx":0000
      Stretch         =   -1  'True
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbaut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSair_Click()
'
'   Escapa
'

    Unload Me
    
End Sub

Private Sub Form_Load()
'
'   Carrega o texto
'

    lb1(0).Caption = App.ProductName
    lb1(1).Caption = App.FileDescription
    lb1(2).Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub tmrSairbtn_Timer()
'
'   Habilita sair depoi de 2 segundos
'

    cmdSair.Visible = True
    cmdSair.Enabled = True
    tmrSairbtn.Enabled = False
    
End Sub

Private Sub tmrUnload_Timer()
'
'   Escapa após um tempo pré determinado de 5 seg.
'

    Unload Me
    
End Sub
