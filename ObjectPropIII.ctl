VERSION 5.00
Begin VB.UserControl ObjectPropIII 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ScaleHeight     =   5910
   ScaleWidth      =   3675
   Begin MilanoPCB.HotTextBox HotTextBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picMembers 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   600
      Width           =   3375
      Begin VB.VScrollBar VS 
         Height          =   1815
         Left            =   3000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   255
      End
      Begin MilanoPCB.ObjectMember ObjectMember1 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   609
         BackColor       =   16777215
      End
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
   Begin VB.Shape shTitle 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "ObjectPropIII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oldVS As Integer
Private bntStart As Boolean

Public Property Get ObjectTarget() As Object
    
    Set ObjectTarget = mObjectTarget
    
End Property

Public Property Let ObjectTarget(NewObject As Object)
    
    Call ClearObjects
    
    Set mObjectTarget = NewObject
    If Not (mObjectTarget Is Nothing) Then
        Call LoadObject
    End If
    
End Property

Private Sub LoadObject()

Dim I               As Integer
Dim bntSuported     As Boolean
Dim objM            As Integer
Dim intWidth        As Integer
Dim strTmp          As String
Dim ff              As String
Dim strMsg          As String

    oldVS = 0
    VS.Enabled = False
    VS.Value = 0
    VS.Enabled = True
    
    Set TLIa = New TLIApplication
    Set Interface = TLIa.InterfaceInfoFromObject(mObjectTarget)

    lbTitle.Caption = LoadResString(115) & " - " & CallByName(mObjectTarget, "Aliases", VbMethod, "&Nome")
    
    For I = 1 To Interface.Members.Count
    
        ' inicia o flag
        bntSuported = False
        
        ' apenas get Propriedades
        If Interface.Members(I).InvokeKind = INVOKE_PROPERTYGET Then
            bntSuported = CheckValidDataType(Interface.Members(I).ReturnType)
            If CallByName(mObjectTarget, "Aliases", VbMethod, Interface.Members(I).Name) = "" Then
                bntSuported = False
            End If
        End If
        
        If bntSuported = True Then
            
            objM = ObjectMember1.Count
            
            Load ObjectMember1(objM)
            
            With ObjectMember1(objM)
            
                If objM = 1 Then
                    .Top = 0
                Else
                    .Top = ObjectMember1(objM - 1).Top + ObjectMember1(objM - 1).Height
                End If

                intWidth = 0
                If VS.Visible = True Then
                    intWidth = intWidth + VS.Width + 75
                End If
                .Width = UserControl.ScaleWidth - intWidth
                .Left = 25
                .MemberID = I
                .Visible = True
                
            End With
            
            Call ObjectMember1(objM).HelpBox(HotTextBox1)
            
        End If
        
    Next
    
    Call picMembers_Resize
    
End Sub

Private Sub ClearObjects()

Dim inControl As Control

    HotTextBox1.ControlString = ""
    For Each inControl In UserControl.Controls
        If TypeOf inControl Is ObjectMember Then
            If inControl.Index <> 0 Then
                Unload inControl
            End If
        End If
    Next
    
End Sub

Private Sub picMembers_Resize()
    
Dim objM        As Integer
Dim I           As Integer
Dim intWidth    As Integer
Dim intMax      As Integer

    With VS
        'oldVS = 0
        If ((ObjectMember1.Count - 1) * ObjectMember1(0).Height) > picMembers.ScaleHeight Then
            .Visible = True
            .Max = ((ObjectMember1.Count - 1) * ObjectMember1(0).Height) - picMembers.ScaleHeight
            .SmallChange = ObjectMember1(0).Height
            .LargeChange = .SmallChange
            .Min = 0
        Else
            VS.Visible = False
        End If
        .Top = 0
        .Left = picMembers.ScaleWidth - .Width
        .Height = picMembers.ScaleHeight
    End With
    
    objM = ObjectMember1.Count - 1
    For I = 1 To objM
        With ObjectMember1(I)
            intWidth = 0
            If VS.Visible = True Then
                intWidth = intWidth + VS.Width
            End If
            .Left = 0
            .Width = picMembers.ScaleWidth - intWidth
        End With
    Next
    
End Sub

Private Sub UserControl_Initialize()

    ObjectMember1(0).Visible = False
    bntStart = False
    
End Sub

Private Sub UserControl_Resize()

    With shTitle
        .Top = 25
        .Height = 250
        .Left = 25
        .Width = UserControl.ScaleWidth - 50
    End With
    
    With lbTitle
        .Left = 75
        .Top = 150 - (.Height / 2)
    End With
    
    With picMembers
        .Left = 25
        .Top = shTitle.Top + shTitle.Height + 75
        .Height = (UserControl.ScaleHeight - 1250) - .Top
        .Width = UserControl.ScaleWidth - 50
    End With
    
    With HotTextBox1
        .Left = 25
        .Top = UserControl.ScaleHeight - 1175
        .Height = (UserControl.ScaleHeight - 75) - .Top
        .Width = UserControl.ScaleWidth - 50
    End With
    
End Sub

Private Sub VS_Change()

Dim I As Integer
    
    For I = 1 To ObjectMember1.Count - 1
        ObjectMember1(I).Top = ObjectMember1(I).Top - (VS.Value - oldVS)
    Next
    oldVS = VS.Value
    
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property




