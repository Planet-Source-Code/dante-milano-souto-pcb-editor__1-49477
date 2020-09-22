VERSION 5.00
Begin VB.UserControl ObjectMember 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1695
   ScaleWidth      =   4800
   Begin MilanoPCB.MemberValue MemberValue1 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "ObjectMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()
Public Event DblClick()
Public Event Resize()
Public Event Change()

Private mMemberID As Integer
Const m_def_CaptionSize = 16

Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = MemberValue1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    MemberValue1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = MemberValue1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    MemberValue1.List(Index) = New_List
    PropertyChanged "List"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = MemberValue1.ListCount
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = MemberValue1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    MemberValue1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = MemberValue1.Sorted
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = MemberValue1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    MemberValue1.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get MemberID() As Integer
    MemberID = mMemberID
End Property

Public Property Let MemberID(ByVal NewValue As Integer)

On Error Resume Next

Dim strTmp As String

    mMemberID = NewValue
    strTmp = CallByName(mObjectTarget, "Aliases", VbMethod, Interface.Members(mMemberID).Name)
    Err.Clear
    If strTmp = "" Then
        strTmp = Interface.Members(mMemberID).Name
    End If
    Label1.Caption = strTmp
    MemberValue1.MemberID = mMemberID

End Property

Private Sub UserControl_EnterFocus()

    Debug.Print "UserControl_EnterFocus"
    
End Sub

Private Sub UserControl_ExitFocus()

    Debug.Print "UserControl_ExitFocus"

End Sub

Private Sub UserControl_Initialize()
    MemberValue1.BorderStyle = 0
End Sub

Private Sub UserControl_Resize()

Dim txComp As Integer
Dim Cord As UDT_LINE_CORD

    UserControl.Height = MemberValue1.Height + 30
    
    UserControl.Cls
    UserControl.AutoRedraw = True
    UserControl.ForeColor = vbBlack
    
    With Cord
    
        ' linha divisória central
        .X1 = UserControl.TextWidth(String(m_def_CaptionSize, "g"))
        .X2 = .X1
        .Y1 = 0
        .Y2 = UserControl.ScaleHeight
        UserControl.Line (.X1, .Y1)-(.X2, .Y2)
        
        ' linha inferior
        .X1 = 0
        .X2 = UserControl.ScaleWidth
        .Y1 = UserControl.ScaleHeight - (m_def_CaptionSize - 2)
        .Y2 = .Y1
        UserControl.Line (.X1, .Y1)-(.X2, .Y2)
        
    End With
    
    With Label1
    
        .Left = UserControl.TextWidth("g")
        .Top = 10
        .Width = UserControl.TextWidth(String((m_def_CaptionSize - 2), "g"))
        .Height = MemberValue1.Height
        
    End With
    
    With MemberValue1
    
        .Left = UserControl.TextWidth(String((m_def_CaptionSize + 1), "g"))
        .Top = 10
        If (UserControl.ScaleWidth - .Left) > 1 + UserControl.TextWidth("l") Then
            .Width = UserControl.ScaleWidth - (.Left + UserControl.TextWidth("l"))
        End If
        
    End With
    
    RaiseEvent Resize
    
End Sub

Private Sub MemberValue1_Change()
    RaiseEvent Change
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    MemberValue1.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    MemberValue1.List(Index) = PropBag.ReadProperty("List" & Index, "")
    MemberValue1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    MemberValue1.Text = PropBag.ReadProperty("Text", "")
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("Alignment", MemberValue1.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label1")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("List" & Index, MemberValue1.List(Index), "")
    Call PropBag.WriteProperty("ListIndex", MemberValue1.ListIndex, 0)
    Call PropBag.WriteProperty("Text", MemberValue1.Text, "")
    
End Sub

'Public Property Get HelpBox() As HotTextBox
'
'    Set HelpBox = MemberValue1.HelpBox
'
'End Property

Public Sub HelpBox(NewValue As HotTextBox)

    Call MemberValue1.HelpBox(NewValue)
    MemberValue1.HelpString = CallByName(mObjectTarget, "HelpContexto", VbMethod, Interface.Members(mMemberID).Name)
    
End Sub
