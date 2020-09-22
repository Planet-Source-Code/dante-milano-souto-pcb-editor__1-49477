VERSION 5.00
Begin VB.UserControl MemberValue 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Label1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "MemberValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents mHelpBox As HotTextBox

Private SpTextbox As New clsTextBox

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Const m_def_Tipo = 0

Enum TpView
    ViewText = 0
    ViewList = 1
End Enum
       
Private mHelpString As String
Private m_Tipo As TpView
Private mMemberID As Integer
Private strOld As String
Private intOld As Integer
'

Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub Combo1_Click()
    Call UpdateChanges
End Sub

Private Sub Combo1_GotFocus()
    intOld = Combo1.ListIndex
End Sub

Private Sub mHelpBox_HotSpotClick(Index As Long)
    Debug.Print "Index"
End Sub

Private Sub Text1_Change()
    Call UpdateChanges
    RaiseEvent Change
End Sub

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub Text1_GotFocus()
    strOld = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MultiLine() As Boolean
    MultiLine = Text1.MultiLine
End Property

Public Property Get ScrollBars() As Integer
    ScrollBars = Text1.ScrollBars
End Property

Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    If Tipo = 0 Then
        Text = Text1.Text
    Else
        Text = Combo1.Text
    End If
End Property

Public Property Let Text(ByVal New_Text As String)
    
    Text1.Text = New_Text
    PropertyChanged "Text"

End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Text1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Text1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Sub UpdateChanges()
On Error GoTo errTrat

Dim strSplit() As String
Dim varTmp As Variant
    
    If mMemberID <> -1 Then
    
        If Tipo = ViewText Then
            varTmp = Text1.Text
        Else
            If InStr(1, Combo1.Text, "=") <> 0 Then
                strSplit = Split(Combo1.Text, "=")
                varTmp = Trim(strSplit(1))
            Else
                varTmp = Combo1.Text
            End If
        End If
    
        Select Case Interface.Members(mMemberID).ReturnType
            Case Is = VarEnum.vtEnum
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CLng(varTmp))
            Case Is = VarEnum.vtBoolean
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CBool(varTmp))
            Case Is = VarEnum.vtByte
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CByte(varTmp))
            Case Is = VarEnum.vtInteger
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CInt(varTmp))
            Case Is = VarEnum.vtLong
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CLng(varTmp))
            Case Is = VarEnum.vtSingle
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CSng(varTmp))
            Case Is = VarEnum.vtDouble
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CDbl(varTmp))
            Case Is = VarEnum.vtString
                Call CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbLet, CStr(varTmp))
        End Select

    End If

Exit Sub
errTrat:

    If Tipo = ViewText Then
        Text1.Text = strOld
    Else
        Combo1.ListIndex = intOld
    End If
        
End Sub

Private Sub UserControl_EnterFocus()

    Label1.Visible = False

    If m_Tipo = ViewList Then
        Combo1.Visible = True
    Else
        Text1.Visible = True
    End If
    
    If Not (mHelpBox Is Nothing) Then
        mHelpBox.ControlString = mHelpString
    End If
    
End Sub

Private Sub UserControl_ExitFocus()

    Label1.Visible = True

    If m_Tipo = ViewList Then
        Combo1.Visible = False
        Label1.Text = Combo1.List(Combo1.ListIndex)
    Else
        Text1.Visible = False
        Label1.Text = Text1.Text
    End If
    
    If Not (mHelpBox Is Nothing) Then
        mHelpBox.ControlString = ""
    End If
    
End Sub

Private Sub UserControl_Initialize()
    mMemberID = -1
    SpTextbox.Target = Text1
    Combo1.Visible = False
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.SelText = PropBag.ReadProperty("SelText", "")
    Text1.Text = PropBag.ReadProperty("Text", "Text1")
    Text1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Combo1.List(Index) = PropBag.ReadProperty("List" & Index, "")
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    m_Tipo = PropBag.ReadProperty("Tipo", m_def_Tipo)

End Sub

Private Sub UserControl_Resize()

    UserControl.Height = Combo1.Height
    If Tipo = 0 Then
        Text1.Visible = True
        Combo1.Visible = False
    Else
        Text1.Visible = False
        Combo1.Visible = True
    End If
        
    With Text1
        .Left = 0
        .Top = 0
        .Width = UserControl.ScaleWidth - 25
        .Height = UserControl.ScaleHeight
    End With
    
    With Label1
        .Left = 0
        .Top = 0
        .Width = UserControl.ScaleWidth - 25
        .Height = UserControl.ScaleHeight
    End With
    
    With Combo1
        .Left = 0
        .Top = 0
        .Width = UserControl.ScaleWidth - 25
    End With
    
    
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Text1.SelText, "")
    Call PropBag.WriteProperty("Text", Text1.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipText", Text1.ToolTipText, "")
    Call PropBag.WriteProperty("List" & Index, Combo1.List(Index), "")
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
    Call PropBag.WriteProperty("Tipo", m_Tipo, m_def_Tipo)

End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Combo1.AddItem Item, Index
End Sub

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = Combo1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Combo1.List(Index) = New_List
    PropertyChanged "List"
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1.ListCount
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    Combo1.RemoveItem Index
End Sub

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = Combo1.Sorted
End Property

Public Property Get Tipo() As TpView
    Tipo = m_Tipo
End Property

Public Property Let Tipo(ByVal New_Tipo As TpView)
    
    m_Tipo = New_Tipo

    If m_Tipo = ViewText Then
        
        Select Case Interface.Members(MemberID).ReturnType
            
            'Case Is = VarEnum.vtBoolean
                'SpTextbox.Allowed =
            Case Is = VarEnum.vtByte
                SpTextbox.Allowed = ALLOW_NUMBERS
            Case Is = VarEnum.vtInteger
                SpTextbox.Allowed = ALLOW_NUMBERS
            Case Is = VarEnum.vtLong
                SpTextbox.Allowed = ALLOW_NUMBERS
            Case Is = VarEnum.vtSingle
                SpTextbox.Allowed = ALLOW_NUMBERS
            Case Is = VarEnum.vtDouble
                SpTextbox.Allowed = ALLOW_NUMBERS
            Case Is = VarEnum.vtString
                SpTextbox.Allowed = ALLOW_ALL
        End Select

    End If
    
    If m_Tipo = ViewList Then
        Combo1.Visible = False
        Label1.Text = Combo1.List(Combo1.ListIndex)
    Else
        Text1.Visible = False
        Label1.Text = Text1.Text
    End If
    
    PropertyChanged "Tipo"
    
End Property

Private Sub UserControl_InitProperties()
    m_Tipo = m_def_Tipo
End Sub

Public Property Get MemberID() As Integer
    MemberID = mMemberID
End Property

Public Property Let MemberID(ByVal NewValue As Integer)

Dim strTmp As String
Dim strSplit() As String
Dim Y As Integer

    mMemberID = NewValue
        
    
    Text = CallByName(mObjectTarget, Interface.Members(mMemberID).Name, VbGet)
    Tipo = ViewText
    
    If Interface.VTableInterface.Members(mMemberID).Parameters.Count = 1 Then
        
        Set vtInfo = Interface.VTableInterface.Members(mMemberID).Parameters(1).VarTypeInfo
        
        If Interface.Members(MemberID).ReturnType = VarEnum.vtBoolean Then
        
            Combo1.Enabled = True
            Combo1.Clear
            Combo1.AddItem "False"
            Combo1.AddItem "True"
            
                If Text = "False" Then
                    Combo1.ListIndex = 0
                Else
                    Combo1.ListIndex = 1
                End If
            
            Tipo = ViewList
            
        ElseIf Not (vtInfo.TypeInfo Is Nothing) Then
        
            If vtInfo.TypeInfo.TypeKind = TKIND_ENUM Then
                
                Combo1.Enabled = True
                Combo1.Clear
                
                For Y = 1 To vtInfo.TypeInfo.Members.Count
                    
                    strTmp = vtInfo.TypeInfo.Members(Y).Name
                    strTmp = strTmp & "="
                    strTmp = strTmp & vtInfo.TypeInfo.Members(Y).Value
                    Combo1.AddItem strTmp
                    
                Next
                
                For Y = 0 To Combo1.ListCount - 1
                    
                    strSplit = Split(Combo1.List(Y), "=")
                    
                    If Trim(Text) = Trim(strSplit(1)) Then
                        Combo1.ListIndex = Y
                    End If
                    
                Next
                
                Tipo = ViewList
                
                
            End If
            
        End If
        
    Else
        
        Text = "Unsuported"
        Text1.Enabled = False
        Tipo = ViewText
       
    End If
    
    
End Property

Public Property Get HelpString() As String
    
    HelpString = mHelpString
    
End Property

Public Property Let HelpString(NewValue As String)
    
    mHelpString = NewValue
    
End Property

'Public Property Get HelpBox() As HotTextBox
'
'    Set HelpBox = mHelpBox
'
'End Property

Public Sub HelpBox(NewValue As HotTextBox)

    Set mHelpBox = NewValue
    
End Sub
