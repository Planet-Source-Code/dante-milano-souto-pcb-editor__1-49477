VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' found this code in http://psprogramming.virtualave.net

Private WithEvents txtTarget As TextBox
Attribute txtTarget.VB_VarHelpID = -1

Private bSelectOnFocus As Boolean
Private bUppercase As Boolean
Private bLowercase As Boolean
Private eAllowed As eTextboxAllows

Public Enum eTextboxAllows
    ALLOW_ALL
    ALLOW_NORETURN
    ALLOW_ALPHANUMERIC
    ALLOW_NUMBERS
    ALLOW_TEXT
End Enum

Public Property Get Allowed() As eTextboxAllows
    Allowed = eAllowed
End Property

Public Property Let Allowed(Value As eTextboxAllows)
    eAllowed = Value
End Property

Public Property Get SelectOnFocus() As Boolean
    SelectOnFocus = bSelectOnFocus
End Property

Public Property Let SelectOnFocus(Value As Boolean)
    bSelectOnFocus = Value
End Property


Public Property Get Target() As TextBox
    Set Target = txtTarget
End Property

Public Property Let Target(Value As TextBox)
    Set txtTarget = Value
End Property


Public Property Get UpperCase() As Boolean
    UpperCase = bUppercase
End Property

Public Property Let UpperCase(Value As Boolean)
    bUppercase = Value
    If Value = True Then bLowercase = False
End Property
Public Property Get LowerCase() As Boolean
    LowerCase = bLowercase
End Property


Public Property Let LowerCase(Value As Boolean)
    bLowercase = Value
    If Value = True Then bUppercase = False
End Property

Private Sub Class_Initialize()
    '// Set Default Properties
    bSelectOnFocus = True
    bUppercase = False
    bLowercase = False
    eAllowed = ALLOW_ALL
End Sub

Private Sub Class_Terminate()
    '// Remove Link to Textbox
    Set txtTarget = Nothing
End Sub


Private Sub txtTarget_Change()
    Dim lLength As Long
    Dim lPos As Long
    
    If bUppercase Then
        '// Change to uppercase
        lPos = txtTarget.SelStart
        lLength = txtTarget.SelLength
        
        txtTarget.Text = UCase$(txtTarget.Text)
    
        txtTarget.SelStart = lPos
        txtTarget.SelLength = lLength
    ElseIf bLowercase Then
        '// Change to lowercase
        lPos = txtTarget.SelStart
        lLength = txtTarget.SelLength
        
        txtTarget.Text = LCase$(txtTarget.Text)
    
        txtTarget.SelStart = lPos
        txtTarget.SelLength = lLength
    End If
End Sub

Private Sub txtTarget_GotFocus()
    '// Select Text
    With txtTarget
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtTarget_KeyPress(KeyAscii As Integer)
    
Dim bAllowed As Boolean
    
    bAllowed = True
    
    If bUppercase Then
    
        '// Change to uppercase
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
    ElseIf bLowercase Then
    
        '// Change to lowercase
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
        
    End If
    
    Select Case eAllowed
        
        Case ALLOW_ALL
            '// Do nothing, everything's allowed
            bAllowed = True
            
        Case ALLOW_NORETURN
            '// Allow everything, but the return key
            If KeyAscii = 13 Then bAllowed = False
        
        Case ALLOW_ALPHANUMERIC
            '// Allow only alphanumeric characters
            If KeyAscii < 48 Then bAllowed = False
            If KeyAscii > 57 And KeyAscii < 65 Then bAllowed = False
            If KeyAscii > 90 And KeyAscii < 97 Then bAllowed = False
            If KeyAscii > 122 Then bAllowed = False
            If KeyAscii = 8 Then bAllowed = True
    
        Case ALLOW_NUMBERS
            '// Allow only numbers
            If KeyAscii < 48 Then bAllowed = False
            If KeyAscii > 57 Then bAllowed = False
            If KeyAscii = 44 Then bAllowed = True   ',
            If KeyAscii = 46 Then bAllowed = True    '.
            
        Case ALLOW_TEXT
            '// Allow only text
            bAllowed = True
            If KeyAscii < 65 Then bAllowed = False
            If KeyAscii > 90 And KeyAscii < 97 Then bAllowed = False
            If KeyAscii > 122 Then bAllowed = False
            If KeyAscii = 8 Then bAllowed = True
    
    End Select
    
    If bAllowed = False Then KeyAscii = 0
    
End Sub


