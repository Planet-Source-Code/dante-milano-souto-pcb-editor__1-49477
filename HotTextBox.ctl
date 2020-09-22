VERSION 5.00
Begin VB.UserControl HotTextBox 
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   2940
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   300
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin VB.Image imgCursor 
      Height          =   480
      Left            =   3420
      Picture         =   "HotTextBox.ctx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "HotTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************
'   MODIFICADO POR DANTE MILANO SOUTO
'*******************************************

'***************************************************
'Hot Text Box Control by Michael Pote (C) 2000
'***************************************************
'Most of the code in this control was generated
'by the ActiveX control interface wizard but
'the 2 main functions of this control are
'picMain_MouseMove and DrawText()

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type HotspotDefenition
    StartX  As Long
    EndX    As Long
    Y       As Long
End Type
Private HotList()   As HotspotDefenition

Private Const m_def_HighlightColor = 255
Private Const m_def_Highlight = 1
Private Const m_def_ForeColor = 0
Private Const m_def_Enabled = True
Private Const m_def_ControlString = "~Hot~Helpbox"
Private Const m_def_HotspotColor = 16711680
Private Const SRCCOPY = &HCC0020

Private m_HighlightColor    As OLE_COLOR
Private m_Highlight         As Boolean
Private m_ForeColor         As OLE_COLOR
Private m_Enabled           As Boolean
Private m_ControlString     As String
Private m_HotspotColor      As OLE_COLOR

Private OverHotSpot         As Boolean
Private HSIndex             As Long
Private HSCount             As Long

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event HotSpotClick(Index As Long)

Private Sub picMain_Click()
    
    RaiseEvent Click

End Sub

Private Sub CountHotSpot()
'Counts how many hotspots there are and
'dynamicaly resizes the hotlist() array

Dim I As Long
Dim C As Long
    
    C = 0
    For I = 1 To Len(m_ControlString)
        If mID$(m_ControlString, I, 1) = "~" Then C = C + 1
    Next
    
    Let HSCount = Int(C / 2)
    
    If HSCount > 0 Then
        ReDim HotList(1 To HSCount) As HotspotDefenition
    End If

End Sub

Private Sub UserControl_Initialize()
HSCount = 0
CountHotSpot
End Sub

Private Sub UserControl_Paint()
    
    Call DrawText
    
End Sub

Private Sub UserControl_Resize()
    
    With picMain
        .Left = 0
        .Top = 0
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With
    
    Call DrawText
    
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    
    BackColor = picPicture.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    picPicture.BackColor() = New_BackColor
    picMain.BackColor = picPicture.BackColor
    PropertyChanged "BackColor"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    
    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    
    Set Font = picMain.Font

End Property

Public Property Set Font(ByVal New_Font As Font)
    
    Set picMain.Font = New_Font
    PropertyChanged "Font"

End Property

Private Sub picMain_DblClick()
    
    RaiseEvent DblClick

End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If OverHotSpot = True Then RaiseEvent HotSpotClick(HSIndex)

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim I As Long

    RaiseEvent MouseMove(Button, Shift, X, Y)

    If HSCount = 0 Then Exit Sub
    
    picMain.MousePointer = 1
    OverHotSpot = False
    
    For I = 1 To UBound(HotList())
        If Y >= HotList(I).Y And Y <= HotList(I).Y + picMain.TextHeight("A") + 2 Then
            If X >= HotList(I).StartX And X <= HotList(I).EndX Then
                picMain.MouseIcon = imgCursor.Picture
                picMain.MousePointer = 99
                HSIndex = I
                OverHotSpot = True
            End If
        End If
    Next
    
    Call DrawText

End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    
    Call DrawText
    
End Sub

Public Property Get ControlString() As String
Attribute ControlString.VB_Description = "Specifies the text in the control, and where the hotspot region and pictures go. ~ - Hotspotstart ` - Hotspot end * - picture "
    
    ControlString = m_ControlString
    Call DrawText

End Property

Public Property Let ControlString(ByVal New_ControlString As String)
    
    m_ControlString = New_ControlString
    PropertyChanged "ControlString"
    Call CountHotSpot
    Call DrawText

End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    
    Set Picture = picPicture.Picture

End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    
    Set picPicture.Picture = New_Picture
    PropertyChanged "Picture"

End Property

Public Property Get HotspotColor() As OLE_COLOR
    
    HotspotColor = m_HotspotColor

End Property

Public Property Let HotspotColor(ByVal New_HotspotColor As OLE_COLOR)
    
    m_HotspotColor = New_HotspotColor
    PropertyChanged "HotspotColor"

End Property

Private Sub UserControl_InitProperties()
    
    m_Enabled = m_def_Enabled
    'm_ControlString = m_def_ControlString
    m_HotspotColor = m_def_HotspotColor
    m_ForeColor = m_def_ForeColor
    m_Highlight = m_def_Highlight
    m_HighlightColor = m_def_HighlightColor

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picPicture.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set picMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ControlString = PropBag.ReadProperty("ControlString", m_def_ControlString)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_HotspotColor = PropBag.ReadProperty("HotspotColor", m_def_HotspotColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    picMain.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_Highlight = PropBag.ReadProperty("Highlight", m_def_Highlight)
    m_HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", picPicture.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", picMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ControlString", m_ControlString, m_def_ControlString)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("HotspotColor", m_HotspotColor, m_def_HotspotColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderStyle", picMain.BorderStyle, 1)
    Call PropBag.WriteProperty("Highlight", m_Highlight, m_def_Highlight)
    Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)

End Sub

Private Sub DrawText()

Dim C       As Long
Dim CurHS   As Long
Dim Update  As Boolean
Dim HS      As Boolean
Dim I       As Integer
Dim Lett    As String
    
    If HSCount = 0 Then CountHotSpot
    
    picMain.Cls
    picMain.CurrentX = 2
    picMain.CurrentY = 2
    HS = False
    
    For I = 1 To Len(m_ControlString) 'Loop through all
    
        Lett = mID$(m_ControlString, I, 1)
        
        If Lett = "~" Then
            HS = (HS = False)
            
            If HS Then
            
                If HSCount >= C + 1 Then
                C = C + 1
                
                    If HotList(C).StartX <> picMain.CurrentX Then
                        Update = True
                        HotList(C).StartX = picMain.CurrentX
                        HotList(C).Y = picMain.CurrentY
                    Else
                        Update = False
                    End If
                    
                End If
                
            End If
        
            If HS = False Then
            
                If HSCount >= C And Update Then HotList(C).EndX = picMain.CurrentX
                    
                    picPicture.Cls
                    
                End If
                
                GoTo NextLett
                
            End If
        
        If Lett = "*" Then 'Picture
            
            BitBlt picMain.hDC, picMain.CurrentX + 3, picMain.CurrentY, picPicture.ScaleWidth, picPicture.ScaleHeight, picPicture.hDC, 0, 0, SRCCOPY: picMain.CurrentX = picMain.CurrentX + picPicture.Width + 3: GoTo NextLett
            
            If HS Then
                
                picMain.Line (picMain.CurrentX + 2, picMain.CurrentY)-(picMain.CurrentX + (picpicure.ScaleWidth - 1), picMain.CurrentY + (picPicture.ScaleHeight - 1)), m_HotspotColor, B
            
            End If
            
        End If
        
        If HS Then
        
            If C = HSIndex And OverHotSpot = True Then
                picMain.ForeColor = m_HighlightColor
            Else
                picMain.ForeColor = m_HotspotColor
            End If
        
            picMain.FontUnderline = True

        Else
        
            picMain.ForeColor = m_ForeColor
            picMain.FontUnderline = False
        
        End If
        
        If picMain.CurrentX + picMain.TextWidth(Lett) > picMain.ScaleWidth Then
            
            picMain.CurrentX = 2
            picMain.CurrentY = picMain.CurrentY + picMain.TextHeight(Lett) + 2
        
        End If
        
        picMain.Print Lett;
        
NextLett:

    Next
    
    picMain.Refresh
    
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    
    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"

End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    
    BorderStyle = picMain.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    
    picMain.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"

End Property

Public Property Get Highlight() As Boolean
Attribute Highlight.VB_Description = "Sets whether the hot spot lights up when the  mouse is over it."
    
    Highlight = m_Highlight

End Property

Public Property Let Highlight(ByVal New_Highlight As Boolean)
    
    m_Highlight = New_Highlight
    PropertyChanged "Highlight"

End Property

Public Property Get HighlightColor() As OLE_COLOR
    
    HighlightColor = m_HighlightColor

End Property

Public Property Let HighlightColor(ByVal New_HighlightColor As OLE_COLOR)
    
    m_HighlightColor = New_HighlightColor
    PropertyChanged "HighlightColor"

End Property

