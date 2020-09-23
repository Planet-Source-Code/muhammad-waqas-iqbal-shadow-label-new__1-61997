VERSION 5.00
Begin VB.UserControl ShadowLabel1 
   Alignable       =   -1  'True
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ScaleHeight     =   285
   ScaleWidth      =   825
   ToolboxBitmap   =   "ShadowLabel1.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Label Label1 
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
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   585
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   25
      TabIndex        =   1
      Top             =   10
      Width           =   585
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ShadowLabel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long
Dim m_LabelStyle As Boolean
Public Enum State_b
    Normal_ = 0
    Default_ = 1
End Enum
Const new_def_LabelStyle = 0

Dim m_State As State_b
Dim m_Font As Font
Dim m_BackStyle As Boolean
Const m_Def_State = State_b.Normal_
Public Enum LabelStyle2    ' button styles
    ThreeD = 0
    Shadow = 1
    End Enum
Private Type POINT_API
    x As Long
    Y As Long
End Type
Const m_def_PasswordChar = ""
Dim m_PasswordChar As String
Private WithEvents F              As Form
Attribute F.VB_VarHelpID = -1
Dim s As Integer
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event Change()
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = Label2.ForeColor
End Property

Public Property Let ShadowColor(ByVal ShadowColor As OLE_COLOR)
    Label2.ForeColor() = ShadowColor
    PropertyChanged "ShadowColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = Label2.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Label1.BackColor() = New_BackColor
    Label2.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Private Sub Shodowlabel1_Change()
RaiseEvent Change
End Sub
Public Property Get FontBold() As Boolean
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Label1.FontBold() = New_FontBold
    Label2.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get LabelStyle() As Boolean
    LabelStyle = m_LabelStyle
End Property

Public Property Let LabelStyle(ByVal new_LabelStyle As Boolean)
m_LabelStyle = new_LabelStyle
PropertyChanged "LabelStyle"
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Label1.FontItalic() = New_FontItalic
    Label2.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property
Public Property Get Font() As Font
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set Label2.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
   Caption = Label1.Caption
    End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    Label2.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub


Private Sub F_Load()
If m_LabelStyle = True Then
Label2.ForeColor = vbBlack
Else
Label2.ForeColor = Shadow
End If
End Sub

Private Sub UserControl_Resize()
'Label1.Left = 40
'Label1.Top = 40
'Label1.Width = UserControl.ScaleWidth - 90
'Label1.Height = UserControl.ScaleHeight - 90
'Label2.Left = 0
'Label2.Top = 0
'Label2.Width = UserControl.ScaleWidth
'Label2.Height = UserControl.ScaleHeight
If UserControl.Height < 180 Then UserControl.Height = 290

End Sub
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_PasswordChar = m_def_PasswordChar
    MsgBox "Created By: Muhammad Waqas Iqbal" & vbCrLf & "Please register your Shadow Label for free." & vbCrLf & "Please log on to http://blueapple.o-f.com", vbInformation, "Shadow Label 1.0"
 new_LabelStyle = new_def_LabelStyle
End Sub
Public Property Get FontName() As String
    FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Label1.FontName() = New_FontName
     Label2.FontName() = New_FontName
    PropertyChanged "FontName"
End Property



Public Property Get FontSize() As Single
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Label1.FontSize() = New_FontSize
    Label2.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get WordWrap() As Boolean
    WordWrap = Label1.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    Label1.WordWrap() = New_WordWrap
     Label2.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = Label1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Label1.FontStrikethru() = New_FontStrikethru
     Label2.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = Label1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Label1.FontUnderline() = New_FontUnderline
    Label2.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
'Public Property Get AutoSize() As Boolean
'    AutoSize = Label1.AutoSize
'End Property

'Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
'    Label1.AutoSize() = New_AutoSize
'    Label2.AutoSize() = New_AAutoSize
'    PropertyChanged "AutoSize"
'End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Label1.Caption = PropBag.ReadProperty("Caption", "Label1")
    Label1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    cadre.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    Label1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Label1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Label1.FontName = PropBag.ReadProperty("FontName", "arial")
    Label1.FontSize = PropBag.ReadProperty("FontSize", 8)
    Label1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Label1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Label1.WordWrap = PropBag.ReadProperty("WordWrap", 0)
    'Label1.AutoSize = PropBag.ReadProperty("AutoSize", 1)
    Set Label1.Font = PropBag.ReadProperty("Font", "Arial")
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    
    
    Label2.Caption = PropBag.ReadProperty("Caption", "Label1")
    Label2.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label2.ForeColor = PropBag.ReadProperty("ShadowColor", &H808080)
    cadre.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    Label2.FontBold = PropBag.ReadProperty("FontBold", 0)
    Label2.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Label2.FontName = PropBag.ReadProperty("FontName", "arial")
    Label2.FontSize = PropBag.ReadProperty("FontSize", 8)
    Label2.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Label2.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Label2.WordWrap = PropBag.ReadProperty("WordWrap", 0)
    'Label2.AutoSize = PropBag.ReadProperty("AutoSize", 1)
    Set Label2.Font = PropBag.ReadProperty("Font", "Arial")
    Label2.Alignment = PropBag.ReadProperty("Alignment", 0)

    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("FontName", Label1.FontName, "arial")
    Call PropBag.WriteProperty("FontSize", Label1.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", Label1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", Label1.FontUnderline, 0)
    Call PropBag.WriteProperty("FontBold", Label1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Label1.FontItalic, 0)
    Call PropBag.WriteProperty("BorderWidth", cadre.BorderWidth, 1)
    Call PropBag.WriteProperty("LineColor", cadre.BorderColor, &H80000005)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Caption")
    Call PropBag.WriteProperty("Caption", Label2.Caption, "Caption")
    Call PropBag.WriteProperty("BackColor", Label1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", Label1.Font, "Arial")
    Call PropBag.WriteProperty("WordWrap", Label1.WordWrap, 0)
    'Call PropBag.WriteProperty("AutoSize", Label1.AutoSize, 0)

    
    'label2
    Call PropBag.WriteProperty("Alignment", Label2.Alignment, 0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("FontName", Label2.FontName, "arial")
    Call PropBag.WriteProperty("FontSize", Label2.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", Label2.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", Label2.FontUnderline, 0)
    Call PropBag.WriteProperty("FontBold", Label2.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Label2.FontItalic, 0)
    Call PropBag.WriteProperty("BorderWidth", cadre.BorderWidth, 1)
    Call PropBag.WriteProperty("LineColor", cadre.BorderColor, &H80000005)
    Call PropBag.WriteProperty("Caption", Label2.Caption, "Caption")
    Call PropBag.WriteProperty("Caption", Label2.Caption, "Caption")
    Call PropBag.WriteProperty("BackColor", Label2.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShadowColor", Label2.ForeColor, &H808080)
    Call PropBag.WriteProperty("Font", Label2.Font, "Arial")
    Call PropBag.WriteProperty("WordWrap", Label2.WordWrap, 0)
    'Call PropBag.WriteProperty("AutoSize", Label2.AutoSize, 0)

    End Sub


