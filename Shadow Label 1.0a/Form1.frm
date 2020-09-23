VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{80B8EDD8-C50A-485A-AF8B-24DCE0FEF5FE}#20.0#0"; "SHADOW LABEL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shadow Label"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   3
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
   End
   Begin VB.CommandButton Command6 
      Caption         =   "3D Text"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Select Font"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select Color"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select Color"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Color"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel16 
      Height          =   855
      Left            =   1800
      TabIndex        =   13
      Top             =   3480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1508
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "Font"
      Caption         =   "Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "Font"
      Caption         =   "Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel15 
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      FontName        =   "Monotype Corsiva"
      FontSize        =   14.25
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Caption         =   "Created By: Muhammad Waqas Iqbal"
      Caption         =   "Created By: Muhammad Waqas Iqbal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Monotype Corsiva"
      FontSize        =   14.25
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Caption         =   "Created By: Muhammad Waqas Iqbal"
      Caption         =   "Created By: Muhammad Waqas Iqbal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel14 
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "ForeColor"
      Caption         =   "ForeColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "ForeColor"
      Caption         =   "ForeColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel13 
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "ShadowColor"
      Caption         =   "ShadowColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "ShadowColor"
      Caption         =   "ShadowColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel12 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "BackColor"
      Caption         =   "BackColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "BackColor"
      Caption         =   "BackColor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel11 
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1085
      FontName        =   "Comic Sans MS"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "Shadow Label ver1.0"
      Caption         =   "Shadow Label ver1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Comic Sans MS"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "Shadow Label ver1.0"
      Caption         =   "Shadow Label ver1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel17 
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "3D Label"
      Caption         =   "3D Label"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   24
      FontBold        =   -1  'True
      Caption         =   "3D Label"
      Caption         =   "3D Label"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShadowLabel.ShadowLabel1 ShadowLabel18 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      FontName        =   "Monotype Corsiva"
      FontSize        =   14.25
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Caption         =   "Email: mwaqasiq007@hotmail.com"
      Caption         =   "Email: mwaqasiq007@hotmail.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
      FontName        =   "Monotype Corsiva"
      FontSize        =   14.25
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      Caption         =   "Email: mwaqasiq007@hotmail.com"
      Caption         =   "Email: mwaqasiq007@hotmail.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     Shadow Label ver1.0           '
'Created By:Muhammad Waqas Iqbal    '
'WebSite: htt://blueapple.o-f.com   '
'         http://softmart.cjb.net   '
'Email: mwaqasiq007@hotmail.com     '
'       waqasiq@gmail.com           '
'''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
On Error Resume Next
cd.ShowColor
ShadowLabel12.BackColor = cd.Color

End Sub

Private Sub Command2_Click()
On Error Resume Next
cd.ShowColor
ShadowLabel13.ShadowColor = cd.Color
End Sub

Private Sub Command3_Click()
On Error Resume Next
cd.ShowColor
ShadowLabel14.ForeColor = cd.Color
End Sub

Private Sub Command4_Click()
On Error Resume Next
cd.ShowFont
ShadowLabel16.FontName = cd.FontName
ShadowLabel16.FontItalic = cd.FontItalic
ShadowLabel16.FontBold = cd.FontBold
ShadowLabel16.FontUnderline = cd.FontUnderline
ShadowLabel16.FontStrikethru = cd.FontStrikethru
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
On Error Resume Next
ShadowLabel17.ShadowColor = vbBlack
End Sub
