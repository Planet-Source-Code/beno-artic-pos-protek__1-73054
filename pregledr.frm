VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form pregledr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PREGLED RACUNOV"
   ClientHeight    =   10500
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton7 
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "TIPKOVNICA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   615
      Left            =   1560
      TabIndex        =   26
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "A4-RACUN "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   855
      Left            =   13080
      TabIndex        =   25
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "A4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16744576
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":0038
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   1680
      TabIndex        =   23
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "RACUNI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":0054
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Naèin plaèila"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   9240
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cdd 
      Left            =   5520
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8640
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   120
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59834369
      CurrentDate     =   40509
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SINHRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13080
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13080
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   5640
      TabIndex        =   11
      Top             =   7920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   2880
      TabIndex        =   10
      Top             =   9960
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   9480
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   9000
      Width           =   7335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2880
      TabIndex        =   7
      Top             =   8520
      Width           =   7335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "INTERNA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9960
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "KARTICA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "GOTOVINA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8760
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   5400
      TabIndex        =   3
      Top             =   1440
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "pregledr.frx":0070
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ZAPRI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10920
      TabIndex        =   1
      Top             =   9360
      Width           =   1935
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "POSODOBI VNOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10920
      TabIndex        =   0
      Top             =   7920
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59834369
      CurrentDate     =   40509
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   5400
      TabIndex        =   24
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "ZAKLJUCKI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":0076
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   615
      Left            =   0
      TabIndex        =   27
      Top             =   7680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PARAGONSKI- RACUN "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":0092
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton6 
      Height          =   615
      Left            =   2880
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "STORNO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   8438015
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":00AE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton8 
      Height          =   615
      Left            =   9120
      TabIndex        =   31
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PO UPORAB - NIKU"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":00CA
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   1
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton9 
      Height          =   615
      Left            =   11400
      TabIndex        =   32
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PO GRUPI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pregledr.frx":00E6
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label5 
      Height          =   135
      Left            =   9120
      TabIndex        =   30
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "ZAPOSLEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "DO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "OD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Davcna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   7920
      Width           =   1095
   End
End
Attribute VB_Name = "pregledr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOSIZE As Long = &H1

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Check1_Click()
Check2.Value = 0
Check3.Value = 0
End Sub

Private Sub Check2_Click()
Check1.Value = 0
Check3.Value = 0
End Sub

Private Sub Check3_Click()
Check1.Value = 0
Check2.Value = 0
End Sub

Private Sub Combo1_Change()
xosw
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
xosw
End Sub

Private Sub Command1_Click()

cdd.Copies = 1
cdd.PrinterDefault = True
cdd.ShowPrinter
Printer.Print Text1.text
Printer.EndDoc
End Sub
Sub aabbb()
Dim strfile As String
strfile = App.path & "\natr.prn"
SaveFileFromTB Text1, strfile, False
Dim P As Printer

End Sub
Private Sub Command2_Click()
Me.List1.SetFocus

Sendkeys "{UP}"
End Sub

Private Sub Command3_Click()
Me.List1.SetFocus

Sendkeys "{DOWN}"
End Sub

Private Sub Command4_Click()
Dim idstr, nall As String
idstr = Getnazi("select ime from po where dav='" & Me.Text2(4).text & "'")
nall = Getnazi("select nasl from po where dav='" & Me.Text2(4).text & "'")
If idstr = "" Then
nall = Getnazi("select nasl from fozD where dav='" & Me.Text2(4).text & "'")
idstr = Getnazi("select ime from fozD where dav='" & Me.Text2(4).text & "'")
End If
Me.Text2(0).text = Left(idstr, 40)
Me.Text2(1).text = LTrim(Mid(idstr, 41))
Me.Text2(2).text = Left(nall, 40)
Me.Text2(3).text = LTrim(Mid(nall, 41))

End Sub

Private Sub Command5_Click()
Dim rst As New ADODB.Recordset
Dim qaha As String
If Left(Me.List1.text, 6) <> "" Then
qaha = "select id_dok,sum(znes) as znes from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "' group by id_dok order by id_dok desc"
'MsgBox (qaha)
rst.Open qaha, myConection, adOpenDynamic, adLockOptimistic
rst.MoveFirst
If Not rst.EOF Then
nacpla.odprnac "PA" & rst.Fields("id_dok"), rst.Fields("znes")
ohonac = 1
End If
End If

End Sub

Private Sub DTPicker1_Change()
xosw
End Sub

Private Sub DTPicker2_Change()
xosw
End Sub

Private Sub Form_Load()
ReSizeForm Me
If rs.State = 1 Then rs.Close
rs.Open "select * from users", myConection, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Me.Combo1.clear
Me.Combo1.AddItem ""
Do While Not rs.EOF
Me.Combo1.AddItem rs.Fields("up") & " - " & rs.Fields("username1")
rs.MoveNext
Loop
Me.DTPicker1.Value = Date
Me.DTPicker2.Value = Date
Dim PE As Printer
Me.Text1.Height = Me.LaVolpeButton5.Top - Me.Text1.Top
LaVolpeButton1_Click

If Getnumb("select nivo from users where username1='" & UPORABNIK & "'") = 1 Then
Me.LaVolpeButton2.Visible = True
Me.LaVolpeButton8.Visible = True
Me.LaVolpeButton9.Visible = True
Else
Me.LaVolpeButton2.Visible = False
Me.LaVolpeButton8.Visible = False
Me.LaVolpeButton9.Visible = False
End If

xosw


End Sub
Sub xosw()
Dim qpog As String
qpog = ""
If Me.Combo1.text <> "" Then
qpog = " and uporabnik='" & Left(Me.Combo1.text, 2) & "'"
End If
Dim das, des
Dim dodajwh, dwod, dwdo As String
dodajwh = ""
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
des = Format(Me.DTPicker2.Value, "dd.mm.yyyy")
dwod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
dwdo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
dodajwh = " and datum between #" & dwod & "# AND #" & dwdo & "# "
Dim rst As New ADODB.Recordset
Dim qaha As String
Dim xni As Integer
xni = Getnumb("select nivo from users where username1='" & UPORABNIK & "'")
If Me.LaVolpeButton1.BackColor = 8421631 Then
qaha = "select id_dok,min(datum) as datum,sum(znes) as znes,max(chk_fix) as chk_fix from nabasif where tip_dok='PA' " & qpog & dodajwh & " group by id_dok order by id_dok desc"
Else
qaha = "select  datum,sum(znes) as znes from nabasif where tip_dok='PA' " & qpog & dodajwh & " group by datum order by datum desc"
End If
'MsgBox (qaha)
Me.Label5.Caption = qaha
If rst.State = 1 Then rst.Close
rst.Open qaha, myConection, adOpenDynamic, adLockOptimistic
If rst.EOF Then
'MsgBox ("Ne najdem nobenega podatka ki ustreza pogoju!")
Exit Sub
End If
rst.MoveFirst
List1.clear
Dim ssst As Integer
ssst = 0
Do While Not rst.EOF
If Me.LaVolpeButton1.BackColor = 8421631 Then
If xni = 1 Then
ssst = 1
Else
ssst = ssst + 1
End If
If ssst < 6 Then
List1.AddItem Left(rst.Fields("id_dok"), 6) & "  " & Format(rst.Fields("datum"), "dd/mm/yyyy") & "    " & levi_pres(FormatNumber(rst.Fields("znes"), 2), 10) & "  " & rst.Fields("chk_fix")
End If
Else
List1.AddItem Format(rst.Fields("datum"), "dd/mm/yyyy") & "    " & FormatNumber(rst.Fields("znes"), 2)
End If

rst.MoveNext
Loop
End Sub

Private Sub Label5_Click()
MsgBox (Me.Label5.Caption)
End Sub

Private Sub LaVolpeButton1_Click()

Me.LaVolpeButton2.BackColor = 14215660
Me.LaVolpeButton1.BackColor = 8421631
Me.LaVolpeButton9.BackColor = 14215660
Me.LaVolpeButton8.BackColor = 14215660
Me.Height = Me.CancelButton.Top + (Me.CancelButton.Height * 2)
Me.LaVolpeButton6.Visible = True
Me.LaVolpeButton4.Caption = "A4 - RAÈUN"
Me.LaVolpeButton5.Caption = "PARAGONSKI - RAÈUN"
xosw
End Sub

Private Sub LaVolpeButton2_Click()

Me.LaVolpeButton1.BackColor = 14215660
Me.LaVolpeButton2.BackColor = 8421631
Me.LaVolpeButton9.BackColor = 14215660
Me.LaVolpeButton8.BackColor = 14215660
Me.Height = Me.LaVolpeButton5.Top + (Me.LaVolpeButton5.Height * 2)
Me.LaVolpeButton6.Visible = False
Me.LaVolpeButton4.Caption = "A 4 - ZAKLJUCEK"
Me.LaVolpeButton5.Caption = "PARAGON - ZAKLJUCEK"
xosw

End Sub

Private Sub LaVolpeButton3_Click()
xrep = "1"
printsql = "select * from mada"
PRINTREP = "report1.rpt"
'Form7.Show vbModal

End Sub

Private Sub LaVolpeButton4_Click()
If Left(Me.LaVolpeButton4.Caption, 2) = "A4" Then
    If Left(Me.List1.text, 6) <> "" Then
    PRINTSNAP "Z_CENAMI", "tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'"
    End If
Else
    If Left(Me.List1.text, 6) <> "" Then
    
    PRINTSNAP "vsizakl", " datum=#" & Left(Me.List1.text, 10) & "#"
    End If

End If
End Sub

Private Sub LaVolpeButton5_Click()

If Left(Me.List1.text, 6) <> "" Then
cdd.Copies = 1
cdd.PrinterDefault = True
cdd.ShowPrinter
Printer.Print Text1.text
Printer.EndDoc
End If
End Sub

Private Sub LaVolpeButton6_Click()
Dim xxiddo, stracun As String

If Left(Me.List1.text, 6) <> "" Then
stracun = Left(Me.List1.text, 6)
If Getnazi("select chk_fix from nabasif where tip_dok='PA' andd id_dok='" & stracun & "'") = "S" Then
 MsgBox ("Ta raèun je že storniran! stornacija ni mogoèa!")
Else
 xxiddo = novast(Val(Getnazi("select max(id_dok) from nabasif where tip_dok='PA'")) + 1, 6)
    If rs.State = 1 Then rs.Close
    rs.Open "insert into glavna (tip_dok,id_dok,opis)values('PA','" & xxiddo & "','STORNO')", myConection, adOpenDynamic, adLockOptimistic
   Dim aaaa, stnr As String
 Dim dass
    Dim datum As String
    
dass = Format(Now, "dd.mm.yyyy hh:mm:ss")
datum = Left(dass, 2) & "." & Mid(dass, 4, 2) & "." & Mid(dass, 7, 4)

aaaa = "insert into nabasif select 'S' as chk_fix,'" & datum & "' as datum,'PA' as tip_dok,'" & xxiddo & "' as id_dok,(tip_dok+id_dok) as kopija,sifra,naziv,faktor,skl,pozicija,'" & Getnazi("select up from users where username1='" & UPORABNIK & "'") & "' as uporabnik,cena,(kol*-1) as kol,pop, x, (znes*-1) as znes," & Pblagajna & " as stdok  from nabasif  where tip_dok='PA' and id_dok='" & stracun & "' order by pozicija"
'MsgBox (aaaa)
myConection.Execute (aaaa)
aaaa = "insert into nacplac select 'PA" & xxiddo & "' as dokument,sifra,znesek*-1 as znesek,'" & datum & "' as datum from nacplac  where dokument='PA" & stracun & "'"
myConection.Execute (aaaa)
aaaa = "update nabasif set chk_fix='S',kopija='PA" & xxiddo & "' where tip_dok='PA' and id_dok='" & stracun & "'"
myConection.Execute (aaaa)
If rs.State = 1 Then rs.Close
rs.Open "select * from nabasif where tip_dok='PA' and id_dok='" & xxiddo & "'", myConection, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
aaaa = "update mada set madazalo=madazalo+" & rs.Fields("kol") & " where madasifr='" & rs.Fields("sifra") & "'"
'MsgBox (aaaa)
myConection.Execute (aaaa)
rs.MoveNext
Loop

MsgBox ("PA" & stracun & " je bil storniran z PA" & xxiddo & "!!!")
xosw
End If
End If
End Sub

Private Sub LaVolpeButton7_Click()
FKeyboard.Show vbModal
End Sub

Private Sub LaVolpeButton8_Click()
Me.LaVolpeButton1.BackColor = 14215660
Me.LaVolpeButton2.BackColor = 14215660
Me.LaVolpeButton9.BackColor = 14215660
Me.LaVolpeButton8.BackColor = 8421631
Me.Height = Me.LaVolpeButton5.Top + (Me.LaVolpeButton5.Height * 2)
Me.LaVolpeButton6.Visible = False
Me.LaVolpeButton4.Caption = "A 4 - ZAKLJUCEK"
Me.LaVolpeButton5.Caption = "PARAGON - ZAKLJUCEK"
xosw

End Sub

Private Sub LaVolpeButton9_Click()
Me.LaVolpeButton1.BackColor = 14215660
Me.LaVolpeButton2.BackColor = 14215660
Me.LaVolpeButton8.BackColor = 14215660
Me.LaVolpeButton9.BackColor = 8421631
Me.Height = Me.LaVolpeButton5.Top + (Me.LaVolpeButton5.Height * 2)
Me.LaVolpeButton6.Visible = False
Me.LaVolpeButton4.Caption = "A 4 - ZAKLJUCEK"
Me.LaVolpeButton5.Caption = "PARAGON - ZAKLJUCEK"
xosw

End Sub

Private Sub List1_Click()
If Me.LaVolpeButton1.BackColor = 8421631 Then
    racuni
Else
    If Me.LaVolpeButton2.BackColor = 8421631 Then
     zaklju
    Else
        If Me.LaVolpeButton9.BackColor = 8421631 Then
            zapog
        Else
        zapou
        End If
    End If
End If
End Sub
Sub zapou()
Text1.text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & _
vbCrLf & vbCrLf
'vbCrLf & "Time:" & vbCrLf
If Me.Combo1.text = "" Then
Text1.text = Text1.text & "PO GRUPI Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Else
Text1.text = Text1.text & "PO GRUPI Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Text1.text = Text1.text & "PREGLED ZAPOSLENEGA: " & Me.Combo1.text & vbCrLf

End If
   If rs.State = 1 Then rs.Close
 Dim das, des
das = Left(Me.List1.text, 10)

dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
If Me.Combo1.text = "" Then
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
Else
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and uporabnik='" & Left(Combo1.text, 2) & "' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
End If
If Not rs.EOF Then
rs.MoveFirst
End If

Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim orr As Double
Dim hrana As Double
Dim pijaca As Double
Dim cig As Double
Dim vsto As Double

Dim kart As Double
Dim gotov As Double
gotov = 0
kart = 0
zne = 0
ddva = 0

ddvb = 0
hrana = 0
pijaca = 0
cig = 0

vsto = 0
Dim davek As Double
Dim vrsta As Integer
Do While Not rs.EOF
If rs.Fields("sifrapart") <> 0 Then
orr = orr + rs.Fields(0)
End If
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

If IsNull(rs.Fields("sifra")) Then
rs.Fields("sifra") = 0
rs.Update
End If
vrsta = Getnumb("select madagrup from mada where madasifr='" & rs.Fields("sifra") & "'")
If Getnumb("select vr from grupa where sifra=" & vrsta) = 0 Then
pijaca = pijaca + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 1 Then
hrana = hrana + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 2 Then
cig = cig + rs.Fields(0)
End If

If Getnumb("select vr from grupa where sifra=" & vrsta) = 4 Then
vsto = vsto + rs.Fields(0)
End If



zne = zne + rs.Fields(0)
If Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'") = "20" Then
ddva = ddva + rs.Fields(0)
End If
If Replace(Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'"), ",", ".") = "8.5" Then
ddvb = ddvb + rs.Fields(0)
End If

rs.MoveNext
Loop
Dim aa As Double
Dim bb As Double
'aa = Format(Me.Text1.text, "0.00")
'bb = Format(Me.Text2.text, "0.00")

If rs.State = 1 Then rs.Close
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic

Dim ee As Long

Dim ff As Long
ee = 0
ff = 0

If Not rs.EOF Then

If IsNull(rs.Fields(0)) Then
ee = 0
Else
ee = rs.Fields(0)
End If
If IsNull(rs.Fields(1)) Then

ff = 0
Else
ff = rs.Fields(1)
End If

If Me.Combo1.text <> "" Then
Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Prijava ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ee, 6) & "'") & vbCrLf
    Text1.text = Text1.text & "Odjava  ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ff, 6) & "'") & vbCrLf
    
End If

    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Zacetna st.rac. : " & ee & vbCrLf
    Text1.text = Text1.text & "Konèna st.rac.  : " & ff & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj izdano raèunov : " & ff - ee + 1 & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj znesek prodaje : " & FormatNumber(zne, 2) & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
   
   Text1.text = Text1.text & Getnacindan(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf
 Text1.text = Text1.text & vbCrLf
 
 Text1.text = Text1.text & "U P O R A B N I K I====================" & vbCrLf
   
If rs.State = 1 Then rs.Close
If Me.Combo1.text = "" Then
rs.Open "SELECT UPORABNIK,SUM(KOL) AS KOL,SUM(ZNES) AS ZNES FROM NABASIF where  tip_dok='PA'  and datum=#" & dod & "# GROUP BY UPORABNIK", myConection, adOpenDynamic, adLockOptimistic

Else
rs.Open "SELECT UPORABNIK,SUM(KOL) AS KOL,SUM(ZNES) AS ZNES FROM NABASIF where  tip_dok='PA' and uporabnik='" & Left(Combo1.text, 2) & "' and datum=#" & dod & "# GROUP BY UPORABNIK", myConection, adOpenDynamic, adLockOptimistic
'rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and uporabnik='" & Left(Combo1.Text, 2) & "' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
End If
If Not rs.EOF Then
rs.MoveFirst
Dim grrr As Long
grrr = 0
Do While Not rs.EOF

   Text1.text = Text1.text & presled(Trim(rs.Fields("UPORABNIK")) & "-" & Getnazi("SELECT USERNAME1 FROM USERS WHERE UP='" & rs.Fields("UPORABNIK") & "'"), 21) & levi_pres(FormatNumber(rs.Fields("kol"), 2), 9) & levi_pres(FormatNumber(rs.Fields("znes"), 2), 9) & vbCrLf
        
rs.MoveNext
Loop
End If
     If ddva <> 0 Or ddvb <> 0 Then
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
        Text1.text = Text1.text & "Osnova  DDV        Znesek DDV   Vrednost" & vbCrLf
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
       
        If ddva <> 0 Then
    
         Text1.text = Text1.text & presled(Format(ddva / 1.2, "standard"), 8) & "20 %" & levi_pres(Format(ddva - (ddva / 1.2), "standard"), 14) & levi_pres(Format(ddva, "standard"), 14) & vbCrLf
   
        End If
        If ddvb <> 0 Then
         Text1.text = Text1.text & presled(Format(ddvb / 1.085, "standard"), 8) & "8,5 %" & levi_pres(Format(ddvb - (ddvb / 1.085), "standard"), 14) & levi_pres(Format(ddvb, "standard"), 14) & vbCrLf
        End If
       Text1.text = Text1.text & "----------------------------------------" & vbCrLf
    End If
If Me.Combo1.text <> "" Then
Text1.text = Text1.text & Getnacindancig(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf
End If
    Text1.text = Text1.text & vbCrLf & vbCrLf
   
      Text1.text = Text1.text & vbCrLf
    If Getnazi("select konec1 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec1 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec2 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
   Text1.text = Text1.text & Getnazi("select konec3 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec4 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
     Text1.text = Text1.text & Getnazi("select konec5 from oblikar") & vbCrLf
    End If
     Text1.text = Text1.text & vbCrLf & vbCrLf & vbCrLf & vbCrLf

End If

End Sub

Sub zapog()

Text1.text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & _
vbCrLf & vbCrLf
'vbCrLf & "Time:" & vbCrLf
If Me.Combo1.text = "" Then
Text1.text = Text1.text & "PO GRUPI Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Else
Text1.text = Text1.text & "PO GRUPI Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Text1.text = Text1.text & "PREGLED ZAPOSLENEGA: " & Me.Combo1.text & vbCrLf

End If
   If rs.State = 1 Then rs.Close
 Dim das, des
das = Left(Me.List1.text, 10)

dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
If Me.Combo1.text = "" Then
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
Else
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and uporabnik='" & Left(Combo1.text, 2) & "' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
End If
If Not rs.EOF Then
rs.MoveFirst
End If

Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim orr As Double
Dim hrana As Double
Dim pijaca As Double
Dim cig As Double
Dim vsto As Double

Dim kart As Double
Dim gotov As Double
gotov = 0
kart = 0
zne = 0
ddva = 0

ddvb = 0
hrana = 0
pijaca = 0
cig = 0

vsto = 0
Dim davek As Double
Dim vrsta As Integer
Do While Not rs.EOF
If rs.Fields("sifrapart") <> 0 Then
orr = orr + rs.Fields(0)
End If
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

If IsNull(rs.Fields("sifra")) Then
rs.Fields("sifra") = 0
rs.Update
End If
vrsta = Getnumb("select madagrup from mada where madasifr='" & rs.Fields("sifra") & "'")
If Getnumb("select vr from grupa where sifra=" & vrsta) = 0 Then
pijaca = pijaca + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 1 Then
hrana = hrana + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 2 Then
cig = cig + rs.Fields(0)
End If

If Getnumb("select vr from grupa where sifra=" & vrsta) = 4 Then
vsto = vsto + rs.Fields(0)
End If



zne = zne + rs.Fields(0)
If Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'") = "20" Then
ddva = ddva + rs.Fields(0)
End If
If Replace(Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'"), ",", ".") = "8.5" Then
ddvb = ddvb + rs.Fields(0)
End If

rs.MoveNext
Loop
Dim aa As Double
Dim bb As Double
'aa = Format(Me.Text1.text, "0.00")
'bb = Format(Me.Text2.text, "0.00")

If rs.State = 1 Then rs.Close
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic

Dim ee As Long

Dim ff As Long
ee = 0
ff = 0

If Not rs.EOF Then

If IsNull(rs.Fields(0)) Then
ee = 0
Else
ee = rs.Fields(0)
End If
If IsNull(rs.Fields(1)) Then

ff = 0
Else
ff = rs.Fields(1)
End If

If Me.Combo1.text <> "" Then
Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Prijava ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ee, 6) & "'") & vbCrLf
    Text1.text = Text1.text & "Odjava  ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ff, 6) & "'") & vbCrLf
    
End If

    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Zacetna st.rac. : " & ee & vbCrLf
    Text1.text = Text1.text & "Konèna st.rac.  : " & ff & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj izdano raèunov : " & ff - ee + 1 & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj znesek prodaje : " & FormatNumber(zne, 2) & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
   
   Text1.text = Text1.text & Getnacindan(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf
   Text1.text = Text1.text & vbCrLf
 
 Text1.text = Text1.text & "G R U P E==============================" & vbCrLf
  
If rs.State = 1 Then rs.Close
If Me.Combo1.text = "" Then
rs.Open "SELECT mada.madagrup,min(nabasif.DATUM) as datum, nabasif.SIFRA, sum(nabasif.KOL) as kol, sum(nabasif.ZNES) as znes, min(nabasif.naziv) as naziv FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR  where  nabasif.TIP_DOK='PA' and nabasif.DATUM=#" & dod & "# group by mada.MADAGRUP,nabasif.sifra order by mada.MADAGRUP,nabasif.sifra", myConection, adOpenDynamic, adLockOptimistic

Else
rs.Open "SELECT mada.madagrup,min(nabasif.DATUM) as datum, nabasif.SIFRA, sum(nabasif.KOL) as kol, sum(nabasif.ZNES) as znes, min(nabasif.naziv) as naziv FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR  where  nabasif.TIP_DOK='PA'  and nabasif.uporabnik='" & Left(Combo1.text, 2) & "' and nabasif.DATUM=#" & dod & "# group by mada.MADAGRUP,nabasif.sifra order by mada.MADAGRUP,nabasif.sifra", myConection, adOpenDynamic, adLockOptimistic
'rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and uporabnik='" & Left(Combo1.Text, 2) & "' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
End If
If Not rs.EOF Then
rs.MoveFirst
Dim grrr As Long
grrr = 0
Do While Not rs.EOF
If grrr <> rs.Fields("madagrup") Then
        Text1.text = Text1.text & novast(rs.Fields("madagrup"), 2) & "-" & presled(Getnazi("select grupa from grupa where sifra=" & rs.Fields("madagrup")), 16) & "--------KOL------VRE" & vbCrLf
End If
   Text1.text = Text1.text & presled(rs.Fields("naziv"), 21) & levi_pres(FormatNumber(rs.Fields("kol"), 2), 9) & levi_pres(FormatNumber(rs.Fields("znes"), 2), 9) & vbCrLf
        

grrr = rs.Fields("madagrup")
rs.MoveNext
Loop
End If
     If ddva <> 0 Or ddvb <> 0 Then
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
        Text1.text = Text1.text & "Osnova  DDV        Znesek DDV   Vrednost" & vbCrLf
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
       
        If ddva <> 0 Then
    
         Text1.text = Text1.text & presled(Format(ddva / 1.2, "standard"), 8) & "20 %" & levi_pres(Format(ddva - (ddva / 1.2), "standard"), 14) & levi_pres(Format(ddva, "standard"), 14) & vbCrLf
   
        End If
        If ddvb <> 0 Then
         Text1.text = Text1.text & presled(Format(ddvb / 1.085, "standard"), 8) & "8,5 %" & levi_pres(Format(ddvb - (ddvb / 1.085), "standard"), 14) & levi_pres(Format(ddvb, "standard"), 14) & vbCrLf
        End If
       Text1.text = Text1.text & "----------------------------------------" & vbCrLf
    End If
If Me.Combo1.text <> "" Then
Text1.text = Text1.text & Getnacindancig(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf
End If
    Text1.text = Text1.text & vbCrLf & vbCrLf
   
      Text1.text = Text1.text & vbCrLf
    If Getnazi("select konec1 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec1 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec2 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
   Text1.text = Text1.text & Getnazi("select konec3 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec4 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
     Text1.text = Text1.text & Getnazi("select konec5 from oblikar") & vbCrLf
    End If
     Text1.text = Text1.text & vbCrLf & vbCrLf & vbCrLf & vbCrLf

End If

End Sub
Sub zaklju()
Text1.text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & _
vbCrLf & vbCrLf
'vbCrLf & "Time:" & vbCrLf
If Me.Combo1.text = "" Then
Text1.text = Text1.text & "REKAPITULACIJA Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Else
Text1.text = Text1.text & "DELNI ZAKLJUCEK Z DNE: " & Left(Me.List1.text, 10) & vbCrLf
Text1.text = Text1.text & "PREGLED ZAPOSLENEGA: " & Me.Combo1.text & vbCrLf

End If
   If rs.State = 1 Then rs.Close
 Dim das, des
das = Left(Me.List1.text, 10)

dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
If Me.Combo1.text = "" Then
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
Else
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  tip_dok='PA' and uporabnik='" & Left(Combo1.text, 2) & "' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
End If
If Not rs.EOF Then
rs.MoveFirst
End If

Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim orr As Double
Dim hrana As Double
Dim pijaca As Double
Dim cig As Double
Dim vsto As Double
Dim storitve As Double
Dim kart As Double
Dim gotov As Double
gotov = 0
kart = 0
zne = 0
ddva = 0

ddvb = 0
hrana = 0
pijaca = 0
cig = 0
storitve = 0

storitve = Getnumb("SELECT  Sum(nabasif.ZNES) AS vv FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR WHERE tip_dok='PA' and datum=#" & dod & "# and (((mada.tip_art)='STO'))")

vsto = 0
Dim davek As Double
Dim vrsta As Integer
Do While Not rs.EOF
If rs.Fields("sifrapart") <> 0 Then
orr = orr + rs.Fields(0)
End If
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

If IsNull(rs.Fields("sifra")) Then
rs.Fields("sifra") = 0
rs.Update
End If
vrsta = Getnumb("select madagrup from mada where madasifr='" & rs.Fields("sifra") & "'")
If Getnumb("select vr from grupa where sifra=" & vrsta) = 0 Then
pijaca = pijaca + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 1 Then
hrana = hrana + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 2 Then
cig = cig + rs.Fields(0)
End If

If Getnumb("select vr from grupa where sifra=" & vrsta) = 4 Then
vsto = vsto + rs.Fields(0)
End If



zne = zne + rs.Fields(0)
If Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'") = "20" Then
ddva = ddva + rs.Fields(0)
End If
If Replace(Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'"), ",", ".") = "8.5" Then
ddvb = ddvb + rs.Fields(0)
End If

rs.MoveNext
Loop
Dim aa As Double
Dim bb As Double
'aa = Format(Me.Text1.text, "0.00")
'bb = Format(Me.Text2.text, "0.00")

If rs.State = 1 Then rs.Close
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where  tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic

Dim ee As Long

Dim ff As Long
ee = 0
ff = 0

If Not rs.EOF Then

If IsNull(rs.Fields(0)) Then
ee = 0
Else
ee = rs.Fields(0)
End If
If IsNull(rs.Fields(1)) Then

ff = 0
Else
ff = rs.Fields(1)
End If

If Me.Combo1.text <> "" Then
Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Prijava ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ee, 6) & "'") & vbCrLf
    Text1.text = Text1.text & "Odjava  ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ff, 6) & "'") & vbCrLf
    
End If

    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Zacetna st.rac. : " & ee & vbCrLf
    Text1.text = Text1.text & "Konèna st.rac.  : " & ff & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj izdano raèunov : " & ff - ee + 1 & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
    Text1.text = Text1.text & "Skupaj znesek prodaje : " & FormatNumber(zne, 2) & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
   Text1.text = Text1.text & "Skupaj znesek storitev : " & FormatNumber(storitve, 2) & vbCrLf
    Text1.text = Text1.text & "=======================================" & vbCrLf
   
   Text1.text = Text1.text & Getnacindan(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf

     If ddva <> 0 Or ddvb <> 0 Then
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
        Text1.text = Text1.text & "Osnova  DDV        Znesek DDV   Vrednost" & vbCrLf
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
       
        If ddva <> 0 Then
    
         Text1.text = Text1.text & presled(Format(ddva / 1.2, "standard"), 8) & "20 %" & levi_pres(Format(ddva - (ddva / 1.2), "standard"), 14) & levi_pres(Format(ddva, "standard"), 14) & vbCrLf
   
        End If
        If ddvb <> 0 Then
         Text1.text = Text1.text & presled(Format(ddvb / 1.085, "standard"), 8) & "8,5 %" & levi_pres(Format(ddvb - (ddvb / 1.085), "standard"), 14) & levi_pres(Format(ddvb, "standard"), 14) & vbCrLf
        End If
       Text1.text = Text1.text & "----------------------------------------" & vbCrLf
    End If
If Me.Combo1.text <> "" Then
Text1.text = Text1.text & Getnacindancig(Left(List1.text, 10), Trim(Left(Combo1.text, 2))) & vbCrLf
End If
    Text1.text = Text1.text & vbCrLf & vbCrLf
   
      Text1.text = Text1.text & vbCrLf
    If Getnazi("select konec1 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec1 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec2 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
   Text1.text = Text1.text & Getnazi("select konec3 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec4 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
     Text1.text = Text1.text & Getnazi("select konec5 from oblikar") & vbCrLf
    End If
     Text1.text = Text1.text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

End If
End Sub
Public Sub racuni()
Me.Text2(0).text = Getnazi("select dod0 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
Me.Text2(1).text = Getnazi("select dod1 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
Me.Text2(2).text = Getnazi("select dod2 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
Me.Text2(3).text = Getnazi("select dod3 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
Me.Text2(4).text = Getnazi("select org from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
Dim plla As Integer
plla = Getnazi("select placilo from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
If plla <> 0 Then
If plla = 1 Then
Me.Check1.Value = 0
Me.Check2.Value = 0
Me.Check3.Value = 1
Else
Me.Check1.Value = 0
Me.Check2.Value = 1
Me.Check3.Value = 0

End If
Else
Check1.Value = 1
Check2.Value = 0
Check3.Value = 0

End If
Text1.text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & _
vbCrLf & vbCrLf & "Datum: " & Mid(Me.List1.text, 9, 10) & vbCrLf & vbCrLf
'vbCrLf & "Time:" & vbCrLf
If Me.Text2(0).text <> "" Then
Text1.text = Text1.text & "Stranka: " & vbCrLf
Text1.text = Text1.text & Getnazi("select dod0 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & _
vbCrLf & Getnazi("select dod1 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & _
vbCrLf & Getnazi("select dod2 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & _
vbCrLf & Getnazi("select dod3 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & _
vbCrLf & vbCrLf & "ID.ST.: SI" & Getnazi("select org from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & vbCrLf & vbCrLf
End If
Text1.text = Text1.text & "RACUN STEVILKA: " & Left(Me.List1.text, 6) & vbCrLf
Text1.text = Text1.text & "Prodajalec:" & Getnazi("select username1 from users where up='" & Getnazi("select uporabnik from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'") & "'") & vbCrLf & vbCrLf
If Getdoba(LTrim(Left(Me.List1.text, 6))) <> "" Then
    
   Text1.text = Text1.text & "Dobavnice: " & vbCrLf
   Text1.text = Text1.text & Getdoba(LTrim(Left(Me.List1.text, 6))) & vbCrLf
    '& " " & Format(Time(), "hh:mm")
End If
Dim stnarr As String
stnarr = "NK" & Getnazi("select id_dok from nabasif where tip_dok='NK' and kopija='" & Left(Me.List1.text, 6) & "'")
If Len(stnarr) > 2 Then
   
   Text1.text = Text1.text & "Naroèilo : " & stnarr & vbCrLf
   Text1.text = Text1.text & Getnazi("select dod1 from glavna where tip_dok+id_dok='" & stnarr & "'") & vbCrLf
    Text1.text = Text1.text & Getnazi("select dod2 from glavna where tip_dok+id_dok='" & stnarr & "'") & vbCrLf
   Text1.text = Text1.text & Getnazi("select dod3 from glavna where tip_dok+id_dok='" & stnarr & "'") & vbCrLf
   Text1.text = Text1.text & Getnazi("select dod4 from glavna where tip_dok+id_dok='" & stnarr & "'") & vbCrLf
    
      Text1.text = Text1.text & Getnazi("select dod5 from glavna where tip_dok+id_dok='" & stnarr & "'") & vbCrLf
    End If
Text1.text = Text1.text & "========================================" & vbCrLf
Text1.text = Text1.text & "Naziv                   kol  pop  znesek" & vbCrLf
Text1.text = Text1.text & "========================================" & vbCrLf

Dim rst As New ADODB.Recordset
rst.Open "select * from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'", myConection, adOpenDynamic, adLockOptimistic
If rst.EOF Then
Exit Sub
End If
rst.MoveFirst
Dim ZNESE As Double
Dim ddva As Double
Dim ddvb As Double
Dim ddvc As Double
Dim placi  As Integer
placi = 0
ZNESE = 0
ddva = 0
ddvb = 0
ddvc = 0
Do While Not rst.EOF
Text1.text = Text1.text & presled(Left(rst.Fields("naziv"), 20), 20) & levi_pres(FormatNumber(rst.Fields("kol"), 2), 7) & levi_pres(FormatNumber(rst.Fields("pop"), 2), 5) & levi_pres(FormatNumber(rst.Fields("znes"), 2), 8) & vbCrLf
ZNESE = ZNESE + rst.Fields("ZNES")
If Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) = 20 Then
ddva = ddva + rst.Fields("ZNES")
End If
If Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) > 7 And Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) < 9 Then
ddvb = ddvb + rst.Fields("ZNES")
End If
If Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) = 0 Then
ddvc = ddvc + rst.Fields("ZNES")
End If
placi = rst.Fields("placilo")


rst.MoveNext
Loop

Text1.text = Text1.text & "========================================" & vbCrLf


Text1.text = Text1.text & "ZA PLACILO EUR " & levi_pres(FormatNumber(ZNESE, 2), 25) & vbCrLf
     If ddva <> 0 Or ddvb <> 0 Then
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
        Text1.text = Text1.text & "Osnova  DDV        Znesek DDV   Vrednost" & vbCrLf
        Text1.text = Text1.text & "----------------------------------------" & vbCrLf
       
        If ddva <> 0 Then
    
         Text1.text = Text1.text & presled(Format(ddva / 1.2, "standard"), 8) & "20 %" & levi_pres(Format(ddva - (ddva / 1.2), "standard"), 14) & levi_pres(Format(ddva, "standard"), 14) & vbCrLf
   
        End If
        If ddvb <> 0 Then
         Text1.text = Text1.text & presled(Format(ddvb / 1.085, "standard"), 8) & "8,5 %" & levi_pres(Format(ddvb - (ddvb / 1.085), "standard"), 14) & levi_pres(Format(ddvb, "standard"), 14) & vbCrLf
        End If
        If ddvc <> 0 Then
         Text1.text = Text1.text & presled(Format(ddvc, "standard"), 8) & " 0 %" & levi_pres(Format(0, "standard"), 14) & levi_pres(Format(ddvc, "standard"), 14) & vbCrLf
        End If
       Text1.text = Text1.text & "----------------------------------------" & vbCrLf
    End If
    Dim pl As String
    pl = "Gotovino"
    If placi = 9999 Then
    pl = "Kartico"
    End If
    If placi = 1 Then
    pl = "Reprezentanco"
    End If
    Text1.text = Text1.text & "Placano z " & vbCrLf
    Text1.text = Text1.text & Getnacin("PA" & Left(Me.List1.text, 6)) & vbCrLf & vbCrLf
    If placi = 1 Then
    Text1.text = Text1.text & vbCrLf & "Podpis:_______________ " & pl & vbCrLf
    End If
      'cPrint.pPrint " Placilo: " & plax, 0.1, False
      
      Text1.text = Text1.text & vbCrLf
    If Getnazi("select konec1 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec1 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec2 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
   Text1.text = Text1.text & Getnazi("select konec3 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    Text1.text = Text1.text & Getnazi("select konec4 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
     Text1.text = Text1.text & Getnazi("select konec5 from oblikar") & vbCrLf
    End If
     Text1.text = Text1.text & vbCrLf & vbCrLf & vbCrLf & vbCrLf
If placi = 1 Then
Me.Check3.Value = 1
Me.Check2.Value = 0
Me.Check1.Value = 0
End If
If placi = 9999 Then
Me.Check3.Value = 0
Me.Check2.Value = 1
Me.Check1.Value = 0
End If


End Sub

Private Sub OKButton_Click()
Dim ahh As Long
ahh = 0
If Check2.Value = 1 Then
ahh = 9999
End If
If Check3.Value = 1 Then
ahh = 1
End If
myConection.Execute ("update nabasif set placilo=" & ahh & " where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
If Me.Text2(4).text <> "" Then
myConection.Execute ("update nabasif set org=" & Me.Text2(4).text & " where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
End If
If Me.Text2(0).text <> "" Then
myConection.Execute ("update glavna set dod0='" & Me.Text2(0).text & "' where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
myConection.Execute ("update glavna set dod1='" & Me.Text2(1).text & "' where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
myConection.Execute ("update glavna set dod2='" & Me.Text2(2).text & "' where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
myConection.Execute ("update glavna set dod3='" & Me.Text2(3).text & "' where tip_dok='PA' and id_dok='" & Left(Me.List1.text, 6) & "'")
End If
racuni
End Sub

Private Sub Text1_GotFocus()
Me.Combo1.SetFocus
End Sub

Private Sub Text2_Click(Index As Integer)
idtipk = 999
For idtipk = 0 To 4
If Index <> idtipk Then
Me.Text2(idtipk).BackColor = &HFFFFFF
End If
Next
If Me.Text2(Index).BackColor = &HC0C0FF Then
Me.Text2(Index).BackColor = &HFFFFFF
Else
Me.Text2(Index).BackColor = &HC0C0FF
idtipk = Index
End If
End Sub
