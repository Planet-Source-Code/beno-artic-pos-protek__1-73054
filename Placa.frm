VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{46E98F52-504C-4B1B-B951-CE2725A20438}#1.1#0"; "gdpicturepro4.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Placa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   10785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form7"
   Picture         =   "Placa.frx":0000
   ScaleHeight     =   10785
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton deloo 
      Height          =   255
      Left            =   11760
      TabIndex        =   219
      Top             =   10440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "delo"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E0A
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
   Begin LVbuttons.LaVolpeButton zapoo 
      Height          =   255
      Left            =   12600
      TabIndex        =   218
      Top             =   10440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Zap."
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E26
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
   Begin LVbuttons.LaVolpeButton des 
      Height          =   375
      Left            =   13560
      TabIndex        =   217
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "==>"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E42
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
   Begin LVbuttons.LaVolpeButton vsi 
      Height          =   375
      Left            =   12360
      TabIndex        =   216
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "VSI"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E5E
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
   Begin LVbuttons.LaVolpeButton lev 
      Height          =   375
      Left            =   11400
      TabIndex        =   215
      Top             =   8400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<=="
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E7A
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
   Begin VB.TextBox Stan 
      Height          =   285
      Left            =   11280
      TabIndex        =   212
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox opi 
      Height          =   285
      Left            =   13440
      TabIndex        =   211
      Text            =   "Text3"
      Top             =   10560
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   13800
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Co 
      Caption         =   "Command1"
      Height          =   255
      Left            =   13680
      TabIndex        =   210
      Top             =   600
      Width           =   615
   End
   Begin GdPicturePro4.GdViewer GdViewer1 
      Height          =   255
      Left            =   12840
      TabIndex        =   209
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   495
      Left            =   11400
      TabIndex        =   208
      Top             =   8880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "IZPIS"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0E96
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   13080
      TabIndex        =   207
      Top             =   9720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "KONCAJ"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0EB2
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
      Left            =   11400
      TabIndex        =   206
      Top             =   9720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "ZAPIŠI"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Placa.frx":0ECE
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   11280
      TabIndex        =   205
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   63
      Left            =   9480
      TabIndex        =   36
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   62
      Left            =   9480
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   61
      Left            =   9480
      TabIndex        =   38
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   60
      Left            =   9480
      TabIndex        =   39
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   59
      Left            =   9480
      TabIndex        =   40
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   58
      Left            =   9480
      TabIndex        =   41
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   57
      Left            =   9480
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   56
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   55
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.MM.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1060
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   199
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46202881
      CurrentDate     =   39670
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   9360
      TabIndex        =   197
      Top             =   10440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   3840
      TabIndex        =   196
      Top             =   10440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   51
      Left            =   9360
      TabIndex        =   174
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   54
      Left            =   8040
      TabIndex        =   110
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   53
      Left            =   8040
      TabIndex        =   109
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   52
      Left            =   8040
      TabIndex        =   108
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   51
      Left            =   8040
      TabIndex        =   107
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   50
      Left            =   8040
      TabIndex        =   106
      Top             =   9600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   49
      Left            =   8040
      TabIndex        =   105
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   48
      Left            =   8040
      TabIndex        =   104
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   47
      Left            =   8040
      TabIndex        =   103
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   46
      Left            =   8040
      TabIndex        =   102
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   45
      Left            =   8040
      TabIndex        =   101
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   50
      Left            =   9360
      TabIndex        =   100
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   49
      Left            =   9360
      TabIndex        =   99
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   48
      Left            =   9360
      TabIndex        =   98
      Top             =   9840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   47
      Left            =   9360
      TabIndex        =   97
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   46
      Left            =   9360
      TabIndex        =   96
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   45
      Left            =   9360
      TabIndex        =   95
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   44
      Left            =   9360
      TabIndex        =   94
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   43
      Left            =   9360
      TabIndex        =   93
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   44
      Left            =   9480
      TabIndex        =   90
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   43
      Left            =   9480
      TabIndex        =   89
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   42
      Left            =   9360
      TabIndex        =   88
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   41
      Left            =   9360
      TabIndex        =   87
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   40
      Left            =   9360
      TabIndex        =   86
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   39
      Left            =   9360
      TabIndex        =   85
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   38
      Left            =   9360
      TabIndex        =   84
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   31
      Left            =   3840
      TabIndex        =   83
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   30
      Left            =   3840
      TabIndex        =   82
      Top             =   9840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   29
      Left            =   3840
      TabIndex        =   81
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   28
      Left            =   3840
      TabIndex        =   80
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   27
      Left            =   3840
      TabIndex        =   79
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   26
      Left            =   3840
      TabIndex        =   78
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   25
      Left            =   3840
      TabIndex        =   77
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   24
      Left            =   3840
      TabIndex        =   76
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   23
      Left            =   3840
      TabIndex        =   75
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   22
      Left            =   3840
      TabIndex        =   74
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   21
      Left            =   3840
      TabIndex        =   73
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   20
      Left            =   3840
      TabIndex        =   72
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   19
      Left            =   3840
      TabIndex        =   71
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   18
      Left            =   3840
      TabIndex        =   70
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   17
      Left            =   3840
      TabIndex        =   69
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   16
      Left            =   3840
      TabIndex        =   68
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   15
      Left            =   3840
      TabIndex        =   67
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   14
      Left            =   3840
      TabIndex        =   66
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   13
      Left            =   3840
      TabIndex        =   65
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   12
      Left            =   3840
      TabIndex        =   64
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   11
      Left            =   3840
      TabIndex        =   63
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   10
      Left            =   3840
      TabIndex        =   62
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   9
      Left            =   3840
      TabIndex        =   61
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   8
      Left            =   3840
      TabIndex        =   60
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   7
      Left            =   3840
      TabIndex        =   59
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   6
      Left            =   3840
      TabIndex        =   58
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   3840
      TabIndex        =   57
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   3840
      TabIndex        =   56
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   55
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   54
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   53
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   52
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   42
      Left            =   9480
      TabIndex        =   51
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   41
      Left            =   9480
      TabIndex        =   50
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   40
      Left            =   9480
      TabIndex        =   49
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   39
      Left            =   9480
      TabIndex        =   48
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   38
      Left            =   9480
      TabIndex        =   47
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   37
      Left            =   9480
      TabIndex        =   46
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   36
      Left            =   9480
      TabIndex        =   45
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   35
      Left            =   9480
      TabIndex        =   44
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   34
      Left            =   9480
      TabIndex        =   43
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   33
      Left            =   9480
      TabIndex        =   42
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   32
      Left            =   2520
      TabIndex        =   35
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   31
      Left            =   2520
      TabIndex        =   34
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   30
      Left            =   2520
      TabIndex        =   33
      Top             =   9600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   29
      Left            =   2520
      TabIndex        =   32
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   28
      Left            =   2520
      TabIndex        =   31
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   27
      Left            =   2520
      TabIndex        =   30
      Top             =   8880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   26
      Left            =   2520
      TabIndex        =   29
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   25
      Left            =   2520
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   24
      Left            =   2520
      TabIndex        =   27
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   23
      Left            =   2520
      TabIndex        =   26
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   22
      Left            =   2520
      TabIndex        =   25
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   21
      Left            =   2520
      TabIndex        =   24
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   20
      Left            =   2520
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   19
      Left            =   2520
      TabIndex        =   22
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   18
      Left            =   2520
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   17
      Left            =   2520
      TabIndex        =   20
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   16
      Left            =   2520
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   15
      Left            =   2520
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   14
      Left            =   2520
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   13
      Left            =   2520
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   12
      Left            =   2520
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   11
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   10
      Left            =   2520
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   9
      Left            =   2520
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   8
      Left            =   2520
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   7
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   6
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11160
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label dd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Saldo:"
      Height          =   255
      Left            =   11400
      TabIndex        =   214
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Saldo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0,00"
      Height          =   255
      Left            =   11280
      TabIndex        =   213
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vrednost ure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   204
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vrednost toèke:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   203
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Število toèk:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   202
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fond ur.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   201
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   255
      Left            =   13440
      TabIndex        =   200
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label dok 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   11880
      TabIndex        =   198
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "IZPLAÈILO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   195
      Top             =   10440
      Width           =   3735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NETO PLAÈA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   194
      Top             =   10440
      Width           =   3735
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Obraèun davka na plaèilno listo"
      Height          =   255
      Left            =   5640
      TabIndex        =   193
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stopnja [% ]    "
      Height          =   255
      Left            =   8040
      TabIndex        =   192
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   9360
      TabIndex        =   191
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrsta prispevka na BOD"
      Height          =   255
      Left            =   5640
      TabIndex        =   190
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stopnja [% ]    "
      Height          =   255
      Left            =   8040
      TabIndex        =   189
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   9360
      TabIndex        =   188
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrsta prispevka iz BOD"
      Height          =   255
      Left            =   120
      TabIndex        =   187
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stopnja [% ]    "
      Height          =   255
      Left            =   2520
      TabIndex        =   186
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   3840
      TabIndex        =   185
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrsta olajšave in dohodek"
      Height          =   255
      Left            =   120
      TabIndex        =   184
      Top             =   8280
      Width           =   3855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   3840
      TabIndex        =   183
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Odbitki"
      Height          =   255
      Left            =   5640
      TabIndex        =   182
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   9360
      TabIndex        =   181
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Dodatki"
      Height          =   255
      Left            =   5640
      TabIndex        =   180
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   9360
      TabIndex        =   179
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrednost  "
      Height          =   255
      Left            =   3840
      TabIndex        =   178
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "URE [% ]    "
      Height          =   255
      Left            =   2520
      TabIndex        =   177
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vrsta obraèuna"
      Height          =   255
      Left            =   120
      TabIndex        =   176
      Top             =   720
      Width           =   2415
   End
   Begin GdPicturePro4.Imaging Imaging1 
      Left            =   11160
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   65
      Left            =   5640
      TabIndex        =   175
      Top             =   10080
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   64
      Left            =   5640
      TabIndex        =   173
      Top             =   9840
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   63
      Left            =   5640
      TabIndex        =   172
      Top             =   9600
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   62
      Left            =   5640
      TabIndex        =   171
      Top             =   9360
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   61
      Left            =   5640
      TabIndex        =   170
      Top             =   9120
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   60
      Left            =   5640
      TabIndex        =   169
      Top             =   8880
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   59
      Left            =   5640
      TabIndex        =   168
      Top             =   8160
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   58
      Left            =   5640
      TabIndex        =   167
      Top             =   7920
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   57
      Left            =   5640
      TabIndex        =   166
      Top             =   7680
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   56
      Left            =   5640
      TabIndex        =   165
      Top             =   7440
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   55
      Left            =   5640
      TabIndex        =   164
      Top             =   7200
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   54
      Left            =   5640
      TabIndex        =   163
      Top             =   6960
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   53
      Left            =   5640
      TabIndex        =   162
      Top             =   6240
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   52
      Left            =   5640
      TabIndex        =   161
      Top             =   5880
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   51
      Left            =   5640
      TabIndex        =   160
      Top             =   5640
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   50
      Left            =   5640
      TabIndex        =   159
      Top             =   5400
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   49
      Left            =   5640
      TabIndex        =   158
      Top             =   5160
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   48
      Left            =   5640
      TabIndex        =   157
      Top             =   4920
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   47
      Left            =   5640
      TabIndex        =   156
      Top             =   4680
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   46
      Left            =   5640
      TabIndex        =   155
      Top             =   4440
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   45
      Left            =   5640
      TabIndex        =   154
      Top             =   4200
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   44
      Left            =   5640
      TabIndex        =   153
      Top             =   3960
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   43
      Left            =   5640
      TabIndex        =   152
      Top             =   3720
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   42
      Left            =   5640
      TabIndex        =   151
      Top             =   3480
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   41
      Left            =   5640
      TabIndex        =   150
      Top             =   3240
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   40
      Left            =   5640
      TabIndex        =   149
      Top             =   2520
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   39
      Left            =   5640
      TabIndex        =   148
      Top             =   2280
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   38
      Left            =   5640
      TabIndex        =   147
      Top             =   2040
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   37
      Left            =   5640
      TabIndex        =   146
      Top             =   1800
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   36
      Left            =   5640
      TabIndex        =   145
      Top             =   1560
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   35
      Left            =   5640
      TabIndex        =   144
      Top             =   1320
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   34
      Left            =   5640
      TabIndex        =   143
      Top             =   1080
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   33
      Left            =   120
      TabIndex        =   142
      Top             =   10080
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   32
      Left            =   120
      TabIndex        =   141
      Top             =   9840
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   31
      Left            =   120
      TabIndex        =   140
      Top             =   9600
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   30
      Left            =   120
      TabIndex        =   139
      Top             =   9360
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   29
      Left            =   120
      TabIndex        =   138
      Top             =   9120
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   28
      Left            =   120
      TabIndex        =   137
      Top             =   8880
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   27
      Left            =   120
      TabIndex        =   136
      Top             =   8640
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   26
      Left            =   120
      TabIndex        =   135
      Top             =   7680
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   25
      Left            =   120
      TabIndex        =   134
      Top             =   7440
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   24
      Left            =   120
      TabIndex        =   133
      Top             =   7200
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   23
      Left            =   120
      TabIndex        =   132
      Top             =   6960
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   22
      Left            =   120
      TabIndex        =   131
      Top             =   6720
      Width           =   2300
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   120
      TabIndex        =   130
      Top             =   6120
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   20
      Left            =   120
      TabIndex        =   129
      Top             =   5880
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   128
      Top             =   5640
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   127
      Top             =   5400
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   126
      Top             =   5160
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   125
      Top             =   4920
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   124
      Top             =   4680
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   123
      Top             =   4440
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   122
      Top             =   4200
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   121
      Top             =   3960
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   120
      Top             =   3720
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   119
      Top             =   3480
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   118
      Top             =   3240
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   117
      Top             =   3000
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   116
      Top             =   2760
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   115
      Top             =   2520
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   114
      Top             =   2280
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   113
      Top             =   2040
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   112
      Top             =   1800
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   111
      Top             =   1560
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   92
      Top             =   1320
      Width           =   2300
   End
   Begin VB.Label Label3 
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   91
      Top             =   1080
      Width           =   2300
   End
End
Attribute VB_Name = "Placa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Private Declare Function EbExecuteLine Lib "vba6.dll" _
        (ByVal pStringToExec As Long, ByVal Foo1 As Long, _
        ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long

Function FExecuteCode(stCode As String, Optional fCheckOnly _
    As Boolean) As Boolean
    FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, _
        Abs(fCheckOnly)) = 0
End Function

Private Sub Command1_Click()
On Error GoTo bb:
Dim res As Boolean
  '  res = FExecuteCode(Text1(0).text)
  '  Me.Label1.Caption = res
 ' Me.PrintForm
Dim nnn As Boolean
Dim ln_id As Variant
Me.Imaging1.SetLicenseNumber ("1519740135762015145551548")

 nnn = Me.Imaging1.CreateImageFromHwnd(Me.hWnd)
 'MsgBox nnn
ln_id = Me.Imaging1.GetNativeImage()
nnn = Me.Imaging1.SaveAsTiff("c:\bb.tiff", CompressionCCITT4)
nnn = Me.Imaging1.CreateNewImage(1100, 1500, 32, White)
nnn = Me.Imaging1.DrawText(Getnazi("select glava1 from oblikar"), 50, 50, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava2 from oblikar"), 50, 70, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava3 from oblikar"), 50, 90, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava4 from oblikar"), 50, 110, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava4 from oblikar"), 50, 130, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawImage(ln_id, 50, 150, 1450, 1350, InterpolationModeHighQuality)
nnn = Me.Imaging1.SaveAsTiff("c:\gg.tiff")
bb:
  Unload Me
End Sub
Private Function t(ii As Integer) As Double
t = 0
If Me.Text1(ii).text <> "" Then
t = Me.Text1(ii).text
End If
End Function

Private Sub Co_Click()
'
End Sub

Private Sub deloo_Click()
Dim contrd As Control
Dim drrs As New ADODB.Recordset
Dim krrs As New ADODB.Recordset
Dim sifr, sqll, sqx As String
sifr = RTrim(LTrim(Left(Me.List1.text, 6)))
drrs.Open "delete from delo where sifd='" & sifr & "'", myConection, adOpenDynamic, adLockOptimistic
drrs.Open "select * from delo where sifd='" & sifr & "'", myConection, adOpenDynamic, adLockOptimistic
If drrs.EOF Then
drrs.AddNew
drrs.Fields("sifd") = sifr
drrs.Update
End If
'MsgBox "update dmat set tock=" & FormatNumber(Me.Text1(57).text, 2) & " where sifd='" & sifr & "'"
myConection.Execute "update dmat set tock=" & Replace(FormatNumber(Me.Text1(57).text, 2), ",", ".") & " where sifd='" & sifr & "'"
myConection.Execute "update dmat set spol=" & Replace(FormatNumber(Me.Text2(27).text, 2), ",", ".") & " where sifd='" & sifr & "'"
myConection.Execute "update dmat set pool=" & Replace(FormatNumber(Me.Text2(29).text, 2), ",", ".") & " where sifd='" & sifr & "'"
For Each contrd In Me.Controls
If contrd.Name = "text1" Then
'If RTrim(LTrim((Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'")))) = "" Then
'MsgBox Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'")
sqx = ""
sqx = RTrim(LTrim(Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'")))
'if isnumeric(contrd.text)
'MsgBox "select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'"
If Not sqx = "" Then

sqll = "update delo set " & sqx & "=" & Replace(FormatNumber(contrd.text, 2), ",", ".") & " where sifd='" & sifr & "'"
myConection.Execute sqll
'End If
End If
End If
If contrd.Name = "text2" Then
'If RTrim(LTrim((Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'")))) = "" Then
sqx = ""

sqx = RTrim(LTrim(Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2X" & levi_pres(LTrim(str(contrd.Index)), 2) & "'")))
'if isnumeric(contrd.text)
If Not sqx = "" Then
sqll = "update delo set " & sqx & "=" & FormatNumber(contrd.text, 2) & " where sifd='" & sifr & "'"
myConection.Execute sqll
'End If
End If
End If

Next
myConection.Execute "delete from dkre where sifd='" & sifr & "'"
Dim ozna As String

For Each contrd In Me.Controls
If contrd.Name = "text1" Then
If contrd.Index >= 33 And contrd.Index <= 44 Then
If RTrim(LTrim(contrd.text)) <> "0,00" Then
ozna = ""
ozna = Left(LTrim(RTrim(Me.Label3(contrd.Index + 8).Caption)), 15)
'krrs.Open "select * from dkre where sifd='" & sifr & "'", myConection, adOpenDynamic, adLockOptimistic
myConection.Execute "insert into dkre (sifd,ozn,obrok,obrok1) values ('" & sifr & "','" & ozna & "'," & Replace(FormatNumber(contrd.text, 2), ",", ".") & "," & Replace(FormatNumber(contrd.text, 2), ",", ".") & ")"
End If
End If
End If
Next
End Sub

Private Sub dok_Click()
If Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1  1'") = "" Then
If MsgBox("Nimaš formul za izraèun plaèe ali jih prenesem iz prejšnjega meseca???", vbOKCancel) = vbOK Then
myConection.Execute ("insert into dokm select tip_dok,tekst,atribut from dokm where tip_dok='PL' and id_dok='82008'")
MsgBox "Konèano"

myConection.Execute ("update dokm set id_dok='" & Mid(dok.Caption, 3, 6) & "' where tip_dok='PL' and isnull(id_dok)")

End If
End If
End Sub

Private Sub DTPicker1_Change()
Me.dok.Caption = "PL" & Month(Format(Me.DTPicker1.Value, "dd.mm.yyyy")) & Year(Me.DTPicker1.Value)
End Sub

Private Sub Form_Load()
'Me.Imag
ReSizeForm Me
Dim contr As Control

For Each contr In Me.Controls
If Left(contr.Name, 4) = "Text" Then
contr.text = "0,00"
End If
Next
If RS.State = 1 Then RS.Close
RS.Open "select * from dokm where id_dok='PLACE' and tip_dok='XX' order by poz", myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
RS.MoveFirst
Dim ppp As Integer
Do While Not RS.EOF
ppp = RS.Fields("poz") - 1

Me.Label3(ppp).Caption = IIf(IsNull(RS.Fields("tekst")), "", RS.Fields("tekst"))
RS.MoveNext
Loop
End If
Me.dok.Caption = "PL" & Month(Format(Me.DTPicker1.Value, "dd.mm.yyyy")) & Year(Me.DTPicker1.Value)
'Me.Timer1.Enabled = True
Fiil List1, "select * from zaposleni order by priimek"
osver
End Sub
Private Sub osver()
Dim contr As Control

For Each contr In Me.Controls
If UCase(Left(contr.Name, 5)) = "TEXT1" Then
If Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='t1'") <> "" Then
contr.text = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='t1'")
Else
contr.text = "0,00"
End If
End If
If UCase(Left(contr.Name, 5)) = "TEXT2" Then
If Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='t2'") <> "" Then
contr.text = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='t2'")
Else
contr.text = "0,00"
End If
End If
If UCase(Left(contr.Name, 6)) = "LABEL3" Then
If Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='l3'") <> "" Then
contr.Caption = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "' and poz=" & contr.Index & " and atribut='l3'")
End If
End If
Next

End Sub
Private Sub VScroll1_Change()
'MsgBox VScroll1.Value
End Sub

Private Sub Label3_Click(Index As Integer)
Dim a, b, c, d As Integer
a = Me.Label3(Index).Left
b = Me.Label3(Index).Top
c = Me.Label3(Index).Height
d = Me.Label3(Index).Width
RaiseEvent opi.Move(a, b, d, c)
Me.opi.Visible = True
Me.opi.text = Me.Label3(Index).Caption
Me.opi.SetFocus

kater = Index
End Sub

Private Sub LaVolpeButton1_click()
Dim contr As Control
deloo_Click
'myConection.Execute ("update zaposleni set porabadop=porabadop-" & Me.Text1(3).text & " where sifra='" & Left(Me.List1.text, 6) & "'")

myConection.Execute ("delete from dokm where tip_dok='PL' and id_dok='" & Me.dok.Caption & "'")
For Each contr In Me.Controls
If UCase(Left(contr.Name, 5)) = "TEXT1" Then
myConection.Execute ("insert into dokm (atribut,tip_dok,id_dok,poz,tekst) values ('t1','PL','" & dok.Caption & "'," & contr.Index & ",'" & contr.text & "')")
End If
If UCase(Left(contr.Name, 5)) = "TEXT2" Then
myConection.Execute ("insert into dokm (atribut,tip_dok,id_dok,poz,tekst) values ('t2','PL','" & dok.Caption & "'," & contr.Index & ",'" & contr.text & "')")
End If
If UCase(Left(contr.Name, 6)) = "LABEL3" Then
myConection.Execute ("insert into dokm (atribut,tip_dok,id_dok,poz,tekst) values ('l3','PL','" & dok.Caption & "'," & contr.Index & ",'" & contr.Caption & "')")
End If
Next
'Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub

Private Sub LaVolpeButton3_Click()
On Error GoTo bb:
Dim res As Boolean
Dim sifrad As String
sifrad = RTrim(LTrim(Left(Me.List1.text, 6)))
  '  res = FExecuteCode(Text1(0).text)
  '  Me.Label1.Caption = res
 ' Me.PrintForm
Dim nnn, xxx As Boolean
Dim ln_id, xln_id As Variant
Me.Imaging1.SetLicenseNumber ("1519740135762015145551548")

 nnn = Me.Imaging1.CreateImageFromHwnd(Me.hWnd)
 'MsgBox nnn
 'me.Imaging1.SaveAsJpeg(aa,12
ln_id = Me.Imaging1.GetNativeImage()
nnn = Me.Imaging1.SaveAsTiff("c:\bb.tiff", CompressionCCITT4)
nnn = Me.Imaging1.CreateNewImage(1100, 1500, 32, White)
nnn = Me.Imaging1.DrawText(Getnazi("select glava1 from oblikar"), 50, 50, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava2 from oblikar"), 50, 70, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava3 from oblikar"), 50, 90, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava4 from oblikar"), 50, 110, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText(Getnazi("select glava4 from oblikar"), 50, 130, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawText("Zaposlen: " & RTrim(Getnazi("select ime from zaposleni where sifra='" & sifrad & "'")) & " " & RTrim(Getnazi("select priimek from zaposleni where sifra='" & sifrad & "'")) & " Davèna:" & RTrim(Getnazi("select davcna from zaposleni where sifra='" & sifrad & "'")) & " Dopust ur: " & Val(Getnazi("select dopustur-porabadop from zaposleni where sifra='" & sifrad & "'")) - Val(Getnazi("select sum(val(tekst)) from dokm where atribut='t1' and poz=3 and id_dok='" & Me.dok.Caption & "'")), 50, 130, 16, FontStyleBold, Black, "Arial", False)
nnn = Me.Imaging1.DrawImage(ln_id, 50, 150, 1450, 1350, InterpolationModeHighQuality)
nnn = Me.Imaging1.SaveAsTiff("c:\gg.tiff")

'Me.GdViewer1.SetNativeImage (Me.Imaging1.GetNativeImage())
'Me.GdViewer1.DisplayFromFile ("c:\gg.tiff")
'GdViewer1.Visible = True
'GdViewer1.PrintImage
Imaging1.PrintImageFit
Dim fso As New FileSystemObject
'fso.DeleteFile ("c:\gg.tiff")
bb:
End Sub

Private Sub List1_Click()

'If MsgBox("Ali želiš shraniti podatke???", vbOKCancel) = vbOK Then
'Call LaVolpeButton1_click

'End If
Me.dok.Caption = presled("PL" & Month(Format(Me.DTPicker1.Value, "dd.mm.yyyy")) & Year(Me.DTPicker1.Value), 10) & LTrim(Left(Me.List1.text, 6))
osver
GdViewer1.Visible = False
End Sub

Private Sub opi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.Label3(kater).Caption = Me.opi.text
Me.opi.Visible = False
End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
If intCtrlDown = 2 Then
xopis = "1X" & levi_pres(LTrim(str(Index)), 2)

intCtrlDown = 0

Else
xopis = "1" & levi_pres(LTrim(str(Index)), 3)
End If
    xid_dok = Left(Trim(dok.Caption), 8)
    Dialog.Show vbModal
'End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)

ozna
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
intCtrlDown = Shift

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
If KeyAscii = 27 Then
If Index > 0 Then
    Text1(Index - 1).SetFocus
    End If
End If

End Sub
Private Sub ozna()
SendKeys "{HOME} +{END}"
End Sub

Private Sub Text1_LostFocus(Index As Integer)
'If Index > 4 Then
Me.Timer1.Enabled = True
'End If

If IsNumber(Text1(Index).text) Then
Text1(Index).text = Format(Trim(Text1(Index).text), "fixed")
End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
If intCtrlDown = 2 Then
xopis = "2X" & levi_pres(LTrim(str(Index)), 2)

intCtrlDown = 0

Else
 xopis = "2" & levi_pres(LTrim(str(Index)), 3)
 End If
    xid_dok = Left(Trim(dok.Caption), 8)
    Dialog.Show vbModal
End Sub


Private Function xpre(sss As String) As String
Dim fds As Boolean
fds = FExecuteCode("xpre=" & sss)
End Function

Private Sub Text2_GotFocus(Index As Integer)

ozna

End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
intCtrlDown = Shift
End Sub

Private Sub Text2_LostFocus(Index As Integer)
If IsNumber(Text2(Index).text) Then
Text2(Index).text = Format(Trim(Text2(Index).text), "fixed")
End If
End Sub
Private Sub Timer1_Timer()

Dim contr, cco As Control
Dim res As Boolean
Dim dohh, ddohh As Long
Dim delo As String

Dim xx As Integer
Dim axx As Double
Dim ahaa, axh As String
For Each contr In Me.Controls
If contr.Name = "text1" Then
Placa.Text1(contr.Index).ToolTipText = "tx(" & contr.Index & ")"
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
Placa.Text1(contr.Index).BackColor = &HC0FFFF
Else
Placa.Text1(contr.Index).BackColor = &HC0E0FF
End If
End If
If contr.Name = "text2" Then
Placa.Text2(contr.Index).ToolTipText = "tx2(" & contr.Index & ")"
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
Placa.Text2(contr.Index).BackColor = &HC0FFFF
Else
Placa.Text2(contr.Index).BackColor = &HC0E0FF
End If
End If
Next



For Each contr In Me.Controls
If Left(UCase(contr.Name), 4) = "text" Then

If contr.Name = "text1" And contr.Index = 20 Then


ahaa = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")

For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text2" Then
axh = "tx2(" & LTrim(str(cco.Index)) & ")"
If IsNumber(Placa.Text2(cco.Index).text) Then
axx = FormatNumber(Placa.Text2(cco.Index).text, 2)
End If
ahaa = Replace(ahaa, axh, axx)
End If
If Val(Me.Text2(30).text) < 0 Then
Me.Text2(30).text = "0,00"
Me.Text2(31).text = "0,00"
End If

Next

For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text1" Then
axh = "tx(" & LTrim(str(cco.Index)) & ")"
'MsgBox Placa.Text1(cco.Index).text
If IsNumber(Placa.Text1(cco.Index).text) Then
axx = FormatNumber(Placa.Text1(cco.Index).text, 2)
End If
ahaa = Replace(ahaa, axh, axx)

End If
Next
'
If Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'") <> "" Then
'MsgBox DateDiff("y", Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"), Date)
delo = str(Round(DateDiff("y", Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"), Date) / 365, 0))
ahaa = Replace(ahaa, "DEL", delo)
End If
ahaa = Replace(ahaa, "REZI", Getnazi("select rezi from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
'MsgBox Val(Me.Text2(30).text)
If Val(Me.Text2(30).text) > 0 Then
If Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") <> 0 Then
dohh = Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1")
ddohh = dohh + (Me.Text2(30).text - Getnazi("select zn2 from ddoh order by zn2")) * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
Else
ddohh = Me.Text2(30).text * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
End If

ahaa = Replace(ahaa, "DOH", ddohh)
End If
ahaa = Replace(ahaa, "POT", Getnazi("select km from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, "OTROK", Getnazi("select otrok from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, ",", ".")
contr.text = Format(ScriptControl1.Eval(ahaa), "fixed")
'MsgBox ahaa
End If
End If
Next

For Each contr In Me.Controls
If Left(UCase(contr.Name), 4) = "text" Then

If contr.Name = "text2" And contr.Index = 19 Then

If contr.Name = "text2" Then
ahaa = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'")
Else
ahaa = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")
End If


For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text2" Then
axh = "tx2(" & LTrim(str(cco.Index)) & ")"
If IsNumber(Placa.Text2(cco.Index).text) Then
axx = FormatNumber(Placa.Text2(cco.Index).text, 2)
End If
ahaa = Replace(ahaa, axh, axx)
End If
If Val(Me.Text2(30).text) < 0 Then
Me.Text2(30).text = "0,00"
Me.Text2(31).text = "0,00"
End If

Next

For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text1" Then
axh = "tx(" & LTrim(str(cco.Index)) & ")"
If IsNumber(Placa.Text1(cco.Index).text) Then
axx = FormatNumber(Placa.Text1(cco.Index).text, 2)
End If
ahaa = Replace(ahaa, axh, axx)

End If
Next
If Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'") <> "" Then

delo = str(Round(DateDiff("y", Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"), Date) / 365, 0))
ahaa = Replace(ahaa, "DEL", delo)
End If
If Val(Me.Text2(30).text) > 0 Then
If Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") <> 0 Then
dohh = Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1")
ddohh = dohh + (Me.Text2(30).text - Getnazi("select zn2 from ddoh order by zn2")) * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
Else
ddohh = Me.Text2(30).text * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
End If

ahaa = Replace(ahaa, "DOH", ddohh)
End If
ahaa = Replace(ahaa, "REZI", Getnazi("select rezi from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, "POT", Getnazi("select km from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, "OTROK", Getnazi("select otrok from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, ",", ".")
contr.text = Format(ScriptControl1.Eval(ahaa), "fixed")
End If
End If
Next




For Each contr In Me.Controls
If Left(UCase(contr.Name), 4) = "text" Then
If contr.Name = "text2" And contr.Index = 19 Or contr.Name = "text1" And contr.Index = 20 Then
Else
If contr.BackColor <> &HC0FFFF Then
ahaa = ""
If contr.Name = "text2" Then
ahaa = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'")
Else
ahaa = Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")
End If
'If ahaa <> "" Then


For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text2" Then
axh = "tx2(" & LTrim(str(cco.Index)) & ")"
If cco.Index <> 31 Then
axx = FormatNumber((Placa.Text2(cco.Index).text), 2)
End If
ahaa = Replace(ahaa, axh, axx)
End If
Next

For Each cco In Me.Controls
If Left(UCase(cco.Name), 5) = "text1" Then
axh = "tx(" & LTrim(str(cco.Index)) & ")"
If IsNumber(Placa.Text1(cco.Index).text) Then
axx = FormatNumber((Placa.Text1(cco.Index).text), 2)
End If
ahaa = Replace(ahaa, axh, axx)

End If
Next
If Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'") <> "" Then

delo = str(Round(DateDiff("y", Getnazi("select dat_zap from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"), Date) / 365, 0))
ahaa = Replace(ahaa, "DEL", delo)
End If
'MsgBox Getnazi("select pd from ddoh where zn1>=" & Me.Text2(30).text & " and zn2<=" & Me.Text2(30).text)
If Val(Me.Text2(30).text) > 0 Then
If Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") <> 0 Then
dohh = Getnazi("select top 1 zn from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1")
ddohh = dohh + (Me.Text2(30).text - Getnazi("select zn2 from ddoh order by zn2")) * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
Else
ddohh = Me.Text2(30).text * (Getnazi("select top 1 pd from ddoh where zn1<" & Val(Me.Text2(30).text) & " and zn2>" & Val(Me.Text2(30).text) & " order by zn1") / 100)
End If

ahaa = Replace(ahaa, "DOH", ddohh)
End If
ahaa = Replace(ahaa, "REZI", Getnazi("select rezi from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, "POT", Getnazi("select km from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, "OTROK", Getnazi("select otrok from zaposleni where sifra='" & Trim(Left(Me.List1.text, 6)) & "'"))
ahaa = Replace(ahaa, ",", ".")
If Val(Me.Text2(30).text) < 0 Then
Me.Text2(30).text = "0,00"
Me.Text2(31).text = "0,00"
End If

If Not Trim(ahaa) = "" Then
If Not Left(ahaa, 4) = "fixx" Then
If Not UCase(Left(ahaa, 3)) = "IIF" Then
On Error GoTo vvv:
contr.text = Format(ScriptControl1.Eval(ahaa), "fixed")
vvv:
End If
End If
End If
'contr.text = ahaa
End If
End If
End If

Next

ozna
Me.Timer1.Enabled = False

End Sub

Private Sub Timer1_Timerx()


Dim contr As Control
Dim res As Boolean
Dim ahaa As String
For Each contr In Me.Controls

If contr.Name = "text2" Then

If contr.Index = 19 Then
Placa.Text2(contr.Index).ToolTipText = "tx2(" & contr.Index & ")"
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
Placa.Text2(contr.Index).text = "0,00"
Else
'Placa.Text2(contr.Index).text = Format(Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3)), "fixed")
ahaa = "Placa.Text2(" & contr.Index & ").text =" & (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'"))
'MsgBox ahaa
'EvaluateExpression (ahaa)

   res = Format(Placa.FExecuteCode(ahaa), "fixed")
    'MsgBox res
End If
End If
End If
If contr.Name = "text1" Then

If contr.Index = 20 Then
Placa.Text1(contr.Index).ToolTipText = "tx(" & contr.Index & ")"
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
Placa.Text1(contr.Index).text = "0,00"
Else
ahaa = "Placa.Text1(" & contr.Index & ").text =" & Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3))
    res = Format(FExecuteCode(ahaa), "fixed")
    'MsgBox res
End If
End If
End If
Next

For Each contr In Me.Controls
If Val(Text2(19).text) <> 0 Then
'MsgBox contr.Name
If contr.Name = "text1" Then
'Dim res As Boolean
'Dim ahaa As String
Placa.Text1(contr.Index).ToolTipText = "tx(" & contr.Index & ")"
'MsgBox ehex
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
'Placa.Text1(contr.Index).text = "0,00"
Placa.Text1(contr.Index).BackColor = &HC0FFFF

Else
Placa.Text1(contr.Index).BackColor = &HC0E0FF

ahaa = "Placa.Text1(" & contr.Index & ").text = " & Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='1" & levi_pres(LTrim(str(contr.Index)), 3))
res = Format(FExecuteCode(ahaa), "standard")
    'MsgBox res
End If
If IsNumber(Placa.Text1(contr.Index).text) Then
Placa.Text1(contr.Index).text = Format(Placa.Text1(contr.Index).text, "standard")
End If
End If
End If
Next

For Each contr In Me.Controls
If contr.Name = "text2" Then
Placa.Text2(contr.Index).ToolTipText = "tx2(" & contr.Index & ")"
'MsgBox ehex
If (Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3) & "'")) = "" Then
Placa.Text2(contr.Index).text = "0,00"
Else
ahaa = "Placa.Text2(" & contr.Index & ").text = " & Getnazi("select tekst from dokm where tip_dok='PL' and id_dok='" & Mid(dok.Caption, 3, 6) & "' and atribut='2" & levi_pres(LTrim(str(contr.Index)), 3))
    res = Format(FExecuteCode(ahaa), "fixed")
    'MsgBox res
    
End If

If IsNumber(Placa.Text2(contr.Index).text) Then
Placa.Text2(contr.Index).text = Format(Placa.Text2(contr.Index).text, "standard")
End If
End If
Next


ozna
Me.Timer1.Enabled = False

End Sub
Function EvaluateExpression(strExpression As String) As String
Dim oScriptControl As Object

On Error GoTo e
Set oScriptControl = CreateObject("MSScriptControl.ScriptControl")
'Late bound if not including reference
oScriptControl.Language = "VBScript"
EvaluateExpression = oScriptControl.Eval(strExpression)
Set oScriptControl = Nothing
Exit Function
e:
MsgBox err.Description, vbCritical, "Error"
End Function

Private Sub zapoo_Click()
myConection.Execute "delete from dmat"
Dim drst As New ADODB.Recordset
Dim drst1 As New ADODB.Recordset
drst.Open "select * from zaposleni", myConection, adOpenDynamic, adLockOptimistic
drst1.Open "select * from dmat", myConection, adOpenDynamic, adLockOptimistic

Do While Not drst.EOF
drst1.AddNew
drst1.Fields("sifd") = drst.Fields("sifra")
drst1.Fields("ime") = drst.Fields("ime")
drst1.Fields("priimek") = Left(drst.Fields("priimek"), 15)
drst1.Fields("davst") = drst.Fields("davcna")
drst1.Fields("ulbiv") = Left(drst.Fields("naslov"), 15)
drst1.Fields("datzap") = drst.Fields("dat_zap")
drst1.Fields("rezi") = drst.Fields("rezi")
drst1.Fields("glde") = "1"
drst1.Fields("zs") = "N"
drst1.Fields("EMSO") = drst.Fields("EMSO")
drst1.Fields("sifizp") = "4"
drst1.Update
drst.MoveNext
Loop
End Sub
