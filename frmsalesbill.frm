VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmsalesbill 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLAGAJNA  "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15525
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   13.5
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15525
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton LaVolpeButton43 
      Height          =   375
      Left            =   8760
      TabIndex        =   149
      Top             =   9120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OPOMBA"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0000
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton42 
      Height          =   735
      Left            =   0
      TabIndex        =   148
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "WEB NK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   255
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":001C
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
   Begin ProsVent.UserControl2 UserControl21 
      Height          =   975
      Left            =   2040
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
   End
   Begin ProsVent.xcKeypad xcKeypad1 
      Height          =   3855
      Left            =   3720
      TabIndex        =   142
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6800
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7560
      Top             =   10320
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton37 
      Height          =   375
      Left            =   0
      TabIndex        =   140
      Top             =   10440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "P"
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
      BCOL            =   8454016
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0038
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton36 
      Height          =   255
      Left            =   14280
      TabIndex        =   136
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Vsi pop."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   13.5
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
      MICON           =   "frmsalesbill.frx":0054
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
   Begin VB.TextBox nassl 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   91
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox imes 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   90
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox dav 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6240
      TabIndex        =   88
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   4455
      Left            =   840
      TabIndex        =   84
      Top             =   4440
      Visible         =   0   'False
      Width           =   7815
      Begin LVbuttons.LaVolpeButton zaprr 
         Height          =   495
         Left            =   7080
         TabIndex        =   141
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   13.5
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
         MICON           =   "frmsalesbill.frx":0070
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
      Begin LVbuttons.LaVolpeButton tvorbara 
         Height          =   495
         Left            =   240
         TabIndex        =   139
         Top             =   3840
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "TVORI RAÈUN ZA IZBRANO NAROÈILO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         MICON           =   "frmsalesbill.frx":008C
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton2532 
         Height          =   495
         Left            =   6120
         TabIndex        =   85
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   2
         TX              =   "Delno"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         MICON           =   "frmsalesbill.frx":00A8
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
         BackStyle       =   0  'Transparent
         Caption         =   "Izberi artikle za raèun"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   87
         Top             =   3720
         Width           =   5025
      End
      Begin MSForms.ListBox ListBox1 
         Height          =   3495
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   7575
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "13361;5821"
         MatchEntry      =   0
         MultiSelect     =   1
         FontEffects     =   1073741825
         FontHeight      =   270
         FontCharSet     =   238
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   2760
      TabIndex        =   83
      Text            =   "Uporabnik:"
      Top             =   10680
      Width           =   1395
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   8040
      TabIndex        =   82
      Text            =   "FIRMA:"
      Top             =   10680
      Width           =   915
   End
   Begin VB.CommandButton shran 
      Height          =   270
      Left            =   720
      MaskColor       =   &H8000000F&
      Picture         =   "frmsalesbill.frx":00C4
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   8880
      Width           =   270
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   5640
      TabIndex        =   78
      Text            =   "1"
      Top             =   2880
      Width           =   615
   End
   Begin LVbuttons.LaVolpeButton plu 
      Height          =   495
      Left            =   6960
      TabIndex        =   75
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      MICON           =   "frmsalesbill.frx":01BE
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
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   74
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin LVbuttons.LaVolpeButton reset 
      Height          =   375
      Left            =   14640
      TabIndex        =   70
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "RESET"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmsalesbill.frx":01DA
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
   Begin VB.ListBox veli 
      BackColor       =   &H00C0FFC0&
      Height          =   1005
      Left            =   15360
      TabIndex        =   68
      Top             =   840
      Width           =   195
   End
   Begin VB.TextBox nazivv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   9120
      TabIndex        =   63
      Top             =   10680
      Width           =   5235
   End
   Begin VB.TextBox pop 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14640
      TabIndex        =   56
      Text            =   "0"
      Top             =   8040
      Width           =   735
   End
   Begin LVbuttons.LaVolpeButton VRNIT 
      Height          =   495
      Left            =   10080
      TabIndex        =   55
      Top             =   8520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Vrniti - F9"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":01F6
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
   Begin LVbuttons.LaVolpeButton prija 
      Height          =   375
      Left            =   8760
      TabIndex        =   53
      Top             =   8040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Prijava"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0212
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2522 
      Height          =   975
      Left            =   11760
      TabIndex        =   52
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Delni - F6"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":022E
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton251 
      Height          =   735
      Left            =   13800
      TabIndex        =   50
      Top             =   9600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Zakljucek"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":024A
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
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   540
      Left            =   15600
      ScaleHeight     =   480
      ScaleWidth      =   255
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   315
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing... Please wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -240
         TabIndex        =   49
         Top             =   480
         Width           =   3405
      End
   End
   Begin VB.TextBox mii 
      Height          =   435
      Left            =   0
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   8760
      TabIndex        =   3
      Top             =   2880
      Width           =   6615
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   735
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0266
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
         Height          =   735
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0282
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
         Height          =   735
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":029E
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
         Height          =   735
         Left            =   0
         TabIndex        =   7
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":02BA
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
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   3600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":02D6
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
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   2880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":02F2
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton7 
         Height          =   735
         Left            =   1320
         TabIndex        =   10
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":030E
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
         Height          =   735
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":032A
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton9 
         Height          =   735
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0346
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton10 
         Height          =   735
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0362
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton11 
         Height          =   735
         Left            =   1320
         TabIndex        =   14
         Top             =   2880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":037E
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton12 
         Height          =   735
         Left            =   1320
         TabIndex        =   15
         Top             =   3600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":039A
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton13 
         Height          =   735
         Left            =   2640
         TabIndex        =   16
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":03B6
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton14 
         Height          =   735
         Left            =   2640
         TabIndex        =   17
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":03D2
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton15 
         Height          =   735
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":03EE
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton16 
         Height          =   735
         Left            =   2640
         TabIndex        =   19
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":040A
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton17 
         Height          =   735
         Left            =   2640
         TabIndex        =   20
         Top             =   2880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0426
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton18 
         Height          =   735
         Left            =   2640
         TabIndex        =   21
         Top             =   3600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0442
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton19 
         Height          =   735
         Left            =   3960
         TabIndex        =   22
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":045E
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton20 
         Height          =   735
         Left            =   3960
         TabIndex        =   23
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":047A
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton21 
         Height          =   735
         Left            =   3960
         TabIndex        =   24
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0496
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton22 
         Height          =   735
         Left            =   3960
         TabIndex        =   25
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":04B2
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton23 
         Height          =   735
         Left            =   3960
         TabIndex        =   26
         Top             =   2880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":04CE
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton24 
         Height          =   735
         Left            =   3960
         TabIndex        =   27
         Top             =   3600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":04EA
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton25 
         Height          =   735
         Left            =   0
         TabIndex        =   122
         Top             =   4320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0506
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton26 
         Height          =   735
         Left            =   1320
         TabIndex        =   123
         Top             =   4320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0522
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton27 
         Height          =   735
         Left            =   2640
         TabIndex        =   124
         Top             =   4320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":053E
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton28 
         Height          =   735
         Left            =   3960
         TabIndex        =   125
         Top             =   4320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":055A
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton29 
         Height          =   735
         Left            =   5280
         TabIndex        =   126
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0576
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton30 
         Height          =   735
         Left            =   5280
         TabIndex        =   127
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0592
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton31 
         Height          =   735
         Left            =   5280
         TabIndex        =   128
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":05AE
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton32 
         Height          =   735
         Left            =   5280
         TabIndex        =   129
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":05CA
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton33 
         Height          =   735
         Left            =   5280
         TabIndex        =   130
         Top             =   2880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":05E6
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton34 
         Height          =   735
         Left            =   5280
         TabIndex        =   131
         Top             =   3600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":0602
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton35 
         Height          =   735
         Left            =   5280
         TabIndex        =   132
         Top             =   4320
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "LaVolpeButton"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmsalesbill.frx":061E
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
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   9840
   End
   Begin VB.TextBox txtInvoiceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Left            =   3240
      TabIndex        =   1
      Text            =   "1"
      Top             =   0
      Width           =   1995
   End
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   0
      Left            =   720
      TabIndex        =   29
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":063A
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton46 
      Height          =   30
      Left            =   1920
      TabIndex        =   30
      Top             =   10185
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   53
      BTYPE           =   2
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmsalesbill.frx":0656
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton45 
      Height          =   735
      Left            =   8760
      TabIndex        =   31
      Top             =   9600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "STORNO - F3"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0672
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
   Begin LVbuttons.LaVolpeButton LaVo1 
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   9360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Mize - F6"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":068E
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
   Begin LVbuttons.LaVolpeButton LaVo2 
      Height          =   735
      Left            =   10560
      TabIndex        =   33
      Top             =   9600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Stranka - F8"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":06AA
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton44 
      Height          =   735
      Left            =   12240
      TabIndex        =   34
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "IZHOD - F10"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":06C6
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   36
      Top             =   960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":06E2
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   37
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":06FE
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   3
      Left            =   0
      TabIndex        =   38
      Top             =   2640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":071A
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   39
      Top             =   3480
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0736
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   5
      Left            =   0
      TabIndex        =   40
      Top             =   4320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0752
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   6
      Left            =   0
      TabIndex        =   41
      Top             =   5160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":076E
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   7
      Left            =   0
      TabIndex        =   42
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":078A
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   8
      Left            =   0
      TabIndex        =   43
      Top             =   6840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":07A6
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   9
      Left            =   0
      TabIndex        =   44
      Top             =   7680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":07C2
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
   Begin LVbuttons.LaVolpeButton mizaa 
      Height          =   735
      Index           =   10
      Left            =   0
      TabIndex        =   45
      Top             =   8520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":07DE
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
   Begin LVbuttons.LaVolpeButton vst5 
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   9960
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   450
      BTYPE           =   2
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmsalesbill.frx":07FA
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
   Begin LVbuttons.LaVolpeButton pred 
      Height          =   495
      Left            =   8760
      TabIndex        =   61
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Predal"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0816
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
   Begin LVbuttons.LaVolpeButton levog 
      Height          =   615
      Left            =   13305
      TabIndex        =   64
      Top             =   8040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "==>"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0832
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
   Begin LVbuttons.LaVolpeButton desnog 
      Height          =   615
      Left            =   13305
      TabIndex        =   65
      Top             =   8760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "<=="
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":084E
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
   Begin LVbuttons.LaVolpeButton zakljucc 
      Height          =   1095
      Index           =   0
      Left            =   1680
      TabIndex        =   73
      Top             =   9360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "ZAKLJUCI - F4"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":086A
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
   Begin LVbuttons.LaVolpeButton min 
      Height          =   495
      Left            =   7560
      TabIndex        =   76
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      MICON           =   "frmsalesbill.frx":0886
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
   Begin LVbuttons.LaVolpeButton bri 
      Height          =   495
      Left            =   8160
      TabIndex        =   77
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      MICON           =   "frmsalesbill.frx":08A2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmsalesbill.frx":08BE
      Height          =   5520
      Left            =   720
      TabIndex        =   79
      Top             =   3600
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   9737
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   49152
      ForeColorSel    =   16777215
      BackColorUnpopulated=   16777152
      GridColor       =   12632256
      GridColorFixed  =   16777215
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLines       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   1
      Left            =   2040
      TabIndex        =   92
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0BC8
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   2
      Left            =   3360
      TabIndex        =   93
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0BE4
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   3
      Left            =   4680
      TabIndex        =   94
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C00
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   4
      Left            =   6000
      TabIndex        =   95
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C1C
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   5
      Left            =   7320
      TabIndex        =   96
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C38
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   6
      Left            =   8640
      TabIndex        =   97
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C54
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   7
      Left            =   9960
      TabIndex        =   98
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C70
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   8
      Left            =   11280
      TabIndex        =   99
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0C8C
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   9
      Left            =   12600
      TabIndex        =   100
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0CA8
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   10
      Left            =   720
      TabIndex        =   101
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0CC4
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   11
      Left            =   2040
      TabIndex        =   102
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0CE0
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   12
      Left            =   3360
      TabIndex        =   103
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0CFC
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   13
      Left            =   4680
      TabIndex        =   104
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0D18
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   14
      Left            =   6000
      TabIndex        =   105
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0D34
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   15
      Left            =   7320
      TabIndex        =   106
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0D50
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   16
      Left            =   8640
      TabIndex        =   107
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0D6C
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   17
      Left            =   9960
      TabIndex        =   108
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0D88
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   18
      Left            =   11280
      TabIndex        =   109
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0DA4
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   19
      Left            =   12600
      TabIndex        =   110
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0DC0
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   20
      Left            =   720
      TabIndex        =   111
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0DDC
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   21
      Left            =   2040
      TabIndex        =   112
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0DF8
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   22
      Left            =   3360
      TabIndex        =   113
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0E14
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   23
      Left            =   4680
      TabIndex        =   114
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0E30
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   24
      Left            =   6000
      TabIndex        =   115
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0E4C
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   25
      Left            =   7320
      TabIndex        =   116
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0E68
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   26
      Left            =   8640
      TabIndex        =   117
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0E84
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   27
      Left            =   9960
      TabIndex        =   118
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0EA0
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   28
      Left            =   11280
      TabIndex        =   119
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0EBC
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   29
      Left            =   12600
      TabIndex        =   120
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0ED8
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
   Begin LVbuttons.LaVolpeButton pregled 
      Height          =   615
      Left            =   14280
      TabIndex        =   121
      Top             =   8760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PREG"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0EF4
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   30
      Left            =   13920
      TabIndex        =   133
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0F10
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   31
      Left            =   13920
      TabIndex        =   134
      Top             =   1560
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0F2C
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
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   630
      Index           =   32
      Left            =   13920
      TabIndex        =   135
      Top             =   2160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0F48
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
   Begin LVbuttons.LaVolpeButton vena 
      Height          =   615
      Left            =   4200
      TabIndex        =   137
      Top             =   9360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "NACIN PLACILA"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0F64
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
   Begin LVbuttons.LaVolpeButton narooc 
      Height          =   495
      Left            =   0
      TabIndex        =   138
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "NAROÈILA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
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
      MICON           =   "frmsalesbill.frx":0F80
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton38 
      Height          =   495
      Left            =   6360
      TabIndex        =   143
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "T"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   17.25
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
      MICON           =   "frmsalesbill.frx":0F9C
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton39 
      Height          =   375
      Left            =   14640
      TabIndex        =   145
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "S+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmsalesbill.frx":0FB8
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton40 
      Height          =   375
      Left            =   15000
      TabIndex        =   146
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "S-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "frmsalesbill.frx":0FD4
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton41 
      Height          =   375
      Left            =   4200
      TabIndex        =   147
      Top             =   10080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Dobavnice"
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
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsalesbill.frx":0FF0
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
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "INTERNA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   144
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label davlb 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ID.ST.:SI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5160
      TabIndex        =   89
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label znees 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   44.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   4800
      TabIndex        =   80
      Top             =   9240
      Width           =   3855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   72
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "VSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   14640
      TabIndex        =   69
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      Height          =   495
      Left            =   14280
      TabIndex        =   67
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   495
      Left            =   14280
      TabIndex        =   66
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSForms.CheckBox inter 
      Height          =   375
      Left            =   11640
      TabIndex        =   62
      Top             =   9120
      Width           =   1815
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3201;661"
      Value           =   "0"
      Caption         =   "INTERNA"
      FontEffects     =   1073741825
      FontHeight      =   270
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DOD"
      Height          =   375
      Left            =   10680
      TabIndex        =   59
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   11280
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "POP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14160
      TabIndex        =   57
      Top             =   8160
      Width           =   615
   End
   Begin MSForms.CheckBox kart 
      Height          =   375
      Left            =   11040
      TabIndex        =   54
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "873;661"
      Value           =   "0"
      Caption         =   "KARTICA-F2"
      FontEffects     =   1073741825
      FontHeight      =   270
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   4320
      TabIndex        =   51
      Top             =   10680
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MIZA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   46
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbst 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11400
      TabIndex        =   35
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label LblDateTime 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rac.st.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "frmsalesbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim gSlno, gItemCode, gItemname, gQty, gRate, gTotal, gpop, Inti, miz, i
Dim Indx
Private nacra As Integer
Public ahha As Long
Private Sub cmbItmcode_LostFocus()

'MsgBox ("0")
If fora = 0 Then
If deln = 1 Then
Else
'SendKeys 1

prvaa = 1
kolik = 1
End If
'MSHFlexGrid1.TextMatrix(Indx, 4) = 1
Else
fora = 0
End If
End Sub


Private Sub bri_Click()
myConection.Execute ("delete from trenutna where stdok='" & Pblagajna & "' and  x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssv
End Sub

Private Sub dav_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Dim idstr, nall As String
idstr = Getnazi("select ime from po where dav='" & Me.dav.Text & "'")
nall = Getnazi("select nasl from po where dav='" & Me.dav.Text & "'")
If idstr = "" Then
nall = Getnazi("select nasl from fozD where dav='" & Me.dav.Text & "'")
idstr = Getnazi("select ime from fozD where dav='" & Me.dav.Text & "'")
End If
If idstr = "" Then
nall = Getnazi("select NAZIV from PARTNER where davcna='" & Me.dav.Text & "'") & " " & Getnazi("select ulica from PARTNER where davcna='" & Me.dav.Text & "'")
idstr = Getnazi("select posta from PARTNER where davcna='" & Me.dav.Text & "'") & " " & Getnazi("select mesto from PARTNER where davcna='" & Me.dav.Text & "'")

End If

Me.imes.Text = idstr
Me.nassl.Text = nall
End If

End Sub

Private Sub desnog_Click()
If Val(Me.Label8.Caption) - 24 > 0 Then
Me.Label8.Caption = str(Val(Me.Label8.Caption) - 24)
Else
Me.Label8.Caption = "0"
End If
Dim q As Integer
q = Val(Me.Label9.Caption)
Hanb (q)
End Sub

Private Sub Form_Activate()
If jefield("mize", "naziv") = False Then
myConection.Execute ("ALTER TABLE mize ADD naziv text 250")
End If


If xzago <> 1 Then
coda852
blagajna = 1
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
For miz = 1 To 10
mizaa(miz).Caption = miz + stevmi
mizaa(miz).BackColor = 14215660

Next
mi

txtInvoiceNo.Text = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)
nazivv.Text = Getnazi("select glava1 from oblikar")
If rs.State = 1 Then rs.Close
   rs.Open "select * from swit WHERE [ItemNumber] > 0 AND [Switchboar]=1 order by [ItemNumber]"
      rs.MoveFirst
      Dim aad As Integer
      aad = 0
      If Not rs.EOF Then

       While (Not (rs.EOF))
       
         nas1(aad).Caption = rs![ITEMTEXT]
         If Val(Getnazi("select poz from dokm where id_dok='" & rs![ITEMTEXT] & "'")) <> 0 Then
         nas1(aad).BackColor = Val(Getnazi("select poz from dokm where id_dok='" & rs![ITEMTEXT] & "'"))
        End If
        If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='BOLDDA'") = "D" Then
        nas1(aad).Font.Bold = True
        Else
        nas1(aad).Font.Bold = False
        End If
         nas1(aad).Tag = rs![ARGUMENT]
        
        aad = aad + 1
            rs.MoveNext
        Wend
        aad = 0
      Do While Not aad = 33
      
      If nas1(aad).Tag = "" Then
      nas1(aad).Visible = False
      End If
      aad = aad + 1
      Loop
      Else
         End If
        Hanb (1)
txtInvoiceNo.Text = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)
Else
xzago = 0
End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'") = "D" Then
Me.LaVolpeButton45.Visible = False
Me.bri.Visible = False
Else
Me.LaVolpeButton45.Visible = True
Me.bri.Visible = True

End If
If Getnumb("select nivo from  users where up='" & LTrim(RTrim(prijavljen)) & "'") > 1 Then
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='ZAKPA'") = "D" Then
Me.LaVolpeButton251.Visible = True
Else
Me.LaVolpeButton251.Visible = False

End If
End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SPLETDA'") = "D" Then
Me.LaVolpeButton37.Visible = True
Me.narooc.Visible = True
Else
Me.LaVolpeButton37.Visible = False
Me.narooc.Visible = False

End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='WEBDA'") = "D" Then
Me.LaVolpeButton42.Visible = True
Else
Me.LaVolpeButton42.Visible = False
End If
osssv
End Sub
Private Sub LoadFlexGridColumnWidths(ByVal flx As MSHFlexGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Cols - 1
        ' Get the column width. Use its current
        ' width as the default value.
        
        flx.ColWidth(i) = GetSetting( _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.ColWidth(i))
    Next i
 
End Sub
Private Sub SaveFlexGridColumnWidths(ByVal flx As MSHFlexGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Cols - 1
        ' Save the column width.
        SaveSetting _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.ColWidth(i)
    Next i
End Sub
Private Sub Form_Load()
ReSizeForm Me
Dim X, velf, velf1 As Integer
velf = velifont()
For velf1 = 0 To velf
For X = 0 To 32

Me.nas1(X).ButtonType = [Windows 32-bit]

Me.nas1(X).Font.Size = Me.nas1(X).Font.Size + 1
Next
For X = 1 To 35
     
      Me("LaVolpeButton" & X).ButtonType = [Windows 32-bit]
      Me("LaVolpeButton" & X).Font.Size = Me("LaVolpeButton" & X).Font.Size + 1
      Next
Next

'MsfRefresh
'FillCombo cmbItmcode, "select MADASIFR from MADA"
 
End Sub




Private Sub Image1_Click()
End
End Sub


Private Sub inter_Click()
Dim rrrt As New ADODB.Recordset
If rrrt.State = 1 Then rrrt.Close


 rrrt.Open "select * from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
 If Not rrrt.EOF Then
 rrrt.MoveFirst
Xvs = Me.MSHFlexGrid1.Rows - 1
 Yvs = 1
Me.UserControl21.opentime
Me.UserControl21.Visible = True
 Do While Not rrrt.EOF
 DoEvents
If Me.inter.Value = True Then
 rrrt.Fields("cena") = Getcena(rrrt.Fields("sifra"), Date)
 ' Getcena(rrrt.Fields("sifra"), Date)
Else
 rrrt.Fields("cena") = Getnumb("select madampcd from mada where madasifr='" & rrrt.Fields("sifra") & "'")
End If
 rrrt.Fields("znes") = rrrt.Fields("kol") * rrrt.Fields("cena")
 rrrt.Fields("pop") = 0
 rrrt.Update
 rrrt.MoveNext
 Yvs = Yvs + 1
 Loop
 End If
 If Me.inter.Value = True Then
Me.Label13.Visible = True
Else
Me.Label13.Visible = False
End If
Me.UserControl21.closetime
Me.UserControl21.Visible = False
osssv
End Sub

Private Sub LaVo1_Click()
If Me.Label11.Caption = "" Then
mize.Show vbModal
Else
shranimi (Me.Label11.Caption)
End If
End Sub

Private Sub LaVo2_Click()
Me.dav.Visible = True
Me.davlb.Visible = True
Me.imes.Visible = True
Me.nassl.Visible = True
Me.dav.SetFocus


End Sub

Private Sub LaVolpeButton1_Click()
Hanbt (1)
End Sub

Private Sub LaVolpeButton10_Click()
Hanbt (10)
End Sub

Private Sub LaVolpeButton11_Click()
Hanbt (11)
End Sub

Private Sub LaVolpeButton12_Click()
Hanbt (12)
End Sub

Private Sub LaVolpeButton13_Click()
Hanbt (13)
End Sub

Private Sub LaVolpeButton14_Click()
Hanbt (14)
End Sub

Private Sub LaVolpeButton15_Click()
Hanbt (15)
End Sub

Private Sub LaVolpeButton16_Click()
Hanbt (16)
End Sub

Private Sub LaVolpeButton17_Click()
Hanbt (17)
End Sub

Private Sub LaVolpeButton18_Click()
Hanbt (18)
End Sub

Private Sub LaVolpeButton19_Click()
Hanbt (19)
End Sub

Private Sub LaVolpeButton2_Click()
Hanbt (2)
End Sub

Private Sub LaVolpeButton20_Click()
Hanbt (20)
End Sub

Private Sub LaVolpeButton21_Click()
Hanbt (21)
End Sub

Private Sub LaVolpeButton22_Click()
Hanbt (22)
End Sub

Private Sub LaVolpeButton23_Click()
Hanbt (23)
End Sub

Private Sub LaVolpeButton24_Click()
Hanbt (24)
End Sub

Private Sub LaVolpeButton25_Click()
' If rs.State = 1 Then rs.Close
'   rs.Open "select * from swit WHERE [ItemNumber] > 0 AND [Switchboar]=1 order by itemnumber"
'      rs.MoveFirst
'      Dim aad As Integer
'      aad = 0
'      If Not rs.EOF Then

'       While (Not (rs.EOF))
'       aad = aad + 1
''         Me("nas" & aad).Caption = rs![ITEMTEXT]
'         Me("nas" & aad).Tag = rs![SWITCHBOAR]
'            rs.MoveNext
'        Wend
'      Else
'         End If
Hanbt (25)


End Sub

Private Sub LaVolpeButton251_Click()
OSE = Me.Label3.Caption
Form3.Show

End Sub

Private Sub LaVolpeButton2532_Click()
deln = 1
   
    Dim strItems As String
    Dim intCurrentRow As Integer
    Dim po As Integer
    Dim a As Integer
    Dim b As Integer
    
   Call LaVolpeButton45_Click
   
   
   
   Dim aaa As String
aaa = Left(Time(), 8)
'MsgBox (aaa)
   If rs.State = 1 Then rs.Close
   
 
rs.Open "select sifra,kol,znesek,datum,ura,stmize,stdok from mize", myConection


  
    For intCurrentRow = 0 To ListBox1.ListCount - 1
       
            
    a = Val(Left(ListBox1.Column(0, intCurrentRow), 13))
    b = Val(Right(ListBox1.Column(0, intCurrentRow), 6))
    If ListBox1.Selected(intCurrentRow) Then
    Sendkeys a & "{enter}+{RIGHT}{BS}" & b & "{enter}"
        '
        '  MSHFlexGrid1.TextMatrix(Indx, 0) = Indx
          
                 'MSHFlexGrid1.TextMatrix(Indx, 0) = Indx
'MSHFlexGrid1.TextMatrix(Indx, 1) = Left(ListBox1.Column(0, intCurrentRow), 13)
'MSHFlexGrid1.TextMatrix(Indx, 2) = Getnazi("select madanazi from mada where madasifr=" & Left(ListBox1.Column(0, intCurrentRow), 13))
'MSHFlexGrid1.TextMatrix(Indx, 4) = Right(ListBox1.Column(0, intCurrentRow), 6)
'Indx = Indx + 1
'po = po + 1
'MSHFlexGrid1.Row = po
Else
If stm1 <> 0 Then
If a <> 0 Then
Dim cen As Double
cen = Getnazi("select madampcd from mada where madasifr='" & a & "'")
rs.AddNew
    rs.Fields(0) = a
    rs.Fields(1) = b
    rs.Fields(2) = b * cen 'Val(MSHFlexGrid1.TextMatrix(i, 5))
    rs.Fields(3) = Date
    rs.Fields(4) = aaa
      rs.Fields(5) = stm1
      rs.Fields(6) = Pblagajna
      rs.Update
End If
End If
    
 
        End If
      
       ' zap = Indx
'          fora = 2
    Next intCurrentRow
rs.Close
'       fora = 2
Me.ListBox1.clear
refr = 1
stm1 = 0
   
Me.Frame2.Visible = False
deln = 0
End Sub

Private Sub LaVolpeButton2522_Click()
If LTrim(RTrim(Me.Label11.Caption)) = "" Then
MsgBox ("Delni deluje le èe je odprta miza!")
Else
Me.Frame2.Visible = True
Dim i
With ListBox1
For i = 1 To Me.MSHFlexGrid1.Rows - 1
.AddItem presled(MSHFlexGrid1.TextMatrix(i, 0), 13) & "  " & presled(MSHFlexGrid1.TextMatrix(i, 1), 17) & "      " & MSHFlexGrid1.TextMatrix(i, 2)
 Next
End With
Me.ListBox1.SetFocus
End If
End Sub

Private Sub LaVolpeButton26_Click()
Hanbt (26)
End Sub

Private Sub LaVolpeButton27_Click()
Hanbt (27)
End Sub

Private Sub LaVolpeButton28_Click()
Hanbt (28)
End Sub

Private Sub LaVolpeButton29_Click()
Hanbt (29)
End Sub

Private Sub LaVolpeButton3_Click()
Hanbt (3)
End Sub

Private Sub LaVolpeButton30_Click()
Hanbt (30)
End Sub

Private Sub LaVolpeButton31_Click()
Hanbt (31)
End Sub

Private Sub LaVolpeButton32_Click()
Hanbt (32)
End Sub

Private Sub LaVolpeButton33_Click()
Hanbt (33)
End Sub

Private Sub LaVolpeButton34_Click()
Hanbt (34)
End Sub

Private Sub LaVolpeButton35_Click()
Hanbt (35)
End Sub

Private Sub LaVolpeButton36_Click()
Me.pop.Text = xpopu
If MsgBox("Ali nastavim vsem pozicijam popust " & str(Me.pop.Text) & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then

myConection.Execute ("update trenutna set pop=" & Me.pop.Text & " where stdok='" & Pblagajna & "'")
'myConection.Execute ("update trenutna set cena=cena*(1-(pop/100)) where x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=cena*(1-(pop/100))*kol where stdok='" & Pblagajna & "'")
osssv
 
End If
End Sub

Private Sub LaVolpeButton37_Click()

ShellExecute 0&, "open", App.path & "\prennar.bat", "", "", 0
On Error GoTo bbb:
Dim fso As scripting.FileSystemObject
    Set fso = New scripting.FileSystemObject
 
    fso.MoveFile App.path & "\*.xml", App.path & "\naro"
bbb:
Sleep 1000
If rs.State = 1 Then rs.Close
rs.Open "select dobavit_id,sum(madazalo) as zalo from mada where dobavit_id<>'' group by dobavit_id", myConection, adOpenDynamic, adLockOptimistic
   
  Open App.path & "\stock.xml" For Output As #1
  Print #1, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
Print #1, "<articles>"
    

   rs.MoveFirst
   Do While Not rs.EOF
   Print #1, "  <article>"
   Print #1, "     <quantity>" & rs.Fields("zalo") & "</quantity>"
   Print #1, "     <code>" & Replace(Replace(rs.Fields("dobavit_id"), "\", ""), "/", "") & "</code>"
   Print #1, "  </article>"
    
   rs.MoveNext
   Loop
Print #1, "</articles>"
Close #1
Sleep 1000
ShellExecute 0&, "open", App.path & "\prenzal.bat", "", "", 0



End Sub

Private Sub LaVolpeButton38_Click()
If Me.xcKeypad1.Visible = True Then
Me.xcKeypad1.Visible = False
Me.Text3.Text = 1
Else
Me.xcKeypad1.Visible = True
Me.Text3.Text = ""
End If
End Sub

Private Sub LaVolpeButton39_Click()
Dim X, cc As Integer
cc = GetSetting("bll", "velfont", "blg", 0)
        
      SaveSetting "bll", "velfont", "blg", cc + 1
For X = 0 To 32

Me.nas1(X).ButtonType = [Windows 32-bit]

Me.nas1(X).Font.Size = Me.nas1(X).Font.Size + 1
Next
For X = 1 To 35
     
      Me("LaVolpeButton" & X).ButtonType = [Windows 32-bit]
      Me("LaVolpeButton" & X).Font.Size = Me("LaVolpeButton" & X).Font.Size + 1
      Next
      
End Sub

Private Sub LaVolpeButton4_Click()
Hanbt (4)
End Sub

Private Sub LaVolpeButton40_Click()
Dim X, cc1 As Integer
cc1 = GetSetting("bll", "velfont", "blg", 0)
        
      SaveSetting "bll", "velfont", "blg", cc1 - 1

For X = 0 To 32
Me.nas1(X).ButtonType = [Windows 32-bit]

Me.nas1(X).Font.Size = Me.nas1(X).Font.Size - 1
Next
For X = 1 To 35
     
      Me("LaVolpeButton" & X).ButtonType = [Windows 32-bit]
      Me("LaVolpeButton" & X).Font.Size = Me("LaVolpeButton" & X).Font.Size - 1
      Next
End Sub

Private Sub LaVolpeButton41_Click()
dobavnice.Show vbModal
End Sub

Private Sub LaVolpeButton42_Click()
nar_web.Show vbModal
End Sub

Private Sub LaVolpeButton43_Click()
'myConection.Execute ("update trenutna set naziv='" & Replace(Me.Text2.Text, "'", "") & "' where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
dolzina = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5)
DOD_OPO.Show vbModal
End Sub

Private Sub LaVolpeButton44_Click()
'End
blagajna = 0

Unload Me
End Sub

Public Sub LaVolpeButton45_Click()
Timer1.Enabled = True
If stalnaprij = 1 Then
Else
prijavljen = ""
End If
Me.kart.Value = 0
Me.inter.Value = 0
myConection.Execute ("delete from trenutna where  stdok='" & Pblagajna & "'")
If nacra = 0 Then
myConection.Execute ("delete from nacplac where dokument='PA" & Me.txtInvoiceNo.Text & "'")
Else
nacra = 0
End If
'Me.UserControl11.Visible = False
'Me.karto.Visible = False
'Me.stranka.Visible = False
'Me.karto.Visible = False
Dim stot, fa
Indx = 1

zap = 1
Me.MSHFlexGrid1.clear
osssv
idstran = 0
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
For miz = 1 To 10
mizaa(miz).Caption = miz + stevmi
mizaa(miz).BackColor = 14215660

Next

mi
On Error GoTo ppp
Indx = 1
zap = 0
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE F6 "
Me.Label12.Caption = ""
Me.Text1.Text = ""
Text2.Text = ""
Text3.Text = 1
'Text2.SetFocus

Me.dav.Visible = False
Me.dav.Text = ""
Me.davlb.Visible = False
Me.imes.Visible = False
Me.imes.Text = ""
Me.nassl.Visible = False
Me.nassl.Text = ""
Me.Label13.Visible = False

Text1.SetFocus
ppp:
End Sub

Private Sub LaVolpeButton46_Click()
'If Me.MSHFlexGrid1.Col = 4 Then
'End If

If Frame2.Visible = True Then
Exit Sub
End If

If Me.dav.Visible = True Then
'idstran = Me.dav.text
End If

    Dim strf As Integer
    If Me.kart.Value = True Then
     plax = "KARTICA"
     
    strf = 1
    Else
    strf = 0
     plax = "GOTOVINA"
    End If
    If strf = 0 Then
    If Me.inter.Value = True Then
      plax = "INTERNA     Podpis ______________________"
    Else
      plax = "GOTOVINA"
    
    End If
    End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='POPPA'") = "D" Then
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PARA2'") = "D" Then
printrac
printrac
Else
printrac
End If
Else
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PARA2'") = "D" Then
printrac2
printrac2
Else
printrac2
End If
End If
Dim i, stot, fa
Dim aaa As String

aaa = Left(Time(), 8)
'MsgBox (aaa)
Dim Rsa As New ADODB.Recordset
   If Rsa.State = 1 Then Rsa.Close

 
Rsa.Open "select * from nabasif", myConection, adOpenStatic, adLockOptimistic
Dim ddd As Integer

Dim ddid As String
ddid = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)
    Dim dass
    Dim datum As String
    
dass = Format(Now, "dd.mm.yyyy hh:mm:ss")
datum = Left(dass, 2) & "." & Mid(dass, 4, 2) & "." & Mid(dass, 7, 4) & " " & Mid(dass, 12, 2) & ":" & Mid(dass, 15, 2) & ":" & Mid(dass, 18, 2)
Dim xdoz As Double
    
For i = 1 To Me.MSHFlexGrid1.Rows - 1
Rsa.AddNew
    Rsa.Fields("SIFRA") = Val(MSHFlexGrid1.TextMatrix(i, 0))
    Rsa.Fields("NAZIV") = MSHFlexGrid1.TextMatrix(i, 1)
    Rsa.Fields("KOL") = Val(MSHFlexGrid1.TextMatrix(i, 2))
    If Val(MSHFlexGrid1.TextMatrix(i, 6)) = 0 Then
    Rsa.Fields("pop") = 0
    Else
    Rsa.Fields("pop") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 6), 2)
    End If
    Rsa.Fields("mpc") = Getcena((MSHFlexGrid1.TextMatrix(i, 0)), Date)


    Rsa.Fields("ZNES") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    Rsa.Fields("cena") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 3), 2)
    
    Rsa.Fields("pozicija") = levi_pres(LTrim(str(i)), 4)
    Rsa.Fields("DAT_K") = datum
     Rsa.Fields("skl") = "GOS"
    Rsa.Fields("DATUM") = Date
    If Me.dav.Text <> "" Then
     Rsa.Fields("org") = Me.dav.Text
     End If
    'If Getnumb("select madadoza from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'") <> "" Then
    xdoz = Getnazi("select madadoza from mada where madasifr='" & Val(MSHFlexGrid1.TextMatrix(i, 0)) & "'")
    Rsa.Fields("DOZA") = xdoz
    'End If
    'MsgBox Getnazi("select madadoza from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 1)) & "'")
    Rsa.Fields("TIP_DOK") = "PA"
    Rsa.Fields("ID_DOK") = ddid
    Rsa.Fields("stdok") = Pblagajna
      Rsa.Fields("SIFRAPART") = ddd
        Rsa.Fields("faktor") = Getnazi("select faktor from dokumenti where tip_dok='PA'")
        Rsa.Fields("UPORABNIK") = RTrim(prijavljen)
        If Me.kart.Value = True Then
                Rsa.Fields("PLACILO") = 9999
                
        End If

        If Me.inter.Value = True Then
                Rsa.Fields("PLACILO") = 1
                Rsa.Fields("ZNES") = Val(MSHFlexGrid1.TextMatrix(i, 2)) * Getcena((MSHFlexGrid1.TextMatrix(i, 0)), Date)
                Rsa.Fields("CENA") = Getcena((MSHFlexGrid1.TextMatrix(i, 0)), Date)
        End If

Rsa.Update
'opisi iz nk
 If stevnaro <> "" Then
'MsgBox ("select tekst from dok_m where tip_dok+id_dok='" & stevnaro & "' and atribut='" & levi_pres(LTrim(str(i)), 4) & "'")
 myConection.Execute ("insert INTO DOKM (tip_dok,id_dok,ATRIBUT,TEKST) values ('PA','" & ddid & "','" & levi_pres(LTrim(str(i)), 4) & "','" & Getnazi("select tekst from dokm where tip_dok+id_dok='" & stevnaro & "' and atribut='" & levi_pres(LTrim(str(i)), 4) & "'") & "')")
End If
Dim SQLLL As String
Dim ses_ko As String
Dim ses_si As String
ses_ko = 0
ses_si = ""
If Getnazi("select sifras from sestavi where sifra=" & Trim(Rsa.Fields("sifra"))) = "" Then
SQLLL = "update mada set madazalo=madazalo-" & Replace(Rsa.Fields("KOL") * IIf(xdoz > 0, xdoz, 1), ",", ".") & " where madasifr='" & Trim(Rsa.Fields("sifra")) & "'"
myConection.Execute (SQLLL)

Else
If rs.State = 1 Then rs.Close
rs.Open "select * from sestavi where sifra=" & Trim(Rsa.Fields("sifra")), myConection, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Do While Not rs.EOF
ses_si = rs.Fields("sifras")
ses_ko = Replace(FormatNumber(Rsa.Fields("KOL") * rs.Fields("kol"), 2), ",", ".")
SQLLL = "update mada set madazalo=madazalo-" & ses_ko & " where madasifr='" & Trim(ses_si) & "'"
myConection.Execute (SQLLL)

rs.MoveNext
Loop
End If
    
       
 
    Next
    Dim nazs, nass, posts, davv As String
    
 If stevnaro <> "" Then
 nazs = Getnazi("select dod0 from glavna where tip_dok+id_dok='" & stevnaro & "'")
 
 nass = Getnazi("select naslov from partner where naziv='" & nazs & "'")
 posts = Trim((Getnazi("select posta from partner where naziv='" & nazs & "'"))) & " " & Trim((Getnazi("select mesto from partner where naziv='" & nazs & "'")))
 End If
 myConection.Execute ("insert into glavna (tip_dok,id_dok,skl,dod0,dod1,dod2) values ('PA','" & ddid & "','GOS','" & nazs & "','" & nass & "','" & posts & "')")
      
 stevnaro = ""
 Rsa.Close
Indx = 1
zap = 1
osssv
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
For miz = 1 To 10
mizaa(miz).Caption = miz + stevmi
mizaa(miz).BackColor = 14215660

Next
'Me.UserControl11.Visible = False
mi
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE F6 "
Me.Label12.Caption = ""
LaVolpeButton45_Click
End Sub

Private Sub LaVolpeButton5_Click()
Hanbt (5)
End Sub

Private Sub LaVolpeButton6_Click()
Hanbt (6)
End Sub

Private Sub LaVolpeButton7_Click()
Hanbt (7)
End Sub

Private Sub LaVolpeButton8_Click()
Hanbt (8)
End Sub

Private Sub LaVolpeButton9_Click()
Hanbt (9)
End Sub



    



Private Sub levog_Click()
Me.Label8.Caption = str(Val(Me.Label8.Caption) + 24)
Dim q As Integer
q = Val(Me.Label9.Caption)
Hanb (q)
End Sub

Private Sub ListBox1_Click()
Me.Text1.Text = Left(Me.ListBox1.Text, 7)
Me.Text2.Text = Getnazi("select madanazi from mada where madasifr='" & Trim(Me.Text1.Text) & "'")
Me.Text3.Text = Right(RTrim(Me.ListBox1.Text), 4)
End Sub

Private Sub mii_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
 Case vbKey0 To vbKey9
      
       mizaa_Click (Chr(KeyCode))
       Me.mii.Visible = False
       
Case Else
 MsgBox ("Vnesti moraš številko!!")
    End Select
End Sub

Private Sub min_Click()
Me.Text3.Text = Me.Text3.Text - 1
myConection.Execute ("update trenutna set kol=" & Me.Text3.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=kol*(1-(pop/100))*cena where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssve (Me.MSHFlexGrid1.Row)
End Sub

Private Sub mizaa_Click(Index As Integer)
stm1 = Index
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
If mizaa(stm1).BackColor = 14215660 Then
shranimi (stm1 + stevmi)


Else
odprimi (stm1 + stevmi)
Dim sSQL As String
    
    'default
    
    
  '  sSQL = "DELETE * FROM mize WHERE stmize=" & Index
  '  myConection.Execute sSQL
  '  mizaa(Index).BackColor = 14215660
'MSHFlexGrid1.SetFocus
fora = 9


End If

End Sub





Private Sub MSHFlexGrid1_Click()
If Me.MSHFlexGrid1.FixedCols = 1 Then

Else
Me.Text1.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)
Me.Text2.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1)
Me.Text3.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
Me.pop.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)
End If
End Sub



Private Sub MSHFlexGrid1_SelChange()
If Me.MSHFlexGrid1.FixedCols = 1 Then

Else
Me.Text1.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)
Me.Text2.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1)
Me.Text3.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
Me.pop.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)
End If
End Sub



Private Sub nas10_Click()
Me.Label8.Caption = "0"
Hanb (10)
Me.Label9.Caption = "10"

End Sub

Private Sub nas11_Click()
Me.Label8.Caption = "0"
Hanb (11)
Me.Label9.Caption = "11"

End Sub

Private Sub nas12_Click()
Me.Label8.Caption = "0"
Hanb (12)
Me.Label9.Caption = "12"

End Sub

Private Sub nas13_Click()
Me.Label8.Caption = "0"
Hanb (13)
Me.Label9.Caption = "13"

End Sub

Private Sub nas14_Click()
Me.Label8.Caption = "0"
Hanb (14)
Me.Label9.Caption = "14"

End Sub

Private Sub nas15_Click()
Me.Label8.Caption = "0"
Hanb (15)
Me.Label9.Caption = "15"

End Sub

Private Sub nas16_Click()
Me.Label8.Caption = "0"
Hanb (16)
Me.Label9.Caption = "16"

End Sub

Private Sub nas17_Click()
Me.Label8.Caption = "0"
Hanb (17)
Me.Label9.Caption = "17"
End Sub

Private Sub nas18_Click()
Me.Label8.Caption = "0"
Hanb (18)
Me.Label9.Caption = "18"
End Sub

Private Sub nas19_Click()
Me.Label8.Caption = "0"
Hanb (19)
Me.Label9.Caption = "19"
End Sub

Private Sub nas2_Click()
Me.Label8.Caption = "0"
Hanb (2)
Me.Label9.Caption = "2"
End Sub

Private Sub nas20_Click()
Me.Label8.Caption = "0"
Hanb (20)
Me.Label9.Caption = "20"
End Sub

Private Sub nas3_Click()
Me.Label8.Caption = "0"
Hanb (3)
Me.Label9.Caption = "3"
End Sub

Private Sub nas4_Click()
Me.Label8.Caption = "0"
Hanb (4)
Me.Label9.Caption = "4"
End Sub

Private Sub nas5_Click()
Me.Label8.Caption = "0"
Hanb (5)
Me.Label9.Caption = "5"
End Sub

Private Sub nas6_Click()
Me.Label8.Caption = "0"
Hanb (6)
Me.Label9.Caption = "6"
End Sub

Private Sub nas7_Click()
Me.Label8.Caption = "0"
Hanb (7)
Me.Label9.Caption = "7"
End Sub

Private Sub nas8_Click()
Me.Label8.Caption = "0"
Hanb (8)
Me.Label9.Caption = "8"
End Sub

Private Sub nas9_Click()
Me.Label8.Caption = "0"
Hanb (9)
Me.Label9.Caption = "9"
End Sub

Private Sub narooc_Click()
Dim parr As String
If Me.narooc.BackColor = &H8080FF Then
Me.Frame2.Visible = True
Dim i
Me.ListBox1.MultiSelect = fmMultiSelectSingle
If rs.State = 1 Then rs.Close
rs.Open "select tip_dok,id_dok,datum from nabasif where tip_dok='NK' and isnull(poknj) group by tip_dok,id_dok,datum order by id_dok", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
parr = Getnazi("select dod0 from glavna where tip_dok='NK' and id_dok='" & rs.Fields("id_dok") & "'")
ListBox1.AddItem presled(rs.Fields("tip_dok") & rs.Fields("id_dok"), 8) & "  " & presled(rs.Fields("datum"), 10) & "  " & parr
rs.MoveNext
Loop

Me.ListBox1.SetFocus
Me.tvorbara.Visible = True
Me.zaprr.Visible = True
End If
End If
End Sub

Private Sub nas1_Click(Index As Integer)

Me.Label8.Caption = "0"
Hanb (Index + 1)
Me.Label9.Caption = LTrim(str(Index + 1))


End Sub

Private Sub plu_Click()
Me.Text3.Text = Me.Text3.Text + 1
myConection.Execute ("update trenutna set kol=" & Me.Text3.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=kol*(1-(pop/100))*cena where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssve (Me.MSHFlexGrid1.Row)
End Sub
Public Sub osssv()
Dim rsa1 As New ADODB.Recordset

rsa1.Open "select sifra,naziv,format(kol,'fixed') as kol,format(cena,'fixed') as cena,format(znes,'fixed') as znesek,X,pop from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic

If rsa1.EOF Then
Me.plu.Visible = False
Me.min.Visible = False
Me.bri.Visible = False
Me.MSHFlexGrid1.Visible = False
Else
Me.MSHFlexGrid1.Visible = True

Me.plu.Visible = True
Me.min.Visible = True
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'") <> "D" Then

Me.bri.Visible = True
Else
Me.bri.Visible = False
End If
    
    Set Me.MSHFlexGrid1.DataSource = rsa1
Me.MSHFlexGrid1.Refresh
'Call FG_AutosizeCols(MSHFlexGrid1, Me, , , False)

MSHFlexGrid1.ColAlignment(2) = flexAlignRightCenter
      MSHFlexGrid1.ColAlignment(3) = flexAlignRightCenter
       MSHFlexGrid1.ColAlignment(4) = flexAlignRightCenter
         MSHFlexGrid1.Redraw = True ' dont forget to do this !
 LoadFlexGridColumnWidths MSHFlexGrid1, "paragonc"

'Me.MSHFlexGrid1.SetFocus
'Me.MSHFlexGrid1.TopRow = Me.MSHFlexGrid1.Rows - 1
Me.MSHFlexGrid1.Row = Me.MSHFlexGrid1.Rows - 1
'Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Rows - 2
MSHFlexGrid1_Click
Me.znees.Caption = Getnazi("select sum(format(znes,'fixed')) from trenutna where  stdok='" & Pblagajna & "'")


End If

End Sub
Private Sub osssve(kkje As Integer)
Dim rsa1 As New ADODB.Recordset

rsa1.Open "select sifra,naziv,kol,format(cena,'fixed') as cena,format(znes,'fixed') as znesek,X,pop from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic

If rsa1.EOF Then
Me.plu.Visible = False
Me.min.Visible = False
Me.bri.Visible = False
Me.MSHFlexGrid1.Visible = False
Else
Me.MSHFlexGrid1.Visible = True

Me.plu.Visible = True
Me.min.Visible = True
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'") <> "D" Then

Me.bri.Visible = True
Else
Me.bri.Visible = False
End If
    
    Set Me.MSHFlexGrid1.DataSource = rsa1
Me.MSHFlexGrid1.Refresh
'Call FG_AutosizeCols(MSHFlexGrid1, Me, , , False)

MSHFlexGrid1.ColAlignment(2) = flexAlignRightCenter
      MSHFlexGrid1.ColAlignment(3) = flexAlignRightCenter
       MSHFlexGrid1.ColAlignment(4) = flexAlignRightCenter
         MSHFlexGrid1.Redraw = True ' dont forget to do this !
 LoadFlexGridColumnWidths MSHFlexGrid1, "paragonc"

'Me.MSHFlexGrid1.SetFocus
'Me.MSHFlexGrid1.TopRow = Me.MSHFlexGrid1.Rows - 1
Me.MSHFlexGrid1.Row = kkje
'Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Rows - 2
MSHFlexGrid1_Click
Me.znees.Caption = Getnazi("select sum(format(znes,'fixed')) from trenutna where  stdok='" & Pblagajna & "'")


End If

End Sub
Private Sub pop_Click()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub pop_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub pop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'myConection.Execute ("update trenutna set cena=cena/(1-(pop/100)) where pop<>0 and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))

myConection.Execute ("update trenutna set pop=" & Me.pop.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
'myConection.Execute ("update trenutna set cena=cena*(1-(pop/100)) where x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=cena*(1-(pop/100))*kol where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssv
Me.Text3.Text = 1
Me.Text2.Text = ""
Me.Text1.Text = ""
Me.pop.Text = 0
Me.Text1.SetFocus

End If
End Sub

Private Sub pop_LostFocus()

xpopu = Me.pop.Text
Me.pop.Text = 0
End Sub

Private Sub pred_Click()

predal
End Sub

Private Sub pregled_Click()
pregledr.Show vbModal
End Sub

Private Sub prija_Click()
UPORABNIKFO.Show vbModal
End Sub

Private Sub reset_Click()
End
End Sub

Private Sub shran_Click()
SaveFlexGridColumnWidths MSHFlexGrid1, "paragonc"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")
Select Case KeyCode

Case vbKeyF3
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'") <> "D" Then

 LaVolpeButton45.SetFocus
 Sendkeys "{enter}"
End If
 Case vbKeyF7
 LaVolpeButton2522.SetFocus
  Sendkeys "{enter}"
 Case vbKeyF2
 If Me.kart.Value = True Then
 Me.kart.Value = False
 Else
 Me.kart.Value = True
 End If
  Case vbKeyF9
VRNIT.SetFocus
   Sendkeys "{enter}"
  Case vbKeyF10
  LaVolpeButton44.SetFocus
   Sendkeys "{enter}"
 Case vbKeyF8
 LaVo2.SetFocus
  Sendkeys "{enter}"
 Case vbKeyF6
fora = 1
 Me.mii.Visible = True
 Me.mii.Text = ""
 
 Me.mii.SetFocus
 
 
Case vbKeyF4
 zakljucc(0).SetFocus
 Sendkeys "{enter}"

' Case vbKeyA To vbKeyZ
Case vbKeyReturn
Dim iid As String
vrjenniz = ""
'idar = Chr(KeyCode)
'   DoSQL "mada", "madasifr", "madanazi", "madanaz1"
 Dim ax As String
 If Not IsNumeric(Text1.Text) Then
 Text1.Text = Replace(Text1.Text, "*", "%")
 iskalni = Text1.Text
       '& Chr(KeyCode)
       pritisk = Text1.Text
       '& Chr(KeyCode)
      ' DoSQL = ""
     

       ax = DoSQLbe("mada", "madasifr", "madanazi", "madanaz1")
      Me.Text1.Text = ax
      Else
      ax = Text1.Text
      Exit Sub
       End If
     'MsgBox (ax)
      
      If Getnazi("select madanazi from mada where madasifr='" & Trim(ax) & "'") <> "" Then
Me.Text2.Text = Getnazi("select madanazi from mada where madasifr='" & Trim(ax) & "'")
Me.Text3.Text = 1
Me.pop.Text = 0
Dim rsa1 As New ADODB.Recordset
rsa1.Open "select sifra,naziv,kol,znes,x,cena,pop,stdok from trenutna where  stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
rsa1.AddNew
rsa1.Fields("sifra") = Me.Text1.Text
rsa1.Fields("naziv") = Me.Text2.Text
rsa1.Fields("kol") = Me.Text3.Text
rsa1.Fields("pop") = Me.pop.Text
rsa1.Fields("stdok") = Pblagajna
rsa1.Fields("cena") = Getnazi("select madampcd from mada where madasifr='" & Trim(ax) & "'")
'rsa1.Fields("mpc") = Getcena(Me.Text1.text)

rsa1.Fields("znes") = Getnazi("select madampcd from mada where madasifr='" & Trim(ax) & "'") * Me.Text3.Text
rsa1.Fields("x") = Getnumb("select max(x) as x  from trenutna where  stdok='" & Pblagajna & "'") + 1
rsa1.Update
osssv
Me.Text3.Text = 1
Me.Text3.SetFocus

Sendkeys "+{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}"
'Sendkeys "{END}"
'Sendkeys "+{HOME}"
End If
Case Else
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not KeyAscii >= 48 And Not KeyAscii <= 57 Then
'Exit Sub
End If
If KeyAscii = 27 Then
Me.Text1.Text = ""
End If
If KeyAscii = 13 Then
If Len(Me.Text1.Text) > 10 Then
If Right(Me.Text1.Text, 1) = "J" Then
Me.Text1.Text = Left(Me.Text1.Text, Len(Me.Text1.Text) - 1)
End If
End If
Dim siii As String
siii = Getnazi("select madasifr from mada where madasifr='" & Trim(Me.Text1.Text) & "'")
If siii = "" Then
siii = Getnazi("select madasifr from mada where madaean='" & Trim(Me.Text1.Text) & "'")
End If
If siii <> "" Then
Me.Text1.Text = siii
Me.Text2.Text = Getnazi("select madanazi from mada where madasifr='" & Trim(Me.Text1.Text) & "'")
Me.Text3.Text = 1
Dim cennnn As Double
cennnn = Getnazi("select madampcd from mada where madasifr='" & Trim(Me.Text1.Text) & "'")
Dim rsa1 As New ADODB.Recordset
rsa1.Open "select sifra,naziv,kol,znes,x,cena,pop,stdok from trenutna where  stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
rsa1.AddNew
rsa1.Fields("sifra") = Me.Text1.Text
rsa1.Fields("naziv") = Me.Text2.Text
rsa1.Fields("kol") = Me.Text3.Text
rsa1.Fields("pop") = Me.pop.Text
rsa1.Fields("stdok") = Pblagajna
rsa1.Fields("cena") = cennnn '
'rsa1.Fields("mpc") = Getcena(Me.Text1.text)

rsa1.Fields("znes") = cennnn * Me.Text3.Text
rsa1.Fields("x") = Getnumb("select max(x) as x  from trenutna where  stdok='" & Pblagajna & "'") + 1
rsa1.Update
osssv
Me.Text3.Text = 1
Me.Text3.SetFocus
Sendkeys "+{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}"
End If
End If
End Sub

Private Sub Text1_LostFocus()
If Getnazi("select sifra from dodatni where sifra='" & Me.Text1.Text & "'") <> "" Then
DOD_AR = "blg"
Me.Text2.Text = Trim(Text2.Text) & " " & Trim(Dodat_a(Text1.Text))
'MsgBox ("update trenutna set naziv=" & Replace(Me.Text2.Text, "'", "") & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set naziv='" & Replace(Me.Text2.Text, "'", "") & "' where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))

End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Me.Text3.Text = "" Then
Me.Text3.Text = 0

End If
myConection.Execute ("update trenutna set kol=" & Replace(Me.Text3.Text, ",", ".") & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=kol*(1-(pop/100))*cena where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssv
Me.Text3.Text = 1
Me.Text2.Text = ""
Me.Text1.Text = ""
Me.pop.Text = 0
Me.Text1.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()


If Me.Text3.Text = "" Then
Me.Text3.Text = 0

End If
'myConection.Execute ("update trenutna set kol=" & Me.Text3.text & " where x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
'myConection.Execute ("update trenutna set znes=kol*cena where x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
'osssv
End Sub

Private Sub Timer1_Timer()
If prijavljen = "" Then
UPORABNIKFO.Show vbModal
Else
Me.Label3.Caption = Trim(prijavljen) & " " & Getnazi("select username1 from users where up='" & Trim(prijavljen) & "'")
End If
LblDateTime.Caption = Time() & " " & Format(Date, "DDDD")
If refr = 1 Then
Dim kerabl As Integer
kerabl = Val(GetSetting("bll", "sifrablg", "odmize", "1"))

For miz = kerabl To kerabl + 9
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
refr = 0
End If

End Sub


Private Function Hanb(intBtn As Integer)
    trenu = intBtn
    Flistvel veli, "select dim from swit WHERE [command]<>1 AND [Switchboar]=" & nas1(intBtn - 1).Tag & " group by dim order by dim"
    
    If rs.State = 1 Then rs.Close
   If sqlb = "" Then
   rs.Open "select * from swit WHERE [ItemNumber] > " & Val(Me.Label8.Caption) + 1 & " and [command]<>1 AND [Switchboar]=" & nas1(intBtn - 1).Tag & " order by [ItemNumber]"
   Else
   rs.Open sqlb
   'sqlb = ""
   End If
      If rs.EOF Then
      Exit Function
      End If
      rs.MoveFirst
      Dim aad As Integer
      aad = 0
      If Not rs.EOF Then
    Do While Not aad = 35
      aad = aad + 1
      Me("LaVolpeButton" & aad).Tag = ""
      Me("LaVolpeButton" & aad).Visible = True
       Me("LaVolpeButton" & aad).BackColor = &HF8F3F1
      Loop
      aad = 0
      rs.MoveFirst
      While Not rs.EOF
       aad = aad + 1
       If aad <= 35 Then
       If Not IsNull(rs![ITEMTEXT]) Then
         Me("LaVolpeButton" & aad).Caption = rs![ITEMTEXT]
         If Val(Getnazi("select kontrola from mada where madasifr='" & rs![ARGUMENT] & "'")) <> 0 Then
          Me("LaVolpeButton" & aad).BackColor = Val(Getnazi("select kontrola from mada where madasifr='" & rs![ARGUMENT] & "'"))
         End If
         Me("LaVolpeButton" & aad).Tag = rs![ARGUMENT]
         End If
         If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='BOLDDA'") = "D" Then
             Me("LaVolpeButton" & aad).Font.Bold = True
         Else
             Me("LaVolpeButton" & aad).Font.Bold = False
         End If
      
       End If
            rs.MoveNext
        Wend
        aad = 0
      Do While Not aad = 35
      aad = aad + 1
      If Me("LaVolpeButton" & aad).Tag = "" Then
      Me("LaVolpeButton" & aad).Visible = False
      End If
      Loop
      Else
         End If
        
    ' If no item matches, report the error and exit the function.
    
    
End Function

Private Function Hanbt(intBt As Integer)
   Me.pop.Text = 0
    Me.Text1.SetFocus
  
    Me.Text1.Text = Me("LaVolpeButton" & intBt).Tag
Sendkeys "{enter}"

 
End Function

Private Function Hanbtx(intBt As Integer)

End Function


Public Function hh()
'MSHFlexGrid1.SetFocus
'SendKeys "{BS}"
End Function
Private Function mi()
Dim strsq As String
strsq = "select stmize from mize group by stmize order by stmize"
If rs.State = 1 Then rs.Close
rs.Open strsq, myConection
Dim ss As String
Dim kerabl As Integer
kerabl = Val(GetSetting("bll", "sifrablg", "odmize", "1"))

ss = ""
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    If rs.Fields("stmize") >= kerabl + 1 And rs.Fields("stmize") <= kerabl + 10 Then
 ss = ss & "," & rs.Fields("stmize")
       Me.mizaa(rs.Fields("stmize") - kerabl).BackColor = 5609
    End If
    rs.MoveNext
    Loop
'MsgBox (ss)
End If
End Function
Public Function shranimi(stm As Integer)
Dim i, stot, fa
Dim aaa As String
aaa = Left(Me.Label12.Caption, 8)

'MsgBox (aaa)
 myConection.Execute ("insert into mize select sifra,naziv,kol,x as mpcd,znes as znesek," & stm & " as stmize,pop as ddva from trenutna where stdok='" & Pblagajna & "'")
 myConection.Execute ("delete from trenutna where stdok='" & Pblagajna & "'")
Indx = 1
zap = 1
osssv
idstran = 0
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
For miz = 1 To 10
mizaa(miz).Caption = miz + stevmi
mizaa(miz).BackColor = 14215660

Next
mi
skumi = 0
  
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE F6 "
Me.Label12.Caption = ""
If stalnaprij = 1 Then
Else
prijavljen = ""
End If
Me.dav.Visible = False
Me.dav.Text = ""
Me.davlb.Visible = False
Me.imes.Visible = False
Me.imes.Text = ""
Me.nassl.Visible = False
Me.nassl.Text = ""
Me.pop.Text = 0
End Function
Public Function odprimi(stm As Integer)
If Me.Label11.Caption <> "" Then
MsgBox "Miza " & Me.Label11.Caption & " je ze odprta! Najprej zapri njo nato lahko odpreš novo!!!!"
Else
Dim i, stot, fa
stm1 = stm
Me.Label11.Caption = LTrim(str(stm))
Me.Label12.Caption = Getnazi("select ura from mize where stmize=" & stm)
Me.LaVo1.Caption = "SHRANI MIZO " & LTrim(str(stm))
Dim aaa As String
aaa = Left(Time(), 8)
'MsgBox (aaa)
Dim aaaa As String
aaaa = "insert into trenutna select sifra,naziv,kol,stmize as doza,ddva as pop,mpcd as x,znesek as znes,'" & Pblagajna & "' as stdok  from mize  where stmize=" & stm & " order by mpcd"
'MsgBox (aaaa)
myConection.Execute (aaaa)
 If rs.State = 1 Then rs.Close
rs.Open "select * from trenutna where doza=" & stm, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
'rs.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & RTrim(LTrim(rs.Fields("sifra"))) & "'")
'If Not rs.Fields("znes") = 0 Then
rs.Fields("cena") = Getnumb("select madampcd from mada where madasifr='" & rs.Fields("sifra") & "'")
'End If
rs.Update
rs.MoveNext

Loop
End If
osssv
myConection.Execute ("delete from mize where stmize=" & stm)
Dim stevmi As Integer
stevmi = GetSetting("bll", "sifrablg", "odmize", "1")
For miz = 1 To 10
mizaa(miz).Caption = miz + stevmi
mizaa(miz).BackColor = 14215660

Next
mi

End If
End Function
Private Sub printracluk()
'luka
 Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
    'frmPrinterSetUp.Show vbModal
    'i f QuitCommand Then
     '   Set cPrint = Nothing
     '   Exit Sub
    'End If

    
SendToPrinter:
    picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 8
    cPrint.FontName = "Courier new"
    cPrint.CurrentY = 0
    ' cPrint.pPrint Chr(27) & Chr(116) & Chr(18), 0.1, False
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
    cPrint.pPrint
    cPrint.pPrint "Prodajalec: " & Me.Label3.Caption
    If Me.imes.Text <> "" Then
    
    cPrint.pPrint
    cPrint.pPrint "Stranka:"
    cPrint.pPrint Left(Me.imes.Text, 40)
cPrint.pPrint Mid(Me.imes.Text, 40, 40)
cPrint.pPrint Left(Me.nassl.Text, 40)
cPrint.pPrint Mid(Me.nassl.Text, 40, 40)
cPrint.pPrint "ID.ST.: SI" & Me.dav.Text

    
    End If
    
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0, True
    cPrint.pPrint txtInvoiceNo.Text, 0.9, True
    cPrint.pPrint " z dne " & Format(Date, "dd/mm/yyyy")
    '& " " & Format(Time(), "hh:mm")
    
    cPrint.pPrint "", 0, False
    cPrint.pPrint "================================", 0, False
    cPrint.pPrint "Naziv                kol  znesek", 0, False
    cPrint.pPrint "================================", 0, False
    Dim i, ass
    Dim popu As Double
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim znep As Double
    Dim ddv2 As Double
    ddv1 = 0
    ddv2 = 0
    popu = 0
    znep = 0
    sku = 0
    For i = 1 To MSHFlexGrid1.Rows - 1
    
   If Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'") = "20" Then
   ddv1 = ddv1 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    If Replace(Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'"), ",", ".") = "8.5" Then
   ddv2 = ddv2 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    stri = Format(MSHFlexGrid1.TextMatrix(i, 2), "standard")
    stri1 = Format(MSHFlexGrid1.TextMatrix(i, 4), "standard")
    If MSHFlexGrid1.TextMatrix(i, 4) <> "" Then
    sku = sku + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    End If
     If stri1 <> "" Then
     If (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)) > 0 Then
     znep = znep + (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2))
     End If
    'MsgBox (Val(Getnazi("select madampcd from mada where madasifr=" & Val(MSHFlexGrid1.TextMatrix(i, 1)))) - (Val(MSHFlexGrid1.TextMatrix(i, 5)) / Val(MSHFlexGrid1.TextMatrix(i, 4))))
    If MSHFlexGrid1.TextMatrix(i, 6) <> 0 Then
    popu = popu + (Getnumb("select madampcd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'")) - FormatNumber(MSHFlexGrid1.TextMatrix(i, 3), 2)
    End If
    End If
    'popu = 0
    'popu = FormatNumber(popu, 2)
    cPrint.pPrint "", 0, False
    cPrint.pPrint Left(MSHFlexGrid1.TextMatrix(i, 1), 17), 0, True
    cPrint.pRightJust stri, tis_a, True
    
    'cPrint.pRightJust Format(MSHFlexGrid1.TextMatrix(i, 6), "standard"), tis_b, True
    cPrint.pRightJust stri1, tis_c, True
    Next
   
    cPrint.pPrint ""
    'cPrint.pPrint ""
    cPrint.pPrint "================================", 0, False
    'cPrint.pPrint ""
    If popu <> 0 Then
    cPrint.pPrint "Popust vracunan v ceni", 0.1, True
    cPrint.pRightJust Format(popu, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "--------------------------------", 0, False
    End If
    cPrint.pPrint "ZA PLACILO EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), 2.5, True
    cPrint.pPrint "", 0.1, False
    If znep > 0 Then
    'cPrint.pPrint "Znesek popusta ", 0.1, True
    'cPrint.pRightJust Format(znep, "standard"), 2.5, True
    'cPrint.pPrint "", 0.1, False
    End If
'    cPrint.pPrint "SKUPAJ SIT", 0.1, True
'    cPrint.pRightJust Format(sku * 239.64, "standard"), 4, True
    zavrnit = sku
    
      cPrint.pPrint
    
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "--------------------------------", 0, False
    cPrint.pPrint "Osnova  DDV Znesek DDV  Vrednost", 0, False
    cPrint.pPrint "--------------------------------", 0, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.7, True
    cPrint.pRightJust "20 %", tis_e, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 2.5 * tiskdol, True
    cPrint.pRightJust Format(ddv1, "standard"), 3.2 * tiskdol, True
    'cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.8, True
    'cPrint.pRightJust " 20 %", 2, True
    'cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3, True
    'cPrint.pRightJust Format(ddv1, "standard"), 4, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.7, True
    cPrint.pRightJust "8.5 %", 1.3, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 2.5, True
    cPrint.pRightJust Format(ddv2, "standard"), 3.2, True
    
   ' cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.8, True
   ' cPrint.pRightJust "8.5 %", 2, True
   ' cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3, True
   ' cPrint.pRightJust Format(ddv2, "standard"), 4, True
    End If
    End If
    Dim pl As String
    
    If Me.kart.Value = True Then
    pl = "KARTICA"
    Else
    pl = "GOTOVINA"
    End If
     If Me.inter.Value = True Then
    pl = "INTERNA     Podpis ______________________"
    Else
    pl = "GOTOVINA"
    End If
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint " Placilo: " & plax, 0.1, False
    If Getnazi("select konec1 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec1 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec2 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec3 from oblikar"), 0.1, False
    End If
 '   cPrint.pPrint Getnazi("select konec4 from oblikar"), 0.1, False
  '  cPrint.pPrint Getnazi("select konec5 from oblikar"), 0.1, False
    cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
        cPrint.pPrint "", 0.1, False
   cPrint.pPrint " ", 0.1, False
 cPrint.pPrint ""
 '  cPrint.pPrint " ", 0.1, False
 '  cPrint.pPrint " ", 0.1, False
 'cPrint.pPrint Chr(27) & "i"
'  cPrint.pPrint Chr(27) & Chr(100) & Chr(48)
'cPrint.pPrint Chr(27) & Chr(105)
 'cPrint.pPrint Chr(7)
'

If FileExist("c:\be.txt") Then
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
End If
    'cPrint.pPrint DEKODIRAJ("rezi")
 'cPrint.pPrint " ", 0.1, False
 
' cPrint.pPrint Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
' cPrint.pPrint " ", 0.1, False
'  cPrint.pPrint " ", 0.1, False
    'odrez
    'cPrint.pPrint Chr(27) & "i", 0.1, False
    predal
   ' cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
     ' If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
 
 
'Print #1, "======================================="
'Print #1, "SKUPAJ SIT  ", Format(asd, "0.00")
'Print #1,
'Print #1, "SKUPAJ EUR  ", Format(asd / DLookup("[eur]", "eur"), "0.00")

'Print #1, "---------------------------------------"
'Print #1, "Osnova DDV-a  DDV  Znesek DDV  Vrednost"
'If ddv > 0 Then
'Print #1, "  " & Format(ddv, "0.00") & "   20.00 %  " & Format(zneddv - ddv, "0.00") & "  " & Format(zneddv, "0.00")
'End If
'If ddv1 > 0 Then
'Print #1, "  " & Format(ddv1, "0.00") & "    8.50 %  " & Format(zneddv1 - ddv1, "0.00") & "  " & Format(zneddv1, "0.00")



'End If
'Print #1, "---------------------------------------"


'Print #1, "---------------------------------------"'
'End If

'Call Shell("print /d:LPT1 c:\be.txt", 6)
   
End Sub
Private Sub printrac()
'original
 Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
    'frmPrinterSetUp.Show vbModal
    'i f QuitCommand Then
     '   Set cPrint = Nothing
     '   Exit Sub
    'End If

    
SendToPrinter:
    picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 8
    cPrint.FontName = "Courier new"
    cPrint.CurrentY = 0
    ' cPrint.pPrint Chr(27) & Chr(116) & Chr(18), 0.1, False
     If Getnazi("select glava1 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    End If
    If Getnazi("select glava2 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    End If
    If Getnazi("select glava3 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    End If
    If Getnazi("select glava4 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    End If
    If Getnazi("select glava5 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    End If
    
    cPrint.pPrint
    cPrint.pPrint "Prodajalec: " & Me.Label3.Caption
    If Me.imes.Text <> "" Then
    
    cPrint.pPrint
    cPrint.pPrint "Stranka:"
    cPrint.pPrint Left(Me.imes.Text, 40)
cPrint.pPrint Mid(Me.imes.Text, 40, 40)
cPrint.pPrint Left(Me.nassl.Text, 40)
cPrint.pPrint Mid(Me.nassl.Text, 40, 40)
cPrint.pPrint "ID.ST.: SI" & Me.dav.Text

    
    End If
    
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0, True
    cPrint.pPrint txtInvoiceNo.Text, 0.9, True
    cPrint.pPrint " z dne " & Format(Date, "dd/mm/yyyy")
    
    If Getdoba(LTrim(txtInvoiceNo.Text)) <> "" Then
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Dobavnice: ", 0.1, False
    cPrint.pPrint Getdoba(LTrim(txtInvoiceNo.Text))
    '& " " & Format(Time(), "hh:mm")
    End If
    If stevnaro <> "" Then
    cPrint.pPrint "", 0, False
    cPrint.pPrint "Naroèilo : " & stevnaro, 0, False
    cPrint.pPrint Getnazi("select dod1 from glavna where tip_dok+id_dok='" & stevnaro & "'"), 0, False
    cPrint.pPrint Getnazi("select dod2 from glavna where tip_dok+id_dok='" & stevnaro & "'"), 0, False
    cPrint.pPrint Getnazi("select dod3 from glavna where tip_dok+id_dok='" & stevnaro & "'"), 0, False
    cPrint.pPrint Getnazi("select dod4 from glavna where tip_dok+id_dok='" & stevnaro & "'"), 0, False
    
    cPrint.pPrint Getnazi("select dod5 from glavna where tip_dok+id_dok='" & stevnaro & "'"), 0, False
    End If
    cPrint.pPrint "", 0, False
    cPrint.pPrint "========================================", 0, False
    cPrint.pPrint "Naziv                     kol pop znesek", 0, False
    cPrint.pPrint "========================================", 0, False
    Dim i, ass
    Dim popu As Double
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim znep As Double
    Dim ddv2 As Double
    ddv1 = 0
    ddv2 = 0
    popu = 0
    znep = 0
    sku = 0
    For i = 1 To MSHFlexGrid1.Rows - 1
    
   If Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'") = "20" Then
   ddv1 = ddv1 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    If Replace(Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'"), ",", ".") = "8.5" Then
   ddv2 = ddv2 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    stri = Format(MSHFlexGrid1.TextMatrix(i, 2), "standard")
    stri1 = Format(MSHFlexGrid1.TextMatrix(i, 4), "standard")
    If MSHFlexGrid1.TextMatrix(i, 4) <> "" Then
    sku = sku + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    End If
     If stri1 <> "" Then
     If (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)) > 0 Then
     znep = znep + (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2))
     End If
    'MsgBox (Val(Getnazi("select madampcd from mada where madasifr=" & Val(MSHFlexGrid1.TextMatrix(i, 1)))) - (Val(MSHFlexGrid1.TextMatrix(i, 5)) / Val(MSHFlexGrid1.TextMatrix(i, 4))))
    If MSHFlexGrid1.TextMatrix(i, 6) <> 0 Then
    popu = popu + (Getnumb("select madampcd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'") * stri) - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    End If
    End If
    'popu = 0
    popu = FormatNumber(popu, 2)
    cPrint.pPrint "", 0, False
    cPrint.pPrint Left(MSHFlexGrid1.TextMatrix(i, 1), 17), 0, True
    cPrint.pRightJust FormatNumber(stri, 0), tis_a, True
    
    cPrint.pRightJust FormatNumber(MSHFlexGrid1.TextMatrix(i, 6), 0), tis_b, True
    cPrint.pRightJust stri1, tis_c, True
    Next
   
    cPrint.pPrint ""
    'cPrint.pPrint ""
    cPrint.pPrint "========================================", 0, False
    'cPrint.pPrint ""
    If popu <> 0 Then
    cPrint.pPrint "Popust vracunan v ceni", 0.1, True
    cPrint.pRightJust Format(popu, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "----------------------------------------", 0, False
    End If
    cPrint.pPrint "ZA PLACILO EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "                              ==========", 0, False
    cPrint.pPrint "", 0, False
    If znep > 0 Then
    'cPrint.pPrint "Znesek popusta ", 0.1, True
    'cPrint.pRightJust Format(znep, "standard"), 2.5, True
    'cPrint.pPrint "", 0.1, False
    End If
'    cPrint.pPrint "SKUPAJ SIT", 0.1, True
'    cPrint.pRightJust Format(sku * 239.64, "standard"), 4, True
    zavrnit = sku
    
      cPrint.pPrint
    
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "----------------------------------------", 0, False
    cPrint.pPrint "Osnova  DDV        Znesek DDV   Vrednost", 0, False
    cPrint.pPrint "----------------------------------------", 0, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), tis_d, True
    cPrint.pRightJust "20 %", tis_e, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), tis_a, True
    cPrint.pRightJust Format(ddv1, "standard"), tis_c, True
    'cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.8, True
    'cPrint.pRightJust " 20 %", 2, True
    'cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3, True
    'cPrint.pRightJust Format(ddv1, "standard"), 4, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), tis_d, True
    cPrint.pRightJust "8.5 %", tis_e, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), tis_b, True
    cPrint.pRightJust Format(ddv2, "standard"), tis_c, True
    
   ' cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.8, True
   ' cPrint.pRightJust "8.5 %", 2, True
   ' cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3, True
   ' cPrint.pRightJust Format(ddv2, "standard"), 4, True
    End If
    End If
    Dim pl As String
    cPrint.pPrint
    If Me.kart.Value = True Then
    pl = "KARTICA"
    Else
    pl = "GOTOVINA"
    End If
     If Me.inter.Value = True Then
    pl = "INTERNA => Podpis ___________"
    Else
    pl = "GOTOVINA"
    End If
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint " Placilo: ", 0.1, False
    cPrint.pPrint Getnacin1("PA" & LTrim(txtInvoiceNo.Text), sku)
    If Getnazi("select konec1 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec1 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec2 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec3 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec4 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec5 from oblikar"), 0.1, False
    End If
 '   cPrint.pPrint Getnazi("select konec4 from oblikar"), 0.1, False
  '  cPrint.pPrint Getnazi("select konec5 from oblikar"), 0.1, False
    cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
        cPrint.pPrint "", 0.1, False
   cPrint.pPrint " ", 0.1, False
 cPrint.pPrint ""
 '  cPrint.pPrint " ", 0.1, False
 '  cPrint.pPrint " ", 0.1, False
 'cPrint.pPrint Chr(27) & "i"
'  cPrint.pPrint Chr(27) & Chr(100) & Chr(48)
'cPrint.pPrint Chr(27) & Chr(105)
 'cPrint.pPrint Chr(7)
'

If FileExist("c:\be.txt") Then
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
End If
    'cPrint.pPrint DEKODIRAJ("rezi")
 'cPrint.pPrint " ", 0.1, False
 
' cPrint.pPrint Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
' cPrint.pPrint " ", 0.1, False
'  cPrint.pPrint " ", 0.1, False
    'odrez
    'cPrint.pPrint Chr(27) & "i", 0.1, False
    predal
   ' cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
     ' If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
 
 
'Print #1, "======================================="
'Print #1, "SKUPAJ SIT  ", Format(asd, "0.00")
'Print #1,
'Print #1, "SKUPAJ EUR  ", Format(asd / DLookup("[eur]", "eur"), "0.00")

'Print #1, "---------------------------------------"
'Print #1, "Osnova DDV-a  DDV  Znesek DDV  Vrednost"
'If ddv > 0 Then
'Print #1, "  " & Format(ddv, "0.00") & "   20.00 %  " & Format(zneddv - ddv, "0.00") & "  " & Format(zneddv, "0.00")
'End If
'If ddv1 > 0 Then
'Print #1, "  " & Format(ddv1, "0.00") & "    8.50 %  " & Format(zneddv1 - ddv1, "0.00") & "  " & Format(zneddv1, "0.00")



'End If
'Print #1, "---------------------------------------"


'Print #1, "---------------------------------------"'
'End If

'Call Shell("print /d:LPT1 c:\be.txt", 6)
   
End Sub



Private Sub Timer2_Timer()
If Me.LaVolpeButton42.Visible = True Then
If Getnazi("select max(id_dok) from nabasif where tip_dok='NK' and isnull(poknj) and x=0") <> "" Then
Me.LaVolpeButton42.BackColor = &HFF&
'MsgBox ("")
Else
'MsgBox ("")
Me.LaVolpeButton42.BackColor = 14215660
End If


End If
If Me.LaVolpeButton37.Visible = True Then
If AllFiles(App.path & "\naro") <> "" Then
narocila.Show vbModal
End If
End If
If Getnazi("select max(id_dok) from nabasif where tip_dok='NK' and isnull(poknj)") <> "" Then
Me.narooc.BackColor = &H8080FF
'MsgBox ("")
Else
'MsgBox ("")
Me.narooc.BackColor = 14215660
End If


If Mid(LTrim(Me.LblDateTime.Caption), 5, 1) = "0" Then
If Me.LaVolpeButton37.Visible = True Then
LaVolpeButton37_Click
End If
End If
txtInvoiceNo.Text = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)

Me.Timer2.Interval = 30000
End Sub

Private Sub tvorbara_Click()
'MsgBox (Me.ListBox1.Text)
Dim aaaa, stnr As String
stnr = Left(Me.ListBox1.Text, 8)
If stnr <> "" Then
aaaa = "insert into trenutna select (tip_dok+id_dok) as kopija,sifra,naziv,cena,kol,pop, x, znes,'" & Pblagajna & "' as stdok  from nabasif  where tip_dok+id_dok='" & stnr & "' order by pozicija"
'MsgBox (aaaa)
stevnaro = stnr
myConection.Execute (aaaa)
Me.tvorbara.Visible = False
Me.zaprr.Visible = False
Me.Frame2.Visible = False
Me.ListBox1.MultiSelect = fmMultiSelectMulti
osssv
End If
End Sub

Private Sub veli_Click()
Me.Label10.Caption = Me.veli.Text
If veli.Text = "VSE" Then
sqlb = ""
Else
sqlb = "select * from swit WHERE [ItemNumber] > " & Val(Me.Label8.Caption) + 1 & " and [command]<>1 AND [Switchboar]=" & Me("nas" & trenu).Tag & " and dim='" & Me.veli.Text & "' order by [ItemNumber]"
End If
Hanb (trenu)
End Sub



Private Sub vena_Click()
nacpla.odprnac "PA" & Trim(Me.txtInvoiceNo.Text), Me.znees.Caption
End Sub

Private Sub VRNIT_Click()
zavrnit = Me.znees.Caption
xzago = 1
Form5.Show vbModal

End Sub



Private Sub printrac2()
 Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
    'frmPrinterSetUp.Show vbModal
    'If QuitCommand Then
    '    Set cPrint = Nothing
    '    Exit Sub
    'End If

    
SendToPrinter:
    picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    
    cPrint.pPrint "Zaposlen: " & Me.Label3.Caption
    If idstran <> 0 Then
    cPrint.pPrint "Stranka:"
    cPrint.pPrint Getnazi("select naziv from partner where sifra=" & idstran)
cPrint.pPrint Getnazi("select ulica from partner where sifra=" & idstran)
cPrint.pPrint Getnazi("select posta from partner where sifra=" & idstran)
cPrint.pPrint Getnazi("select mesto from partner where sifra=" & idstran)
cPrint.pPrint Getnazi("select davcna from partner where sifra=" & idstran)

    
    End If
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0.1, True
    cPrint.pPrint Me.txtInvoiceNo.Text, 1, True
    cPrint.pPrint "z dne " & Format(Date, "dd/mm/yyyy") & " "
    '& Format(Time(), "hh:mm"), 1.6, True
    
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Naziv                   kol      znesek ", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    Dim i, ass
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim ddv2 As Double
    ddv1 = 0
    ddv2 = 0
    sku = 0
Dim vss As Integer
Dim v As Integer
v = 15
vss = Me.MSHFlexGrid1.Row

    For i = 1 To MSHFlexGrid1.Row
   If Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 1)) & "'") = "20" Then
   ddv1 = ddv1 + Val(MSHFlexGrid1.TextMatrix(i, 5)) / vss
   End If
    If Replace(Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 1)) & "'"), ",", ".") = "8.5" Then
   ddv2 = ddv2 + Val(MSHFlexGrid1.TextMatrix(i, 5)) / vss
   End If
    stri = Format(MSHFlexGrid1.TextMatrix(i, 4), "standard")
    stri1 = Format(v / vss, "standard")
    sku = 15
    
cPrint.pPrint "", 0.1, False
    cPrint.pPrint MSHFlexGrid1.TextMatrix(i, 2), 0.1, True
    cPrint.pRightJust stri, 3, True
    cPrint.pRightJust stri1, 4, True
    Next
    cPrint.pPrint ""
    'cPrint.pPrint ""
    cPrint.pPrint "=======================================", 0.1, False
    'cPrint.pPrint ""
    cPrint.pPrint "SKUPAJ EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    
    'cPrint.pPrint "SKUPAJ SIT", 0.1, True
    'cPrint.pRightJust Format(sku * 239.64, "standard"), 4, True
    zavrnit = sku
      cPrint.pPrint
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 1.2, True
    cPrint.pRightJust " 20 %", 1.9, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3, True
    cPrint.pRightJust Format(ddv1, "standard"), 4, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 1.2, True
    cPrint.pRightJust "8.5 %", 1.9, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3, True
    cPrint.pRightJust Format(ddv2, "standard"), 4, True
    End If
    End If
    Dim pl As String
    
  
    pl = "Gotovina"
   
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Placilo: " & pl
    cPrint.pPrint Getnazi("select konec1 from oblikar")
    cPrint.pPrint Getnazi("select konec2 from oblikar")
    cPrint.pPrint Getnazi("select konec3 from oblikar")
    cPrint.pPrint Getnazi("select konec4 from oblikar")
    cPrint.pPrint Getnazi("select konec5 from oblikar")
    'cPrint.pPrint "", 0.1, False
    'cPrint.pPrint "", 0.1, False
    ' cPrint.pPrint "", 0.1, False
     ' cPrint.pPrint "", 0.1, False
     '  cPrint.pPrint "", 0.1, False
     '   cPrint.pPrint "", 0.1, False
     '   cPrint.pPrint "", 0.1, False
     cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
    
Call Shell("print /d:LPT2 c:\be.txt", 6)
    'cPrint.pPrint Chr(27), 0.1, False
   ' predal
   ' odrez
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
     ' If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
 
 
'Print #1, "======================================="
'Print #1, "SKUPAJ SIT  ", Format(asd, "0.00")
'Print #1,
'Print #1, "SKUPAJ EUR  ", Format(asd / DLookup("[eur]", "eur"), "0.00")

'Print #1, "---------------------------------------"
'Print #1, "Osnova DDV-a  DDV  Znesek DDV  Vrednost"
'If ddv > 0 Then
'Print #1, "  " & Format(ddv, "0.00") & "   20.00 %  " & Format(zneddv - ddv, "0.00") & "  " & Format(zneddv, "0.00")
'End If
'If ddv1 > 0 Then
'Print #1, "  " & Format(ddv1, "0.00") & "    8.50 %  " & Format(zneddv1 - ddv1, "0.00") & "  " & Format(zneddv1, "0.00")



'End If
'Print #1, "---------------------------------------"


'Print #1, "---------------------------------------"'
'End If

'Call Shell("print /d:LPT1 c:\be.txt", 6)
   
End Sub

Private Sub predal()
If FileExist("c:\be1.txt") Then
Else
Open "c:\be1.txt" For Output As #1
Print #1, Chr(7)
'Print #1, Chr(27) & Chr(105)
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
'Print #1, Chr(27) & "p" & Chr(0) & Chr(25) & "·"
Close #1
End If
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be1.txt", 6)
'Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & "c:\be1.txt", 6)
   
End Sub
Private Sub odrez()
'Open "c:\be.txt" For Output As #1
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, Chr(27) & Chr(105)
'Print #1, Chr(7)
'
'Print #1, Chr(27) & Chr(100) & Chr(48)
'
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
'Close #1
Call Shell("print /d:lpt2 c:\be.txt", 6)
   
End Sub
Private Sub coda852()
'Open "c:\be1.txt" For Output As #1
'Print #1, Chr(27) & Chr(116) & Chr(18)
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
'Close #1
'Call Shell("print /d:LPT1 c:\be1.txt", 6)
   
End Sub

Private Sub xcKeypad1_Key(KeyPressed As Variant)
 Select Case KeyPressed
        Case "BS"
            If Len(Me.Text3.Text) > 0 Then
                Me.Text3.Text = (Left$(Me.Text3.Text, Len(Me.Text3.Text) - 1))
            End If
        Case "Sign"
        Me.Text3.Text = Val(Replace(Me.Text3.Text, ",", ".")) * -1
        Case "Clear"
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            Me.Text3.Text = ""
           Case "ZAPRI"
           If Text3.Text = "" Then
           Text3.Text = 0
           End If
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            myConection.Execute ("update trenutna set kol=" & Me.Text3.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
            myConection.Execute ("update trenutna set znes=kol*(1-(pop/100))*cena where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
            osssv
            Me.Text3.Text = 1
            Me.Text2.Text = ""
            Me.Text1.Text = ""
            Me.pop.Text = 0
            Me.Text1.SetFocus
            Me.xcKeypad1.Visible = False
        Case Else
            Me.Text3.Text = (Me.Text3.Text & KeyPressed)
    End Select
End Sub


Public Sub zakljucc_Click(Index As Integer)
'MsgBox (Pblagajna)
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='NACA'") = "D" Then
vena_Click
End If
txtInvoiceNo.Text = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)

If jedobavnica <> "" Then
myConection.Execute ("UPDATE DOBAVN SET FAKTURA='" & Trim(novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)) & "' where stranka='" & jedobavnica & "' and faktura='.'")
jedobavnica = ""
End If
If stevnaro <> "" Then
myConection.Execute ("UPDATE nabasif SET kopija='" & Trim(novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)) & "' where tip_dok+id_dok='" & stevnaro & "'")

End If
If rs.State = 1 Then rs.Close
 rs.Open "update nabasif set org=" & Val(Getnazi(" select davcna from partner where naziv='" & Getnazi("select dod0 from glavna where tip_dok+id_dok='" & stevnaro & "'") & "'")) & ",poknj='K' where tip_dok+id_dok='" & stevnaro & "'", myConection, adOpenDynamic, adLockOptimistic
 
Dim rsa1 As New ADODB.Recordset
nacra = 1
rsa1.Open "select sifra,naziv,kol,format(cena,'fixed') as cena,format(znes,'fixed') as znesek,kopija,X,pop from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic

If Not rsa1.EOF Then
 'If Me.MSHFlexGrid1.Rows > 1 Then
'Me.karto.Visible = False
    Call LaVolpeButton46_Click
   ' Me.pop.Text = 0
End If
   
     
     
End Sub

Private Function FileExist(FileName As String) As Boolean

  On Error GoTo FileDoesNotExist
  
  Call FileLen(FileName)
  FileExist = True
  Exit Function
  
FileDoesNotExist:
  FileExist = False
  
End Function

Private Sub zaprr_Click()
Me.tvorbara.Visible = False

Me.Frame2.Visible = False
Me.zaprr.Visible = False
End Sub
