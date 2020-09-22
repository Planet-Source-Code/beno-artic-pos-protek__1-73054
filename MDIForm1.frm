VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.MDIForm frmMAIN 
   BackColor       =   &H8000000C&
   Caption         =   "Poslovanje"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   -60
   ClientWidth     =   14700
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FF8080&
      FillColor       =   &H00E0E0E0&
      Height          =   8625
      Left            =   0
      ScaleHeight     =   8565
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   38
         Top             =   3720
         Width           =   2775
         Begin VB.Image MenuHeader 
            Height          =   495
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avtomobili"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   13
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relacije"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   11
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vnos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   9
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   1440
            Width           =   465
         End
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "POTNI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   3
            Left            =   150
            TabIndex        =   39
            ToolTipText     =   "Choose Tasks"
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   4
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   2775
         TabIndex        =   23
         Top             =   7320
         Width           =   2775
         Begin MSComctlLib.ProgressBar ProgressBar 
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   25
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   16
         Top             =   3240
         Width           =   2775
         Begin LVbuttons.LaVolpeButton LaVolpeButton2 
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "PREVZEM"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":091E
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
            Height          =   495
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "POS-BLAG."
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":093A
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   20
            Top             =   1680
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DOBAVNICA"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0956
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   2160
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "PREDRACUN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0972
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   2
            Left            =   360
            TabIndex        =   22
            Top             =   2640
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "FAKTURA"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":098E
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   3
            Left            =   360
            TabIndex        =   27
            Top             =   3120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":09AA
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Visible         =   0   'False
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":09C6
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   5
            Left            =   0
            TabIndex        =   29
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":09E2
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   6
            Left            =   0
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":09FE
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   7
            Left            =   0
            TabIndex        =   31
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0A1A
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   8
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0A36
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   9
            Left            =   0
            TabIndex        =   33
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0A52
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   10
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0A6E
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   11
            Left            =   0
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0A8A
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   12
            Left            =   0
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0AA6
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
         Begin LVbuttons.LaVolpeButton PREG 
            Height          =   495
            Index           =   13
            Left            =   0
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "DN"
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
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "MDIForm1.frx":0AC2
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
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "DOKUMENTI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Choose Tasks"
            Top             =   120
            Width           =   1785
         End
         Begin VB.Image MenuHeader 
            Height          =   615
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   2775
         TabIndex        =   10
         Top             =   2760
         Width           =   2775
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "AKCIJE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Choose Tasks"
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TDR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   8
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":0ADE
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   1455
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pregled prodaje"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   7
            Left            =   240
            MouseIcon       =   "MDIForm1.frx":0DE8
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Negotovinski"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   6
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":10F2
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pregled zakljuckov"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   5
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":13FC
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   2160
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pregled nabave"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   4
            Left            =   255
            MouseIcon       =   "MDIForm1.frx":1706
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   720
            Width           =   1485
         End
         Begin VB.Image MenuHeader 
            Height          =   495
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox FrameXPMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Index           =   0
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   2775
         TabIndex        =   4
         Top             =   240
         Width           =   2775
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zaposleni"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   10
            Left            =   360
            MouseIcon       =   "MDIForm1.frx":1A10
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   2160
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Artikli"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   2
            Left            =   360
            MouseIcon       =   "MDIForm1.frx":1D1A
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dobavitelji"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   1
            Left            =   360
            MouseIcon       =   "MDIForm1.frx":2024
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stranke"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   0
            Left            =   360
            MouseIcon       =   "MDIForm1.frx":232E
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kategorije"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Index           =   3
            Left            =   360
            MouseIcon       =   "MDIForm1.frx":2638
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "šIFRANT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   0
            Left            =   15
            TabIndex        =   6
            ToolTipText     =   "Choose Tasks"
            Top             =   120
            Width           =   1005
         End
         Begin VB.Image MenuHeader 
            Height          =   495
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComctlLib.ImageList i32x32 
         Left            =   6240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   27
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":2942
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":361C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3D70
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":4A4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":5724
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":63FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":70D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":7DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":8A8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":9766
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":A440
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":B11A
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":BDF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":CACE
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":D7A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":E482
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":F15C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":FE36
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":10B10
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":117EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":11C3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":12092
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":12592
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":126D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1281A
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":12942
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":12A96
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList SmallImages 
         Left            =   10560
         Top             =   150
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   42
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":12B9A
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":13874
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":16028
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":16E7A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":17CCC
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":185A6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":18E80
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1975A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1A124
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1A9FE
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1AD18
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1B5F2
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1BECC
               Key             =   "IMG12"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1C7A6
               Key             =   "IMG13"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1CAC0
               Key             =   "IMG14"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1D39A
               Key             =   "IMG15"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1DC74
               Key             =   "IMG16"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1E54E
               Key             =   "IMG17"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1EE28
               Key             =   "IMG18"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1F702
               Key             =   "IMG19"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1FFDC
               Key             =   "IMG20"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":208B6
               Key             =   "IMG21"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":21190
               Key             =   "IMG22"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":21A6A
               Key             =   "IMG23"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":22344
               Key             =   "IMG24"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":22C1E
               Key             =   "IMG25"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":234F8
               Key             =   "IMG26"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":23DD2
               Key             =   "IMG27"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":246AC
               Key             =   "IMG28"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":24F86
               Key             =   "IMG29"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":25860
               Key             =   "IMG30"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":26116
               Key             =   "IMG31"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":269F0
               Key             =   "IMG32"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":26E42
               Key             =   "IMG33"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":27294
               Key             =   "IMG34"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":29A46
               Key             =   "IMG35"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":29BF6
               Key             =   "IMG36"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":29CFA
               Key             =   "IMG37"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":29DE6
               Key             =   "IMG38"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":29F0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":2A076
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":2A1BA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgUp 
         Height          =   375
         Left            =   0
         Picture         =   "MDIForm1.frx":2A30A
         Top             =   6000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image imgDown 
         Height          =   375
         Left            =   0
         Picture         =   "MDIForm1.frx":2D998
         Top             =   6480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSForms.Label Label6 
         Height          =   135
         Left            =   120
         TabIndex        =   2
         Top             =   7200
         Visible         =   0   'False
         Width           =   135
         ForeColor       =   -2147483640
         BackColor       =   -2147483636
         Size            =   "238;238"
         FontHeight      =   165
         FontCharSet     =   238
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox iml16 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14670
      TabIndex        =   1
      Top             =   0
      Width           =   14700
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8625
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Picture         =   "MDIForm1.frx":31026
            Text            =   "Uporabnik:"
            TextSave        =   "Uporabnik:"
            Object.ToolTipText     =   "Trenutni uporabnik"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Cakam.."
            TextSave        =   "Cakam.."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2937
            Picture         =   "MDIForm1.frx":31E7A
            Text            =   "Cas prijave:"
            TextSave        =   "Cas prijave:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "cakam.."
            TextSave        =   "cakam.."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2/26/2008"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Text            =   "Caps Lock"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Text            =   "Num Lock"
            TextSave        =   "Num Lock"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Text            =   "Insert"
            TextSave        =   "Insert"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuTop 
      Caption         =   "Datoteka"
      Begin VB.Menu mnuFileNew 
         Caption         =   "Nova      "
         Begin VB.Menu mnuNew 
            Caption         =   "Stranke...."
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Artikli"
            Index           =   3
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Kategorije"
            Index           =   4
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Printaj"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Nastavi printanje"
      End
      Begin VB.Menu mnuPrintPrv 
         Caption         =   "Predogled printanja"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Shrani"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Izhod"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Urejanje"
      Visible         =   0   'False
      Begin VB.Menu mnuModify 
         Caption         =   "Uredi "
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Briši       "
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Podrobnosti      "
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Osveži"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Najdi"
      End
   End
   Begin VB.Menu mnUsers 
      Caption         =   "Uporabniki"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Dodaj uporabnika"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Briši uporabnika"
      End
      Begin VB.Menu mnuChangeUsername 
         Caption         =   "Spremeni username"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Spremeni  Password"
      End
      Begin VB.Menu mnuViewall 
         Caption         =   "Pregled vseh uporabnikov"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Orodja"
      Begin VB.Menu mnuvoz 
         Caption         =   "UVOZ"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CdlgEx1 As New CdlgEx
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Const speed = 150 'Speed of he menu animations
Dim i As Integer
Dim expand As Boolean 'Tells the timer wether to expand or contract the menus
Dim frame As Integer 'Specifys the menuon  to resize
Dim framex As Integer
Dim preframe As Integer 'Specifys the menu to resize
Dim dex As Integer
Dim doresizex As Boolean
Dim doresize As Boolean 'Tells the timer to resize the menus




Private Sub lblHeader_Click(Index As Integer)
 MenuHeader_Click (Index)
End Sub

Private Sub MenuHeader_Click(Index As Integer)

       
        
           


    'If not currently resizeing allow menu to be resized
    If doresize = False Then
        
        'Works out if the menu needs to expand or contract
        If FrameXPMenu(Index).Height = MenuHeader(Index).Height Then 'minimised
            expand = True
        Else
            expand = False
        End If
        
        'Tell the timer to do the resizing to what frame
        doresize = True
        frame = Index
        
    End If


End Sub

Public Sub size(frm As Form)
    frm.Width = Me.ScaleWidth
    frm.Height = Me.ScaleHeight
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim tempSql As String
    frmControlMain.MSHFlexGrid1.Visible = True
    frmControlMain.WBrow.Visible = False
    frmControlMain.Show
End Sub



Private Sub Command2_Click()
'MODIFYID = "100"
'ADDING = False
'C_frmProduct.Show vbModal

End Sub

Private Sub Command3_Click()
Form2.Show vbModal
End Sub

Private Sub PREG_Click(Index As Integer)
CatalogueName = "MATE"
frmControlMain.WBrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
tip_dok = Left(Me.PREG(Index).Caption, 2)
If dejpre = 1 Then
        'SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from storno where [datum] between #" & dod & "# AND #" & ddo & "# and stw<>'A' group by st order by st"
          SQL = "Select tip_dok,id_dok,min(stdok) as stdok,min(datum) as datum,sum(cena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif  where [datum] between #" & dod & "# AND #" & ddo & "# and tip_dok='" & Left(PREG(Index).Caption, 2) & "' group by tip_dok,id_dok order by tip_dok,id_dok"
        Else
           SQL = "Select tip_dok,id_dok,min(stdok) as stdok,min(datum) as datum,sum(cena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif where tip_dok='" & Left(PREG(Index).Caption, 2) & "' group by tip_dok,id_dok order by tip_dok,id_dok"
   End If
        'CatalogueName = "Purchase Registry"
'Form7.Show
Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "" Then

Set Rs1 = DCON.Execute(SQL)
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
    frmControlMain.osv_Click
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing

End Sub

Private Sub Label1_Click(Index As Integer)
cst = Index
    frmControlMain.WBrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
    Select Case Index
    Case 9
         SQL = "Select * from potni "
        CatalogueName = "potni"
    Case 0
         SQL = "Select * from partner "
        CatalogueName = "Customer"
    Case 13
         SQL = "Select * from avto "
        CatalogueName = "avtom"
    Case 11
         SQL = "Select * from relacija "
        CatalogueName = "relacija"
    Case 10
         SQL = "Select * from zaposleni "
        CatalogueName = "zaposleni"
    Case 1
       ' Sql = "Select * from Supplier WHERE SuppliersID<>'CASH'"
        SQL = "Select * from partner "
        CatalogueName = "Supplier"
    Case 2
    'If CatalogueName = "Category" Then
   'Else
    'Call zaloga
    'End If
     SQL = "Select madasifr,madanazi,madampcd,madazalo from mada order by madagrup,madanazi"
     CatalogueName = "Category"
    Case 3
        SQL = "Select * from grupa order by sifra"
        CatalogueName = "Location"
    Case 4
'        SQL = "Select * from PurchaseOrderHeader"
        SQL = "Select stdok,min(datum) as datum,sum(nabcena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif group by stdok"
        CatalogueName = "Purchase Order"
    Case 5
'        SQL = "Select * from PurchaseReturnHeader"
        If dejpre = 1 Then
        SQL = "Select datum,  sum(znesek) as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by datum order by datum"
        Else
        SQL = "Select datum,  sum(znesek) as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif group by datum order by datum"
        End If
        CatalogueName = "Purchase Return"
    Case 6
        If dejpre = 1 Then
        'SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from storno where [datum] between #" & dod & "# AND #" & ddo & "# and stw<>'A' group by st order by st"
          SQL = "Select tip_dok,id_dok,min(stdok) as st_dokument,min(datum) as datum,sum(nabcena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif  where [datum] between #" & dod & "# AND #" & ddo & "# group by tip_dok,id_dok order by tip_dok,id_dok"
        Else
           SQL = "Select tip_dok,id_dok,min(stdok) as st_dokument,min(datum) as datum,sum(nabcena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif group by tip_dok,id_dok order by tip_dok,id_dok"
         End If
        CatalogueName = "Purchase Registry"
    Case 7
     If dejpre = 1 Then
        SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from racusif group by st order by st"
         End If
        CatalogueName = "Sales Return"
    Case 8
    Call tdrr
      If dejpre = 1 Then
        SQL = "Select * from tdr where [datum] between #" & dod & "# AND #" & ddo & "#"
        Else
        SQL = "Select * from tdr"
        End If
        CatalogueName = "Sales Registry"
    End Select
    

Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "" Then

Set Rs1 = DCON.Execute(SQL)
ssqq = SQL
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
    frmControlMain.osv_Click
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing

    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Call moveShape(Shape1, Label1(Index))
End Sub
Public Function moveShape(shape As Object, Cntrl As Object)
        shape.Visible = True
        shape.Move Cntrl.Left - 150, Cntrl.Top - 60, 1845, 300
End Function


Private Sub Label5_dblClick()
If odp = 0 Then
odp = 1
Else
odp = 0
End If
End Sub

Public Sub Label6_Click()
Call fresh
End Sub

Private Sub LaVolpeButton1_Click()
frmsalesbill.Show
End Sub

Private Sub LaVolpeButton2_Click()
frmPR.Show
 CatalogueName = "MATE"
End Sub

Private Sub LaVolpeButton3_Click()

End Sub

Private Sub lblTask_Click()
    Load frmControlMain
    frmControlMain.MSHFlexGrid1.Visible = False
    frmControlMain.WBrow.Visible = True
    
    
    Dim SqLargs As String
    SqLargs = "SELECT madasifr,madanazi,madazalo,madampcd From mada WHERE ((madazalo)<=0) Order by madazalo DESC"
    Call frmControlMain.CreateStartPage(SqLargs)
End Sub

Private Sub MDIForm_Resize()
    Call size(frmControlMain)
End Sub
Private Sub mnuCashBook_Click()
    Form1.Show
End Sub
Private Sub mnuDaybook_Click()
    GetNewConnection2
   Call frmControlMain.CreateDataPage("Select  TOP 2 * From V1", "Day book")
    
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
End
End Sub

Private Sub mnPret_Click()
'ADDING = True
'Load tempPRET
'tempPRET.Show vbModal

End Sub

Private Sub mnuAddUser_Click()
    Load frmAddUser
    frmAddUser.Show vbModal
End Sub

Private Sub mnuBackup_Click()
Load frmBackup
frmBackup.Show vbModal

End Sub

Private Sub mnuChangePassword_Click()
Load frmChange
frmChange.Show vbModal

End Sub

Private Sub mnuChangeUsername_Click()
Load frmuser
frmuser.Show vbModal

End Sub

Private Sub mnuDelete_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Select Case CatalogueName



  Case "Customer"

        
          If MsgBox("Zbrišem Stranko?", vbInformation + vbYesNo) = vbYes Then
          
            SQL = "Delete From partner Where sifra=" & frmControlMain.MSHFlexGrid1.text
            End If
        
       
  Case "Supplier"
        
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
            SQL = "Delete From partner Where sifra=" & frmControlMain.MSHFlexGrid1.text
         End If
       
        
  Case "Category"
     ' Set Rs1 = myConection.Execute("Select * From mada where madasifr=" & frmControlMain.DataGrid1.Columns(0).text)
      '  If Rs1.RecordCount = 0 Then
        If MsgBox("Zbrišem Artikel?", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from mada Where madasifr=" & Val(frmControlMain.MSHFlexGrid1.text)
        End If
       'Else
        '    MsgBox "Neuspešno!", vbInformation
           
        'End If
        
  Case "Location"
    If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from grupa Where sifra=" & Val(frmControlMain.MSHFlexGrid1.text)
    End If
  Case "Purchase Order"
    Set Rs1 = DCON.Execute("Select * From nabasif where stdok='" & frmControlMain.MSHFlexGrid1.text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.text))
        'If Rs1.RecordCount = 0 Then
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from nabasif Where stdok='" & frmControlMain.MSHFlexGrid1.text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.text)
        End If
       
  Case "Purchase Return"
     MsgBox "Neuspešno!", vbInformation
         
        'Delete from PurchaseReturnDetail Where PurchaseReturnID='" & "'"
        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Purchase Registry"
     MsgBox "Neuspešno!", vbInformation
         
       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
       'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Sales Return"
        If MsgBox("Naredim storno raèuna " & frmControlMain.MSHFlexGrid1.text & " ? ", vbInformation + vbYesNo) = vbYes Then
        DCON.Execute "insert into storno select * from racusif where st=" & Val(frmControlMain.MSHFlexGrid1.text)
        SQL = "Delete from racusif Where st=" & Val(frmControlMain.MSHFlexGrid1.text)
    End If
       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Sales Registry"
     MsgBox "Neuspešno!", vbInformation
         
        'Delete from SalesRegistryDetail Where SalesRegistryID='" & "'"
        'Delete from SalesRegistryHeader Where SalesRegistryID='" & "'"
End Select

DCON.Execute SQL

'Call GridRefresh
Call fresh


Set DCON = Nothing
End Sub

Private Sub mnuDeleteUser_Click()
SQL = "Select username1 from users where username1 <> '" & CurUser & "'"
Load frmDelUser
frmDelUser.Show vbModal

End Sub

Private Sub mnuDetails_Click()
On Error GoTo adder:
If frmControlMain.MSHFlexGrid1.Rows >= 1 Then
Select Case (CatalogueName)
  Case "Purchase Order"
    CreateH_Page "Select *  from nabasif where  stdok='" & frmControlMain.MSHFlexGrid1.text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.text) & "", " Prevzemni list št.: " & frmControlMain.MSHFlexGrid1.text
  Case "Purchase Return"
        CreateH_Page "Select datum,min(st) as zac,max(st) as kon,sum(znesek) as znesek from racusif  where datum=#" & (frmControlMain.MSHFlexGrid1.text) & "# group by datum", " Rekapitulacija "
  Case "Purchase Registry"
    CreateH_Page "Select sifra,naziv,kol,znesek from storno where st=" & Val(frmControlMain.MSHFlexGrid1.text), " Raèun št: " & frmControlMain.MSHFlexGrid1.text & ", z dne : " & frmControlMain.MSHFlexGrid1.text
   Case "Sales Return"
        CreateH_Page "Select sifra,naziv,kol,znesek from racusif where st=" & Val(frmControlMain.MSHFlexGrid1.text), " Raèun št: " & frmControlMain.MSHFlexGrid1.text & ", z dne : " & frmControlMain.MSHFlexGrid1.text
   Case "Sales Registry"
        CreateH_Page "Select * from tdr", " Details "
   Case "Category"
 
        CreateH_Page "select madasifr,madanazi,madazalo,madampcd from mada", "  Podrobnosti zalog "
  Case Else
           MakeShortReport SQL, " Podrobnosti za"
End Select
End If
Exit Sub
adder:
End Sub

Private Sub mnuExit_Click()
On Error GoTo adder
   PostMessage Me.hWnd, WM_SYSCOMMAND, SC_CLOSE, 0
   
'little bit of blabla this method is very widely used in Scripting Lanuguage very powerful code but very gentle
adder:
Exit Sub
End Sub

Private Sub mnuFind_Click()
Dim sFind As String

sFind = InputBox("Najdi zapis", "Record")
sFind = Replace(sFind, "'", "", 1, Len(sFind), vbTextCompare)


If sFind <> "" Then
  Select Case CatalogueName
  
  Case "Customer"
         Call GRIDBIND("Customer", frmControlMain.MSHFlexGrid1, " Where customerid like'" & sFind & "%' OR Company like'" & sFind & "%'")

  Case "Supplier"
         Call GRIDBIND("Supplier", frmControlMain.MSHFlexGrid1, " Where suppliersid like'" & sFind & "%' Or BusinessName like'" & sFind & "%'")

  Case "Category"
         Call GRIDBIND("mada", frmControlMain.MSHFlexGrid1, " Where madanazi like'" & sFind & "%'")
  Case "Location"
         Call GRIDBIND("grupa", frmControlMain.MSHFlexGrid1, " Where grupa like'" & sFind & "%'")
  Case "Purchase Order"
        Call GRIDBIND("nabasif", frmControlMain.MSHFlexGrid1, " Where stdok like'" & sFind & "%'")

  Case "Purchase Return"
            Call GRIDBIND("PurchaseReturnHeader", frmControlMain.MSHFlexGrid1, " Where PurchaseReturnID like'" & sFind & "%'")
  Case "Purchase Registry"
      Call GRIDBIND("PurchaseRegistryHeader", frmControlMain.MSHFlexGrid1, " Where PurchaseRegistryID like'" & sFind & "%'")
  Case "Sales Return"
            Call GRIDBIND("racusif", frmControlMain.MSHFlexGrid1, " Where st=" & sFind)
  Case "Sales Registry"
            ' Call GRIDBIND("SalesRegistryHeader", frmControlMain.DataGrid1, " Where SalesRegistryID like'" & sFind & "%'")
    
End Select

End If

End Sub

Private Sub mnuModify_Click()
ADDING = False
MODIFYID = frmControlMain.MSHFlexGrid1.text


Select Case CatalogueName

 Case "MATE"
 ma_ured = 1
 blag.Show
 
  Case "Customer"
    Load C_frmCustomer
    C_frmCustomer.Show vbModal
  Case "Supplier"
 
  std = Trim(frmControlMain.MSHFlexGrid1.text) & Trim(frmControlMain.MSHFlexGrid1.text)
  MsgBox std
  Case "Category"
    Load frmProdEntry
    frmProdEntry.ShowEdit MODIFYID
  Case "Location"
    Load C_frmLocation
    C_frmLocation.Show vbModal
 
  Case "Purchase Order"
 ' edi = 1
 ' dob = Val(frmControlMain.DataGrid1.Columns(3).text)
 ' std = frmControlMain.DataGrid1.Columns(0).text
 '   frmPR.Show
     'MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
   ' tempPO.Show vbModal
   

''
''       SQL = "Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
''       SQL = "Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Purchase Return"
   ' tempPRET.Show vbModal
    MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'        'Delete from PurchaseReturnDetail Where PurchaseReturnID='" & "'"
'        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Purchase Registry"
        MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
'       'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Sales Return"
 MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
  
'       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
'        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
 Case "Sales Registry"
 'MsgBox "Transaction has Already been Recieved:" & vbCrLf & "Please Use Return Modules For Returning.", vbInformation
 
  
'        'Delete from SalesRegistryDetail Where SalesRegistryID='" & "'"
        'Delete from SalesRegistryHeader Where SalesRegistryID='" & "'"
End Select
    


    
End Sub


Private Sub mnuPO_Click()
'ADDING = True
'Load tempPO
'tempPO.Show vbModal

End Sub

Private Sub mnuPreg_Click()
' ADDING = True
Load frmPR
frmPR.Show vbModal

End Sub

Private Sub mnuPrint_Click()
On Error GoTo adder:
    If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPrintPrv_Click()
On Error GoTo adder:
    If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub
Private Sub mnuPageSetup_Click()
On Error GoTo adder:
If frmControlMain.WBrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End If
    Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPurchaseRegister_Click()
 frmControlMain.WBrow.Visible = True
    frmControlMain.MSHFlexGrid1.Visible = False
    
rptState = "PurchaseRegistry"

Load Form1
Form1.Show vbModal

End Sub

Private Sub mnuRefresh_Click()
    'Call GridRefresh
    Call fresh
End Sub

Private Sub mnuSalesRegister_Click()
 frmControlMain.WBrow.Visible = True
    frmControlMain.MSHFlexGrid1.Visible = False
rptState = "SalesRegistry"

Load Form1
Form1.Show vbModal



End Sub

Private Sub mnuSave_Click()
'    On Error GoTo adder:
'If frmControlMain.wbrow.Visible = True Then
        frmControlMain.WBrow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER, 1
        
'End If
    'Exit Sub
'adder:
'    Exit Sub
End Sub

Private Sub mnuNew_Click(Index As Integer)
'0   customer
'1   vendor
'2   (space)
'3   product
'4   cat
'5   engmod
'6   (space)
'7   uom
'8   (space)
'9   bank
'10   Location
Select Case Index
    Case 0
    ADDING = True
    C_frmCustomer.Show vbModal
    
    Case 1
    ADDING = True
    'C_frmSupplier.Show vbModal
    
    Case 3
    ADDING = True
    frmProdEntry.ShowAdd
    
    Case 4
     ADDING = True
    
    C_frmLocation.Show vbModal
    Case 10
    ADDING = True
   C_frmCategory.Show vbModal
    
End Select
End Sub



Private Sub mnuSReg_Click()
' ADDING = True
'Load tempSalesReg
'tempSalesReg.Show vbModal

End Sub

Private Sub mnuSRet_Click()
' ADDING = True
'Load frmSalesReturn
'frmSalesReturn.Show vbModal

End Sub

Private Sub MDIForm_Load()
osve = 0
'setBitmaps
'makebar
'LoadtabLeft
'

Dim i As Integer
 For i = 0 To MenuHeader.Count - 1
        'Decide if the menu is already contracted and if so display the expand header
        If FrameXPMenu(i).Height = imgUp.Height Then
            MenuHeader(i).Picture = imgDown.Picture
        Else
            MenuHeader(i).Picture = imgUp.Picture
        End If
        MenuHeader(i).Height = 375
        MenuHeader(i).Width = FrameXPMenu(i).Width
    Next
    
    doresize = False
'lblHeader_Click (1)
'lblHeader_Click (2)



Dim RSS As New ADODB.Recordset
RSS.Open "select * from dokumenti", myConection, adOpenDynamic, adLockOptimistic
RSS.MoveFirst
Dim a As Integer

a = 0
Do While Not RSS.EOF
If RSS.EOF Then
Exit Sub
End If

Me.PREG(a).Visible = True
Me.PREG(a).Caption = RSS.Fields(0) & "-" & RSS.Fields(1)
Me.PREG(a).Font.size = 12
If a > 3 Then
Me.PREG(a).Top = Me.PREG(0).Height + Me.PREG(a - 1).Top
Me.PREG(a).Left = Me.PREG(0).Left
End If
a = a + 1
RSS.MoveNext
Loop

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim FrameExpandHeight As Integer: FrameExpandHeight = 0 'Height for menu to expand to
    Dim i As Integer
    
    'Adjust top of each menu
    For i = 1 To FrameXPMenu.Count - 1
        FrameXPMenu(i).Top = FrameXPMenu(i - 1).Top + FrameXPMenu(i - 1).Height + 120
    Next
    
    'Animates a menu if one is resizing
    If doresize = True Then
        
        If expand = False Then 'If not expanding
            MenuHeader(frame).Picture = imgDown.Picture 'Change headder pic
            FrameXPMenu(frame).Height = FrameXPMenu(frame).Height - speed 'Animate
            If FrameXPMenu(frame).Height <= MenuHeader(frame).Height Then doresize = False: FrameXPMenu(frame).Height = MenuHeader(frame).Height 'If we are fully contracted no need to do further resizing
        Else
            MenuHeader(frame).Picture = imgUp.Picture 'Change headder pic
            FrameXPMenu(frame).Height = FrameXPMenu(frame).Height + speed 'Animate
            
            'This finds out the bottom position of the last control in a menu and limits expanding appropiatly
            For i = 0 To Me.Count - 1
                'If the control is in the frame we are resizing
                If Controls(i).Container.Name = FrameXPMenu(frame).Name Then
                    If Controls(i).Container.Index = FrameXPMenu(frame).Index Then
                        'If the control top value is larger than any previous, prepare to resize to it
                        If Controls(i).Top + Controls(i).Height > FrameExpandHeight Then
                            FrameExpandHeight = Controls(i).Top + Controls(i).Height
                        End If
                    End If
                End If
            Next
            If FrameXPMenu(frame).Height >= FrameExpandHeight + 120 Then doresize = False: FrameXPMenu(frame).Height = FrameExpandHeight + 120 'Stop the frame from resizing
        End If
    End If

End Sub



Sub LoadtabLeft()

'tabLeft.Pinned = False
'tabLeft.ImageList = ImageList1
'   Dim tabX As cTab
'   Set tabX = tabLeft.Tabs.Add("Main", , "Main", 8)
'        tabX.Panel = Picture1
'    Set tabX = tabLeft.Tabs.Add("Catalogue", , "Catalogue", 2)
'        tabX.Panel = picSearch'

   'Set tabX = tabLeft.Tabs.Add("StockINQ", , "Stock Inquiry", 3)
   'Set tabX = tabLeft.Tabs.Add("EXPLORER2", , "Explorer", 4)
        'tabX.Panel = Picture1
 '  Set tabX = tabLeft.Tabs.Add("EXPLORER3", , "Explorer", 5)
End Sub

Sub setBitmaps()
'    With PopMenu1
'    .SubClassMenu Me
'    .ImageList = ImageList1
'    .ItemIcon("mnuFileNew") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuNew(0)") = ImageList1.ListImages("C1").Index - 1
'    .ItemIcon("mnuNew(1)") = ImageList1.ListImages("C2").Index - 1
'    .ItemIcon("mnuNew(3)") = ImageList1.ListImages("C3").Index - 1
'    .ItemIcon("mnuNew(4)") = ImageList1.ListImages("C4").Index - 1
'    .ItemIcon("mnuNew(5)") = ImageList1.ListImages("C5").Index - 1
'    .ItemIcon("mnuNew(7)") = ImageList1.ListImages("C5").Index - 1
 ''   .ItemIcon("mnuFind") = ImageList1.ListImages("C7").Index - 1
 '   .ItemIcon("mnuDelete") = ImageList1.ListImages("C6").Index - 1
  '  .ItemIcon("mnuModify") = ImageList1.ListImages("C8").Index - 1
 '   .ItemIcon("mnuDetails") = ImageList1.ListImages("C9").Index - 1
 '   .ItemIcon("mnuDetails") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuCashBook") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuLedger") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuGSummary") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuSalesRegister") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuPurchaseRegister") = ImageList1.ListImages("C9").Index - 1
'.ItemIcon("mnuDaybook") = ImageList1.ListImages("C9").Index - 1
''.ItemIcon("mnuSOA") = ImageList1.ListImages("C9").Index - 1
  '  .ItemIcon("mnuRecievable") = ImageList1.ListImages("C9").Index - 1
  '  .ItemIcon("mnuRecievable") = ImageList1.ListImages("C9").Index - 1
  '  .ItemIcon("mnuPayables") = ImageList1.ListImages("C9").Index - 1
  '  .ItemIcon("mnuSALedger") = ImageList1.ListImages("C9").Index - 1'

'    .ItemIcon("mnuStockItem") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuGroupSummary") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuMovementAnalysis") = ImageList1.ListImages("C9").Index - 1
'   .ItemIcon("mnuPhysicalStockRegister") = ImageList1.ListImages("C9").Index - 1
'    .ItemIcon("mnuInventoryStatement") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuLocation") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuReorderStatus") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuCategories") = ImageList1.ListImages("C9").Index - 1
'        .ItemIcon("mnuPendings") = ImageList1.ListImages("C9").Index - 1
 ''   .UnsubclassMenu
 '   End With
End Sub
''///////////////////////////////////////////////
Sub makebar()


'Dim barX As cListBar
'Dim itmX As cListBarItem
'Dim i As Long
   ' With listbar1
'        .ImageList(evlbLargeIcon) = iml16 ' ilsIcons32
'        .ImageList(evlbSmallIcon) = iml16 'ilsIcons32
'//////////Catalogue Entries
 '     Set barX = .Bars.Add("Catalogue", , "Catalogue")
 '       Set itmX = barX.Items.Add("AutoCompany", , "AutoCompany", 16)
 '           itmX.HelpText = "Catalogue Your Information By Categorizing record for AutoCompany"
 '       Set itmX = barX.Items.Add("EngineModel", , "EngineModel", 2)
 '           itmX.HelpText = "Cataglogue your Information By Assigning EngineModel"
 '       Set itmX = barX.Items.Add("Location", , "Location", 3)
 '           itmX.HelpText = "Add/Edit/Remove Location for Stock and Inventory"
 '       Set itmX = barX.Items.Add("Category", , "Category", 4)
  ''          itmX.HelpText = "Add/Edit/Remove Category For easy Groupings"
  '      Set itmX = barX.Items.Add("Product", , "Product", 5)
  '          itmX.HelpText = "Add/Edit/Remove Products and Details"
  '      Set itmX = barX.Items.Add("UOM", , "Units of Measures", 6)
  '              itmX.HelpText = "Enter Units of Measures for Packaging"
'//''////////Accounts
   '   Set barX = .Bars.Add("Account", , "Accounts")
   '         Set itmX = barX.Items.Add("Supplier", , "Supplier", 7)
   '             itmX.HelpText = ""
   '         Set itmX = barX.Items.Add("Customer", , "Customer", 8)
   '             itmX.HelpText = ""
   '         Set itmX = barX.Items.Add("Bank", , "Banks", 8)
   '             itmX.HelpText = ""
'//'////////Vouchers
    '    Set barX = .Bars.Add("Voucher", , "Vouchers")
    '        Set itmX = barX.Items.Add("Payments", , "Payments", 10)
    '            itmX.HelpText = ""
    '        Set itmX = barX.Items.Add("Receipts", , "Receipts", 11)
    '            itmX.HelpText = ""
    '        Set itmX = barX.Items.Add("Deposit", , "Deposit", 12)
     '           itmX.HelpText = ""
     '       Set itmX = barX.Items.Add("DNote", , "Delivery Notes", 13)
     '           itmX.HelpText = ""
     '       Set itmX = barX.Items.Add("CNote", , "Counter Notes", 14)
     '           itmX.HelpText = ""
     '       Set itmX = barX.Items.Add("ExpBook", , "Expense Books", 15)
     '           itmX.HelpText = ""
     '   Set barX = .Bars.Add("Transaction", , "Transaction")
     '       Set itmX = barX.Items.Add("SalesReg", , "Sales Registry", 1)
      '          itmX.HelpText = "Record Purchases of Items"
 '           Set itmX = barX.Items.Add("SalesRet", , "Sales Return", 17) '
'                itmX.HelpText = "Make Vocher and Do Payments to Vendors'"
  '          Set itmX = barX.Items.Add("SalesOrd", , "Sales Order", 18)
  '              itmX.HelpText = ""
  '          Set itmX = barX.Items.Add("PurReg", , "Purchase Registry", 19)
  '              itmX.HelpText = ""
  '          Set itmX = barX.Items.Add("PurRet", , "Purchase Return", 20)
  '              itmX.HelpText = ""
   ''         Set itmX = barX.Items.Add("PurOrd", , "Purchase Order", 21)
   '             itmX.HelpText = ""
   '     Set barX = .Bars.Add("IBooks", , "Inventory Books")
   ''         Set itmX = barX.Items.Add("Stock Item", , "Stock Item", 22)
    '            itmX.HelpText = "Record Purchases of Items"
    '        Set itmX = barX.Items.Add("Group Summary", , "Group Summary", 23)
    '            itmX.HelpText = "Make Vocher and Do Payments to Vendors"
    '        Set itmX = barX.Items.Add("Physical Stock register", , "Physical Stock register", 24)
    '            itmX.HelpText = ""
    ''        Set itmX = barX.Items.Add("Inventory Statement", , "Inventory Statement", 25)
     '           itmX.HelpText = ""
     '       Set itmX = barX.Items.Add("Movement Analysis", , "Movement Analysis", 26)
    '            itmX.HelpText = ""

'        Set barX = .Bars.Add("ABooks", , "Accounts Books")
'            Set itmX = barX.Items.Add("Cash Book", , "Cash Book", 21)
'                itmX.HelpText = "Record Purchases of Items"
'            Set itmX = barX.Items.Add("Ledger", , "Ledger", 3)
'                itmX.HelpText = "Make Vocher and Do Payments to Vendors"
'            Set itmX = barX.Items.Add("Group Summary", , "Group Summary", 4)
 '               itmX.HelpText = ""
 ''           Set itmX = barX.Items.Add("Sales Register", , "Sales Register", 5)
 '               itmX.HelpText = ""
 '           Set itmX = barX.Items.Add("Purchase Register", , "Purchase Register", 6)
 '               itmX.HelpText = ""
  '          Set itmX = barX.Items.Add("Statement of Accounts", , "Statement of Accounts", 7)
  ''              itmX.HelpText = ""
  '          Set itmX = barX.Items.Add("Day Book", , "Day Book", 8)
  '              itmX.HelpText = ""
'


 '       .Bars(1).OfficeXpStyle = True
 '       .Bars(2).OfficeXpStyle = True
 '       .Bars(3).OfficeXpStyle = True
 '       .Bars(4).OfficeXpStyle = True
 '       .Bars(5).OfficeXpStyle = True
 '       .Bars(6).OfficeXpStyle = True




  ' End With
  ' Set itmX = Nothing
  ' Set barX = Nothing
End Sub

'
Private Sub mnuStockItem_Click()
Load Form2
Form2.Show vbModal

'Load Form1
' frmControlMain.WBrow.Visible = True
'    frmControlMain.DataGrid1.Visible = False
''SQL = "SELECT Product.Name," _
'    & " prodOut.Quantity as [QntyOut]," _
'    & " prodIn.Quantity as [QntyIN]," _
'    & " prodIn.TotalPurchase," _
'    & " prodOut.Total AS TotalSales" _
'    & " FROM prodOut RIGHT JOIN " _
'    & " (prodIn RIGHT JOIN Product ON prodIn.ProductID = Product.ProductID) ON prodOut.ProductID = Product.ProductID"
'
'SQL = "Select * From dprodstata"
'
'Call frmControlMain.CreateSubPage(SQL, "Stock Analysis")
End Sub

Private Sub mnuViewall_Click()
SQL = "Select username1 from users"

frmDelUser.Caption = "View Users"
frmDelUser.Command1.Visible = False
Load frmDelUser
frmDelUser.Show vbModal

End Sub

Private Sub tdrr()
counter = 0
Set Rs1 = New Recordset
Set RS2 = New Recordset
Dim rs3 As ADODB.Recordset
Set rs3 = New Recordset
Dim da As String
da = "01/01/20" & Right(Date, 2)
Dim dattum As Date
Dim sql1 As String
Dim SQL As String
Dim pro As Long
Dim zap As Integer
Dim zalo As Long
Dim sql2 As String
Dim z As Integer
dattum = da
zap = 1
z = 1
myConection.Execute ("delete from tdr")
Do While Not Right(dattum, 2) > Right(Date, 2)
counter = z / 364 * 100
z = z + 1
Label5.Caption = (Round(counter, 2))
 If counter > 100 Then counter = 100
frmMAIN.ProgressBar.Value = counter
'MsgBox (counter)

'frmMAIN.Label5.Caption = Str(counter) & " %"
SQL = "select datum,stdok,sum(nabcena*kol) as zne from nabasif where datum=#" & dattum & "# group by stdok,datum order by datum"
Set Rs1 = myConection.Execute(SQL)
'MsgBox (sql)

'Set rs3 = myConection.Execute("select * from tdr")
 sql2 = "select * from tdr"
    If ConnectRS(myConection, rs3, sql2) = False Then
        
    End If
    
If Not Rs1.EOF Then
rs3.AddNew
rs3.Fields("datum") = dattum
rs3.Fields("opis") = Rs1.Fields("stdok")
rs3.Fields("nabava") = Round(Rs1.Fields("zne"), 2)
zalo = zalo + Round(Rs1.Fields("zne"))
rs3.Fields("zaloga") = zalo
rs3.Fields("zap") = zap
rs3.Update
zap = zap + 1
End If
sql1 = "select sifra,znesek from racusif where datum=#" & dattum & "#"
Set RS2 = myConection.Execute(sql1)
pro = 0
If Not RS2.EOF Then
RS2.MoveFirst
Do While Not RS2.EOF
If Not IsNull(RS2.Fields("sifra")) Then
pro = pro + Round(RS2.Fields("znesek") / (1 + (Val(Getnazi("select madapd from mada where madasifr=" & RS2.Fields("sifra"))) / 100)), 2)
Else
pro = pro + Round(RS2.Fields("znesek"), 2)

End If
RS2.MoveNext
Loop
End If
If pro > 0 Then
rs3.AddNew
rs3.Fields("datum") = dattum
rs3.Fields("prodaja") = pro
zalo = Round(zalo, 2) - pro
rs3.Fields("zaloga") = zalo
rs3.Fields("opis") = "REKAPITULACIJA"
rs3.Fields("zap") = zap
rs3.Update
zap = zap + 1
End If
dattum = dattum + 1
Loop

End Sub
Private Sub fresh()
'cst = Index
    frmControlMain.WBrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
    Select Case cst
    
    Case 0
         SQL = "Select * from partner "
        CatalogueName = "Customer"
    Case 1
       ' Sql = "Select * from Supplier WHERE SuppliersID<>'CASH'"
        SQL = "Select * from partner "
        CatalogueName = "Supplier"
    Case 2
     SQL = "Select madasifr,madanazi,madampcd,madazalo from mada order by madagrup,madanazi"
     CatalogueName = "Category"
    Case 3
        SQL = "Select * from grupa order by sifra"
        CatalogueName = "Location"
    Case 4
'        SQL = "Select * from PurchaseOrderHeader"
        SQL = "Select stdok,min(datum) as datum,sum(nabcena*kol) as nabcena, min(sifrapart) as sifrapart from nabasif group by stdok"
        CatalogueName = "Purchase Order"
    Case 5
'        SQL = "Select * from PurchaseReturnHeader"
        If dejpre = 1 Then
        SQL = "Select datum,  sum(znesek) as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by datum order by datum"
        Else
        SQL = "Select datum,  sum(znesek) as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif group by datum order by datum"
        End If
        CatalogueName = "Purchase Return"
    Case 6
        If dejpre = 1 Then
        SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from storno where [datum] between #" & dod & "# AND #" & ddo & "# and stw<>'A' group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from storno where stw<>'A' group by st order by st"
         End If
        CatalogueName = "Purchase Registry"
    Case 7
     If dejpre = 1 Then
        SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from racusif group by st order by st"
         End If
        CatalogueName = "Sales Return"
    Case 8
    Call tdrr
      If dejpre = 1 Then
        SQL = "Select * from tdr where [datum] between #" & dod & "# AND #" & ddo & "#"
        Else
        SQL = "Select * from tdr"
        End If
        CatalogueName = "Sales Registry"
    End Select
    

Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "" Then
If Rs1.State = 1 Then Rs1.Close

'Rs1.Open SQL, myConection, adOpenStatic, adLockOptimistic

Set Rs1 = DCON.Execute(SQL)

If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
   
    
End If
End If
Set Rs1 = Nothing


End Sub
Private Sub zaloga()

counter = 0
Set Rs1 = New Recordset
Set RS2 = New Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim rs55 As ADODB.Recordset
Set rs3 = New Recordset
Dim da As String
Dim sql5 As String
Dim sql1 As String
  Dim sql55 As String
Dim SQL As String
Dim pro As Long
Dim aa As Double
Dim zalo As Long
Dim sql2 As String
Dim z As Integer
Dim opa As String

Dim pr As Double
Dim na As Double
myConection.Execute "Update mada set madazalo=0 "

z = 1
sql2 = "select * from mada"

   
'If rs3.State = 1 Then rs3.Close
'If ConnectRS(DCON, rs3, sql2) = False Then
      
   

 '  End If
  If rs3.State = 1 Then rs3.Close
   
rs3.Open sql2, myConection, adOpenStatic, adLockOptimistic


rs3.MoveFirst
Do While Not rs3.EOF
    na = 0
    pr = 0
    counter = z / rs3.RecordCount * 100
    z = z + 1
     If counter > 100 Then counter = 100
    frmMAIN.ProgressBar.Value = counter
    frmMAIN.Label5.Caption = Str(counter) & " %"
    SQL = "select embalaza,kol from nabasif where sifra=" & Round(rs3.Fields("madasifr"), 0)
    Set Rs1 = myConection.Execute(SQL)
    If Not Rs1.EOF Then
    Rs1.MoveFirst
    End If
    Do While Not Rs1.EOF
    na = na + (Round(Rs1.Fields("embalaza"), 2) * Round(Rs1.Fields("kol"), 2))
    Rs1.MoveNext
    Loop

    
     sql55 = "select * from mada where madasifr=" & Round(rs3.Fields("madasifr"), 0)
    If ConnectRS(myConection, rs55, sql55) = False Then
            
        End If
        
    
    rs55.MoveFirst
    rs55.Fields("madazalo") = Round(na, 2)
    rs55.Fields("madasest") = "N"
    rs55.Update
  
  Set rs55 = Nothing
  
    
    
    'myConection.Execute "Update mada set madazalo=" & Round(na, 2) & " where madasifr=" & rs3.Fields("madasifr")
    'MsgBox ("nabava " & na)
    opa = Getnazi("select madaenme from mada where madasifr=" & Round(rs3.Fields("madasifr"), 0))
    If opa = "KOM" Then
    sql1 = "select kol from racusif where sifra=" & Round(rs3.Fields("madasifr"), 0)
    Set RS2 = myConection.Execute(sql1)
    If Not RS2.EOF Then
    
    RS2.MoveFirst
    End If
    Do While Not RS2.EOF()
    pr = pr + RS2.Fields("kol")
    If Not RS2.EOF Then
    RS2.MoveNext
    End If
    Loop
    sql55 = "select * from mada where madasifr=" & Round(rs3.Fields("madasifr"), 0)
    If ConnectRS(myConection, rs55, sql55) = False Then
            
        End If
        
    
    rs55.MoveFirst
    rs55.Fields("madazalo") = rs55.Fields("madazalo") - Round(pr, 2)
    rs55.Fields("madasest") = "N"
    rs55.Update
  
    Set rs55 = Nothing
    
    
    Else
    sql1 = "select kol,doza from racusif where sifra=" & Round(rs3.Fields("madasifr"), 0)
    Set RS2 = myConection.Execute(sql1)
    If Not RS2.EOF Then
    RS2.MoveFirst
    End If
    Do While Not RS2.EOF
    pr = pr + (RS2.Fields("kol") * Round(RS2.Fields("doza"), 2))
    RS2.MoveNext
    Loop
    sql55 = "select * from mada where madasifr=" & Round(rs3.Fields("madasifr"), 0)
    If ConnectRS(myConection, rs55, sql55) = False Then
            
        End If
        
    
    rs55.MoveFirst
    rs55.Fields("madazalo") = rs55.Fields("madazalo") - Round(pr, 2)
    rs55.Fields("madasest") = "N"
    rs55.Update
  
    Set rs55 = Nothing
    
    
    
    
    
    End If
    
  ' MsgBox (rs3.Fields("madasifr") & "     " & na & "     " & pr)
    
    
    'myConection.Execute "Update mada set madazalo=madazalo-" & Round(pr, 2) & " where madasifr=" & rs3.Fields("madasifr")
    
    'MsgBox (rs3.Fields("madasifr") & " " & na & " " & pr)
  
    
    'aa = Round(na - pr, 2)
    
    'myConection.Execute "Update mada set madazalo=" & Round(na - pr, 2) & " where madasifr=" & rs3.Fields("madasifr")
    'myConection.Execute "Update mada set madasest='N' where madasifr=" & rs3.Fields("madasifr")
    
    sql5 = "select * from sestavi where sifra=" & Round(rs3.Fields("madasifr"), 0)
    Set rs4 = myConection.Execute(sql5)
    If Not rs4.EOF Then
    myConection.Execute "Update mada set madazalo=0 where madasifr=" & Round(rs3.Fields("madasifr"), 0)
    myConection.Execute "Update mada set madasest='D' where madasifr=" & Round(rs3.Fields("madasifr"), 0)
    
    End If
    
    'If aa <> 0 Then
    'MsgBox (Round(na - pr, 2))
    'End If
    'End If
    
    rs3.MoveNext
Loop
    
    sql2 = "select * from sestavi"
     If ConnectRS(myConection, rs3, sql2) = False Then
            
        End If
        rs3.MoveFirst
        Dim ss As Integer
        Dim kk As Double
        Dim ssa As Integer
        Dim ops As Double
    Do While Not rs3.EOF
    ss = rs3.Fields("sifras")
    ssa = rs3.Fields("sifra")
    opa = Getnazi("select madaenme from mada where madasifr=" & ss)
    ops = rs3.Fields("kol")
    SQL = "select sum(kol) as kkoo from racusif where sifra=" & ssa
       If RS.State = 1 Then RS.Close
       
    RS.Open SQL, myConection, adOpenStatic, adLockOptimistic
    If IsNull(RS.Fields("kkoo")) Then
    kk = 0
    Else
    kk = RS.Fields("kkoo") * ops
    End If
    'sql = "select * from mada where madasifr=" & ss
    '   If Rs.State = 1 Then Rs.Close
       
    'Rs.Open sql, myConection, adOpenStatic, adLockOptimistic
    myConection.Execute "Update mada set madazalo=madazalo-" & kk & " where madasifr=" & ss
    'Rs.Fields("madazalo") = Rs.Fields("madazalo") - kk
    '    Rs.Update
    'MsgBox (Rs.Fields("madazalo"))
    'MsgBox (ss & " " & kk)
    rs3.MoveNext
Loop
'myConection.Close

End Sub

Private Sub mnuvoz_Click()
Dim aa As String
Dim dcon1 As ADODB.Connection
 CdlgEx1.CancelError = False
                CdlgEx1.Filter = "Access files|*.mdb|All files|*.*"
                CdlgEx1.ShowOpen
                aa = CdlgEx1.FileName
Dim sCNSTR As String

Set dcon1 = New ADODB.Connection

dcon1.CursorLocation = adUseClient

    sCNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aa + ";"
    dcon1.Open sCNSTR
Set Rs1 = New Recordset

SQL = "select max(st) as st from racusif"
Set Rs1 = dcon1.Execute(SQL)
'If Rs1.RecordCount <= 0 Then
'MsgBox (Rs1.Fields("st"))
'End If
Dim sql1 As String
sql1 = "select max(st) as stma from racusif"
Set RS = myConection.Execute(sql1)
If RS.Fields("stma") - Rs1.Fields("st") > 0 Then
If MsgBox("Ali želiš uvoziti podatke???", vbOKCancel) = vbOK Then
myConection.Execute ("delete from racusif where st<=" & Rs1.Fields("st"))

SQL = "select * from racusif"
Set Rs1 = dcon1.Execute(SQL)

If RS.State = 1 Then RS.Close

RS.Open "select * from racusif ", myConection, adOpenStatic, adLockOptimistic

Rs1.MoveFirst
Do While Not Rs1.EOF
RS.AddNew
RS.Fields("st") = Rs1.Fields("st")
RS.Fields("datum") = Rs1.Fields("datum")
RS.Fields("oseba") = Rs1.Fields("oseba")
RS.Fields("kol") = Rs1.Fields("kol")
RS.Fields("znesek") = Rs1.Fields("znesek")
RS.Fields("doza") = Rs1.Fields("doza")
RS.Fields("placilo") = Rs1.Fields("placilo")
RS.Fields("vst") = Rs1.Fields("vst")
RS.Fields("ura") = Rs1.Fields("ura")
RS.Fields("sifra") = Rs1.Fields("sifra")

RS.Update

Rs1.MoveNext
Loop
End If
End If
End Sub

Private Sub Timer2_Timer()

On Error Resume Next
'Dim FrameExpandHeight As Integer: FrameExpandHeight = 0 'Height for menu to expand to
    Dim i As Integer
If dex = 1 Then
    'Adjust top of each menu
    For i = 1 To FrameXPMenu.Count - 1
        FrameXPMenu(i).Top = FrameXPMenu(i - 1).Top + FrameXPMenu(i - 1).Height + 120
    Next
    
    'Animates a menu if one is resizing
    If doresizex = True Then
        
            MenuHeader(framex).Picture = imgDown.Picture 'Change headder pic
            FrameXPMenu(framex).Height = FrameXPMenu(framex).Height - speed 'Animate
            If FrameXPMenu(framex).Height <= MenuHeader(framex).Height Then doresize = False: FrameXPMenu(framex).Height = MenuHeader(framex).Height 'If we are fully contracted no need to do further resizing
        
    End If
End If
dex = 2
End Sub
