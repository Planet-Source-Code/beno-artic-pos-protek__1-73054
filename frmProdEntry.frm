VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmProdEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artikli"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2880
      TabIndex        =   57
      Text            =   "0"
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   7680
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   45
      Top             =   0
      Width           =   375
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   375
      Left            =   7680
      TabIndex        =   44
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "B"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
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
      MICON           =   "frmProdEntry.frx":000C
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
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "touch screen"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
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
      MICON           =   "frmProdEntry.frx":0028
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
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   19
      Top             =   0
      Width           =   3135
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6240
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Izpolni vsa polja s *"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   180
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmProdEntry.frx":0044
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artikli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   600
         TabIndex        =   26
         Top             =   0
         Width           =   870
      End
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   5745
      Left            =   0
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   817
      TabIndex        =   17
      Top             =   600
      Width           =   12255
      Begin LVbuttons.LaVolpeButton LaVolpeButton12 
         Height          =   255
         Left            =   3720
         TabIndex        =   68
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Enable"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":090E
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
         Height          =   375
         Left            =   5640
         TabIndex        =   67
         Top             =   4560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Dodatni nazivi"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":092A
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
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   5160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Slika"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   8454143
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmProdEntry.frx":0946
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
         Height          =   615
         Left            =   9840
         TabIndex        =   65
         Top             =   4920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Shrani"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   12648384
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmProdEntry.frx":0962
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
         Left            =   11040
         TabIndex        =   64
         Top             =   4920
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Prekini"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "frmProdEntry.frx":097E
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
      Begin VB.TextBox naziv1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Customer Name"
         Top             =   960
         Width           =   5055
      End
      Begin ProsVent.UserControl2 UserControl21 
         Height          =   975
         Left            =   2520
         Top             =   1800
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1720
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton6 
         Height          =   495
         Left            =   3840
         TabIndex        =   61
         Top             =   5040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Preracun kartice"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":099A
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
      Begin VB.TextBox fakx 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   55
         Top             =   2280
         Width           =   675
      End
      Begin LVbuttons.LaVolpeButton Pros 
         Height          =   375
         Left            =   6840
         TabIndex        =   54
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "DID"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":09B6
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
      Begin ProsVent.UserControl1 tip_a 
         Height          =   375
         Left            =   1680
         TabIndex        =   53
         Top             =   4320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Locked          =   0   'False
         polje           =   "tip"
         ssql            =   "select * from tip_art"
         TextLocked      =   0   'False
      End
      Begin ProsVent.UserControl1 eme 
         Height          =   375
         Left            =   1680
         TabIndex        =   52
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Locked          =   0   'False
         polje           =   "em"
         ssql            =   "select * from em"
         TextLocked      =   0   'False
      End
      Begin ProsVent.UserControl1 grupa 
         Height          =   375
         Left            =   1680
         TabIndex        =   51
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Locked          =   0   'False
         polje           =   "grupa"
         ssql            =   "select * from grupa"
         TextLocked      =   0   'False
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   4800
         TabIndex        =   48
         Text            =   "1"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3960
         TabIndex        =   47
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3360
         TabIndex        =   46
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   495
         Left            =   3840
         TabIndex        =   43
         Top             =   4440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NORMATIV FIKSNI"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":09D2
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
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmProdEntry.frx":09EE
         Left            =   9600
         List            =   "frmProdEntry.frx":09FB
         TabIndex        =   41
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Customer Name"
         Top             =   1440
         Width           =   5055
      End
      Begin LVbuttons.LaVolpeButton kart 
         Height          =   495
         Left            =   1680
         TabIndex        =   36
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "KARTICA ARTIKLA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":0A17
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
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Text            =   "0"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Text            =   "1"
         Top             =   3360
         Width           =   615
      End
      Begin VB.ListBox List87 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1740
         ItemData        =   "frmProdEntry.frx":0A33
         Left            =   5640
         List            =   "frmProdEntry.frx":0A35
         TabIndex        =   11
         Top             =   2760
         Width           =   6360
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmProdEntry.frx":0A37
         Left            =   9600
         List            =   "frmProdEntry.frx":0A44
         TabIndex        =   7
         Text            =   "20"
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmProdEntry.frx":0A57
         Left            =   9600
         List            =   "frmProdEntry.frx":0A64
         TabIndex        =   9
         Text            =   "20"
         Top             =   1440
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   5760
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProdEntry.frx":0A75
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDeleteProdPack 
         Caption         =   "&Briši"
         Height          =   315
         Left            =   10080
         TabIndex        =   25
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdEditProdPack 
         Caption         =   "&Uredi"
         Height          =   315
         Left            =   9240
         TabIndex        =   24
         Top             =   2280
         Width           =   825
      End
      Begin VB.CommandButton cmdAddProdPack 
         Caption         =   "&Dodaj"
         Height          =   315
         Left            =   8400
         TabIndex        =   23
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox ean_C 
         Height          =   315
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00F8F8F8&
         Caption         =   "Akti&ven"
         Height          =   255
         Left            =   390
         TabIndex        =   16
         Top             =   5280
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox nabc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   8
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox mpcc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1800
         Width           =   1635
      End
      Begin VB.TextBox naziv 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "Customer Name"
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtProdID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton4 
         Height          =   495
         Left            =   4680
         TabIndex        =   49
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NORMA VIŠINA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":100F
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
         Height          =   495
         Left            =   3840
         TabIndex        =   50
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NORMA ŠIRINA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "frmProdEntry.frx":102B
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
      Begin VB.Image slikka 
         Height          =   1095
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dodatni opis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   63
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H. OD      DO         H.CENA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   60
         Top             =   3120
         Width           =   1950
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   59
         Top             =   3960
         Width           =   270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faktor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6960
         TabIndex        =   56
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&DODAJ FORMO VŠD"
         Height          =   195
         Index           =   3
         Left            =   7440
         TabIndex        =   42
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sestavnica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   40
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Šifra dobavitelja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip_art"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Min. Zaloga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   3960
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " * Doza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label zaloga 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SESTAVLJEN ARTIKEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   33
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ZALOGA "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   4800
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " * GRUPA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Izstopni DDV"
         Height          =   195
         Index           =   2
         Left            =   7380
         TabIndex        =   28
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Vstopni DDV"
         Height          =   195
         Index           =   1
         Left            =   7380
         TabIndex        =   29
         Top             =   720
         Width           =   870
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00808080&
         Height          =   1935
         Left            =   5580
         Top             =   2640
         Width           =   6525
      End
      Begin VB.Label emm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&EM:"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   2880
         Width           =   270
      End
      Begin VB.Label lblRM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1320
         TabIndex        =   21
         Top             =   4050
         Width           =   45
      End
      Begin VB.Label lblRC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1320
         TabIndex        =   20
         Top             =   3870
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nabavna cena"
         Height          =   195
         Index           =   0
         Left            =   7380
         TabIndex        =   14
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPC :"
         Height          =   195
         Left            =   7380
         TabIndex        =   15
         Top             =   1860
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Naziv"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Šifra Izdelka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton7 
      Height          =   495
      Left            =   9720
      TabIndex        =   62
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Preracun kartice vseh artiklov"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
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
      MICON           =   "frmProdEntry.frx":1047
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
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   58
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Min. Zaloga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   37
      Top             =   4560
      Width           =   1005
   End
End
Attribute VB_Name = "frmProdEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim curProd As tProd
Dim newProd As tProd

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean
Public xzalogg As Double

Public Function ShowAdd(Optional ByVal sProdDescription As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newProd.ProdDescription = sProdDescription
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function


Public Function ShowAddRetID(ByRef lProdID As Long, Optional ByVal sProdDescription As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    newProd.ProdDescription = sProdDescription
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAddRetID = mShowAdd
    lProdID = newProd.ProdID
    
End Function

Public Function ShowEdit(ByVal lProdID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curProd.ProdID = lProdID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function






'Private Sub cmbCat_GotFocus()
'    Dim tmpS As String
    
'    tmpS = cmbcat.text
    
    'load Category List
'    modRSCat.FillCatToCMB cmbcat
    
'    cmbcat.text = tmpS
'End Sub

Private Sub cmbPack_GotFocus()
    
    Dim tmpS As String
    
  
    
End Sub

Private Sub cmbPack_LostFocus()
        
  
    
End Sub

Private Sub cmbProdPack_GotFocus()

    Dim tmpS As String
    
    'tmpS = cmbProdPack.text
    
    'load package list for other package
   ' modRSPack.FillPackToCMB cmbProdPack

   ' cmbProdPack.text = tmpS
    
End Sub

Private Sub cmdDeleteProdPack_Click()
Dim rst As ADODB.Recordset
Dim xxxsql As String
Dim aa As String
If Me.List87.Text <> "" Then
aa = Left(Me.List87.Text, 13)
Set rst = myConection.Execute("delete * from sestavi where sifra=" & Val(Me.txtProdID.Text) & " and sifras=" & Val(aa))
End If
     Me.List87.clear
     xxxsql = "sELECT sifras,kol from sestavi where sifra=" & Val(Me.txtProdID.Text)
     Filllist1 List87, xxxsql
     
End Sub

Private Sub cmdEditProdPack_Click()
Dim rst As ADODB.Recordset
Dim xxxsql As String
Dim aa As String
If Me.List87.Text <> "" Then
aa = Left(Me.List87.Text, 13)
'Set rst = myConection.Execute("delete * from sestavi where sifra=" & Val(Me.txtProdID.text) & " and sifras=" & Val(aa))
uredi = 1
izbrko = Val(aa)
siff = Val(Me.txtProdID.Text)
kosovnica.Show vbModal
 End If
     
End Sub

Private Sub cmdSave_Click()


End Sub

Private Sub cmdCancel_click()
    
    
End Sub

Private Sub Combo1_Change()
'Me.Text4.Text = Me.Combo1.Text
End Sub

Private Sub Combo2_Change()
'Me.Text3.Text = Me.Combo2.Text

End Sub

Private Sub Command1_Click()
'MsgBox (tip_a.BoundDatax)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'MsgBox (tip_a.BoundDatax)

End Sub
Private Sub Form_Activate()

    Me.AutoRedraw = False
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
    
    
    
    
    Select Case mFormState
        Case "add"
                        
            'set form caption
            Me.Caption = "Dodaj artikel"
            
            'generate new Prod ID
            txtProdID.Text = Val(Getnazi("SELECT MAX(val(MADASIFR)) AS CC FROM MADA")) + 1

            
        Case "edit"
        
            Dim vPack As tPack
            Dim vCat As tCat
    
            'set form caption
            Me.Caption = "Urejanje artikla"
            
            'get product info
            If GetProdByID(curProd.ProdID, curProd) = False Then
                
            End If
            
         Dim xxxsql As String
       
            'set form fields
            With curProd
                txtProdID.Text = Trim(str(Val(MODIFYID)))
                
               Dim asi, sifx As String
               asi = (txtProdID.Text)
               If Getnazi("select dobavit_id from mada where madasifr='" & asi & "'") <> "" Then
               sifx = Mid(Trim(Getnazi("select dobavit_id from mada where madasifr='" & asi & "'")), 2, Len(Trim(Getnazi("select dobavit_id from mada where madasifr='" & asi & "'"))) - 2)
                Text3.Text = Replace(sifx, "/,/", ",")
                End If
                tip_a.BoundDatax = Getnazi("select tip_art from mada where madasifr='" & asi & "'")
                
                chkActive.Value = IIf(.Active = True, vbChecked, vbUnchecked)
             
            
                naziv.Text = Getnazi("select madanazi from mada where madasifr='" & asi & "'")
                naziv1.Text = Getnazi("select madanaz1 from mada where madasifr='" & asi & "'")
                ean_C.Text = Getnazi("select madaean from mada where madasifr='" & asi & "'")
                 eme.BoundDatax = Getnazi("select madaenme from mada where madasifr='" & asi & "'")
                grupa.BoundDatax = Getnazi("select grupa from grupa where sifra=" & Getnazi("select madagrup from mada where madasifr='" & Trim(txtProdID.Text) & "'"))
                Text1.Text = Getnazi("select madadoza from mada where madasifr='" & asi & "'")
                Text2.Text = Getnazi("select madaminz from mada where madasifr='" & asi & "'")
                Combo2.Text = Getnazi("select madapdv from mada where madasifr='" & asi & "'")
                nabc.Text = Getnazi("select madanabc from mada where madasifr='" & asi & "'")
                Combo1.Text = Getnazi("select madapd from mada where madasifr='" & asi & "'")
               mpcc.Text = Getnazi("select madampcd from mada where madasifr='" & asi & "'")
                Combo4.Text = Getnazi("select postava from mada where madasifr='" & asi & "'")
                If Getnazi("select kontrola from mada where madasifr='" & asi & "'") <> "" Then
                Picture1.BackColor = Getnazi("select kontrola from mada where madasifr='" & asi & "'")
                Else
                Picture1.BackColor = 0
                End If
       Text4.Text = Getnazi("select odure from mada where madasifr='" & asi & "'")
       Text5.Text = Getnazi("select doure from mada where madasifr='" & asi & "'")
       Text6.Text = Getnazi("select happy from mada where madasifr='" & asi & "'")
       Text7.Text = Getnazi("select MADAZACS from mada where madasifr='" & asi & "'")
       If Getnazi("select madaemba from mada where madasifr='" & asi & "'") = "" Then
       fakx.Text = "1"
       Else
       fakx.Text = FormatNumber(Getnazi("select madaemba from mada where madasifr='" & asi & "'"), 4)
       End If
            End With
        
        xxxsql = "sELECT sifras,kol from sestavi where sifra=" & Val(Me.txtProdID.Text)
         If rs.State = 1 Then rs.Close
       rs.Open xxxsql, myConection
       If rs.EOF Then
       Me.zaloga.Caption = Getnazi("select madazalo from mada where madasifr='" & Me.txtProdID.Text & "'") & " " & Getnazi("select madaenme from mada where madasifr='" & Me.txtProdID.Text & "'")
       xzalogg = 0
       If Getnazi("select madazalo from mada where madasifr='" & Me.txtProdID.Text & "'") <> "" Then
       xzalogg = Getnazi("select madazalo from mada where madasifr='" & Me.txtProdID.Text & "'")
       End If
       End If
        Filllist1 List87, xxxsql
        
    End Select
    If FileExist(App.path & "\slike\ma_" & Me.txtProdID.Text & ".jpg") Then
    slikka.Picture = LoadPicture(App.path & "\slike\ma_" & Me.txtProdID.Text & ".jpg")
    End If
   If obstaja("dodatni") Then
   Else
    myConection.Execute ("select sifra,naziv into dodatni from trenutna ")
     myConection.Execute ("delete from dodatni")
   End If
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
    Me.AutoRedraw = True
    
End Sub

Private Function FileExist(FileName As String) As Boolean

  On Error GoTo FileDoesNotExist
  
  Call FileLen(FileName)
  FileExist = True
  Exit Function
  
FileDoesNotExist:
  FileExist = False
  
End Function
Private Sub Form_Load()
'ReSizeForm Me
    isOn = False
 '   PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0
   'Call CMB1("grupa", "grupa", cmdpro)
    
    'set list column
    
    
End Sub

Private Sub cmbProdPack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
      
    End If
End Sub

Private Sub cmdAddProdPack_Click()

 Dim rst As ADODB.Recordset
Dim xxxsql As String
Dim aa As String
siff = Val(Me.txtProdID.Text)
kosovnica.Show vbModal

    
End Sub

Private Sub listProdPack_DblClick()
    
End Sub

Private Sub listProdPack_RequestEdit(Row As Long, Col As Long, Cancel As Boolean)
    
    Select Case Col
        Case 0
          
    End Select
    
End Sub

Private Sub listProdPack_RequestUpdate(Row As Long, Col As Long, NewValue As String, Cancel As Boolean)
    
    Dim vPack As tPack
    
    
    Select Case Col
        Case 0 'pack
  
            'If IsEmpty(cmbProdPack.text) = True Then
            '    Exit Sub
            'End If
            
            'check pack duplication
            
            
            
            'add pack if new
            'If modRSPack.AddPack(cmbProdPack.text) = True Then
            '    NewValue = cmbProdPack.text
            'Else
               ' WriteErrorLog Me.Name, "listProdPack_RequestUpdate", "Falied on: 'modRSPack.AddPack(cmbProdPack.Text) = True'"
            'End If
            
            'add id
            If modRSPack.GetPackByTitle(NewValue, vPack) = True Then
              '  listProdPack.CellText(Row, 1) = vPack.PackID
            Else
              '  WriteErrorLog Me.Name, "listProdPack_RequestUpdate", "Falied on: 'modRSPack.GetPackByTitle(NewValue, vPack) = True'"
            End If
            
        
        Case 2 'qty
            If Not ((NewValue) > 0) Then
                MsgBox "Vnesi pravo kolièino.", vbExclamation
                Cancel = True
                Exit Sub
            End If
            
            'generate supplier price
           
        Case 3 'supplier price
            If Not ((NewValue) > 0) Then
                MsgBox "Vnesi pravo nab. ceno.", vbExclamation
                Cancel = True
            End If
            
        Case 4 'srp
            If Not ((NewValue) > 0) Then
                MsgBox "Vnesi pravo MPC.", vbExclamation
                Cancel = True
            End If
            
    End Select
    
End Sub







Private Sub SaveAdd()
    
    Dim tmpProd As tProd
    Dim vPack As tPack
    Dim vCat As tCat
    
    
    
    'validate
    
    'check description
    
    'check description duplication
    
    
    'check code duplication
    
    'check package
   
    'set package
   
    'check Category
    
    'set Category
    
    
    'check prices
    
    
    
    'check other packages (last item)
    
    
    'set new product
    With newProd
        .ProdID = (txtProdID)

        '.ProdCode = Trim(txtProdCode.text)
        '.ProdDescription = Trim(txtProdDescription.text)
    
        '.FK_PackID = Text1.text
        '.FK_CatID = vCat.CatID
            
        .BegInvStock = 0
        
        '.SupPrice = FormatNumber((txtSupPrice.text), 2)
        '.SRPrice = FormatNumber((txtSRPrice.text), 2)
        
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        '.RC = Combo1.text
        '.RM = Combo2.text
        '.RCU = Combo3.text
        '   .RMU = Text2.text
    End With
    
    'save
     If modRSProd.AddProd(newProd) = True Then
        
        'save other packages
        Dim li As Long
        'Dim vProdPack As tProdPack
        
       
        
        'set flag
        mShowAdd = True
        Me.Text3.Text = "/" & Trim(Me.Text3.Text) & "/"
        myConection.Execute ("update mada set tip_art='" & tip_a.BoundDatax & "' where madasifr='" & Me.txtProdID.Text & "'")
myConection.Execute ("update mada set dobavit_id='" & Replace(Trim(Me.Text3.Text), ",", "/,/") & "' where madasifr='" & Me.txtProdID.Text & "'")
        'unload this form
        Unload Me
    Else
      '  WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProd.AddProd(newProd) = True'"
    End If
    
    

   ber = 1


End Sub


Private Sub SaveEdit()
    
    Dim tmpProd As tProd
    Dim vPack As tPack
    Dim vCat As tCat
    
    'validate
    
    'check description
   ' If IsEmpty(txtProdDescription.text) Then
   '     MsgBox "Vnesi naziv.", vbExclamation
      '  HLTxt txtProdDescription
   '     Exit Sub
   ' End If
    
    'If LCase(Trim(curProd.ProdDescription)) <> LCase(Trim(txtProdDescription.text)) Then
    '    'check description duplication
    '    If modRSProd.GetProdByDescription(txtProdDescription.text, tmpProd) = True Then
    '        MsgBox "Naziv že obstaja.", vbExclamation
          '  HLTxt txtProdDescription
    '        Exit Sub
    '    End If
    'End If
    
    'check code duplication
    'If Not IsEmpty(txtProdCode.text) Then
    '    If LCase(Trim(curProd.ProdCode)) <> LCase(Trim(txtProdCode.text)) Then
    '        If modRSProd.GetProdByCode(txtProdCode.text, tmpProd) = True Then
    '            MsgBox "že obstaja koda.", vbExclamation
               ' HLTxt txtProdCode
     '           Exit Sub
     '       End If
     '   End If
    'End If
    
    'check package
    'If cmbPack.ListIndex < 0 Then
     '   If IsEmpty(cmbPack.Text) Then
     '      MsgBox "Vnesi pravo kategorijo.", vbExclamation
     '     cmbPack.SetFocus
     '       Exit Sub
     '   Else
     '       'add new package
     '       modRSPack.AddPack Trim(cmbPack.Text)
     '   End If
  '  End If
    
    'set package
   ' If modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False Then
   '     WriteErrorLog Me.Name, "SaveAdd", "Failed on: modRSPack.GetPackByTitle(cmbPack.Text, vPack) = False'  |  PackTitle: " & cmbPack.Text
   '     Exit Sub
   ' End If
    
    'check Category
    
    
    'check prices
    
    
    
    'check other packages (last item)
    
    
    
    
    
    'set cur product
    With curProd
        .ProdID = (txtProdID)

       ' .ProdCode = Trim(txtProdCode.text)
       ' .ProdDescription = Trim(txtProdDescription.text)
    
        '.FK_PackID = Text1.text
        '.FK_CatID = vCat.CatID
            
        .BegInvStock = 0
        '
        '.SupPrice = FormatNumber((txtSupPrice.text), 2)
        '.SRPrice = FormatNumber((txtSRPrice.text), 2)
        
        .Active = IIf(chkActive.Value = vbChecked, True, False)
        '.RC = Combo1.text
        '.RM = Combo2.text
        '.RCU = Combo3.text
        '.RMU = Text2.text
        
    End With
    
    'save
    
    If modRSProd.EditProd(curProd) = True Then
        
        'delete all other packages
       ' If modRSProdPack.DeleteAllProdPack(curProd.ProdID) = False Then
       '     WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSProdPack.DeleteAllProdPack(curProd.ProdID) = False'"
       ' End If
        
        'save other packages
        Dim li As Long
        'Dim vProdPack As tProdPack
        
        
        'set flag
        mShowEdit = True
        Me.Text3.Text = "/" & Trim(Me.Text3.Text) & "/"
        myConection.Execute ("update mada set tip_art='" & tip_a.BoundDatax & "' where madasifr='" & Me.txtProdID.Text & "'")
myConection.Execute ("update mada set dobavit_id='" & Replace(Trim(Me.Text3.Text), ",", "/,/") & "' where madasifr='" & Me.txtProdID.Text & "'")

        'unload this form
        Unload Me
    Else
       ' WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSProd.EditProd(newProd) = True'"
    End If
     
        ber = 1
        
End Sub



Private Sub kart_Click()
Call skla


 CreateH_Page "Select datum,opis,nabava,prodaja,cen,zaloga,nace as vredzal  from tdr", " Kartica artikla.: " & Getnazi("select madanazi from mada where madasifr='" & Me.txtProdID.Text & "'")



''''''''''''''''''''''''''
'On Error GoTo adder:
End Sub

Private Sub skla()
Dim arti As Integer
Set Rs1 = New Recordset
Set RS2 = New Recordset
Dim rs3 As ADODB.Recordset
Set rs3 = New Recordset
Dim da As String
da = "01/01/20" & Right(Date, 2)
Dim dattum As Date
Dim sql1 As String
Dim SQL As String
Dim pro As Double
Dim zap As Integer
Dim zalo As Double
Dim sql2 As String
Dim z As Integer
dattum = da
zap = 1
z = 1
arti = Me.txtProdID.Text
If obstaja("tdr") Then
 myConection.Execute ("DROP TABLE tdr")
 End If
If rs.State = 1 Then rs.Close

'myConection.Execute ("delete from tdr")
'Do While Not Right(dattum, 2) > Right(Date, 2)
z = z + 1
SQL = "select top 1 datum,(tip_dok+id_dok) as stdok,sum(kol*faktor) as nabava,sum(kol*faktor) as zaloga,sum(kol*faktor) as prodaja,sum(kol*faktor) as zap,sum(kol*faktor) as koli,space(250) as opis,min(faktor) as faktor,min(cena) as cen,min(mpc) as nace,min(poknj) as poknjizen into tdr from nabasif where faktor<>0 and sifra='" & arti & "' group by tip_dok,id_dok,datum,sifra order by datum"
Set Rs1 = myConection.Execute(SQL)
'MsgBox (sql)
myConection.Execute ("delete  from tdr")
SQL = "select datum,(tip_dok+id_dok) as stdok,sum(kol*faktor) as koli,min(faktor) as faktor,min(cena) as cen,min(mpc) as nace,min(poknj) as poknjizen  from nabasif where faktor<>0 and sifra='" & arti & "' group by tip_dok,id_dok,datum,sifra order by datum"
Set Rs1 = myConection.Execute(SQL)
'Set rs3 = myConection.Execute("select * from tdr")
 sql2 = "select * from tdr"
    If ConnectRS(myConection, rs3, sql2) = False Then
        
    End If
    
If Not Rs1.EOF Then
Rs1.MoveFirst
Dim dozz As Double
Dim VREDZ As Double
dozz = Getnazi("select madadoza from mada where madasifr='" & arti & "'")
Do While Not Rs1.EOF
rs3.AddNew
rs3.Fields("datum") = Rs1.Fields("datum")
If Rs1.Fields("poknjizen") = "K" Then
rs3.Fields("opis") = Rs1.Fields("stdok") & " - POKNJIŽEN"
Else
rs3.Fields("opis") = Rs1.Fields("stdok") & " - NEPOKNJIŽEN"
End If

If Rs1.Fields("faktor") > 0 Then
rs3.Fields("nabava") = FormatNumber(Rs1.Fields("koli"), 2)
zalo = zalo + Rs1.Fields("koli")

rs3.Fields("cen") = FormatNumber(Rs1.Fields("cen"), 2)
VREDZ = VREDZ + (rs3.Fields("nabava") * rs3.Fields("cen"))

Else
rs3.Fields("prodaja") = FormatNumber(Rs1.Fields("koli") * dozz * -1, 2)
zalo = zalo + Rs1.Fields("koli") * dozz
rs3.Fields("cen") = FormatNumber(Rs1.Fields("nace"), 2)
VREDZ = VREDZ + (rs3.Fields("prodaja") * (1 / dozz) * rs3.Fields("cen") * -1)

End If
rs3.Fields("nace") = FormatNumber(VREDZ, 2)
rs3.Fields("zaloga") = FormatNumber(zalo, 2)
rs3.Fields("zap") = zap
rs3.Update
Rs1.MoveNext
zap = zap + 1
Loop
End If
'dattum = dattum + 1
'Loop
End Sub



Private Sub LaVolpeButton1_Click()
touch
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub LaVolpeButton10_Click()
trenslika = Trim(Me.txtProdID.Text)
slike.Show vbModal
End Sub

Private Sub LaVolpeButton11_Click()
dodatni_ar = Trim(Me.txtProdID.Text)
DOD_AR = "art"
dodatni.Show vbModal

End Sub

Private Sub LaVolpeButton12_Click()
Me.txtProdID.Enabled = True
End Sub

Private Sub LaVolpeButton2_Click()
If Me.tip_a.BoundDatax = "IZD" Then
tip_dok = "NT"
ma_ured = 0
normati = "NT" & Me.txtProdID.Text
frmblag.Show vbModal
Else
MsgBox ("Normativ lahko kreiraš le izdelku!")
End If
End Sub

Private Sub LaVolpeButton3_Click()
   On Error GoTo NoColorChosen
   With CommonDialog1
      .CancelError = True
      ' Entire dialog box is displayed, including the Define Custom Colors section
      .flags = cdlCCFullOpen
      .ShowColor  ' Launch the Color Dialog
      Me.Picture1.BackColor = .Color  ' Assign selected color to background of Picture1
      Exit Sub
   End With
NoColorChosen:
   ' Get here if user clicks the Cancel button
   MsgBox "NISI SI IZBRAL BARVE!", vbInformation, "Cancelled"
   Exit Sub

End Sub

Private Sub LaVolpeButton4_Click()
'tip_dok = "NT"
'ma_ured = 0
'normati = "NTX" & Me.txtProdID.text
'frmblag.Show vbModal
brow_nt.odpri "select (id_dok) as dokument,stdok from nabasif where tip_dok='NT' and id_dok like 'X" & LTrim(Me.txtProdID.Text) & "%' group by id_dok,stdok", "X" & LTrim(Me.txtProdID.Text)
End Sub

Private Sub LaVolpeButton5_Click()
'tip_dok = "NT"
'ma_ured = 0
'normati = "NTY" & Me.txtProdID.text
'frmblag.Show vbModal
brow_nt.odpri "select (id_dok) as dokument, stdok from nabasif where tip_dok='NT' and id_dok like 'Y" & LTrim(Me.txtProdID.Text) & "%' group by id_dok,stdok", "Y" & LTrim(Me.txtProdID.Text)
End Sub

Private Sub LaVolpeButton6_Click()
Dim ratt As New ADODB.Recordset
'MsgBox (Getnazi("select sifra from nabasif where tip_dok='NA' and sifra='" & Me.txtProdID.Text & "' order by datum"))
If Getnazi("select sifras from sestavi where sifra=" & Val(Me.txtProdID.Text)) <> "" Then
Dim sises As String
Dim rrr As New ADODB.Recordset
'MsgBox ("select sifras from sestavi where sifra=" & Val(Me.txtProdID.Text))
rrr.Open "select sifras from sestavi where sifra=" & Val(Me.txtProdID.Text), myConection, adOpenDynamic, adLockOptimistic
If Not rrr.EOF Then
rrr.MoveFirst
Do While Not rrr.EOF
If sises = "" Then
sises = Trim(str(rrr.Fields("sifras")))
Else
sises = sises & "','" & Trim(str(rrr.Fields("sifras")))
End If
rrr.MoveNext
Loop
'MsgBox ("select datum from nabasif where tip_dok='NA' and sifra in ('" & sises & "') order by datum")
ratt.Open "select datum from nabasif where tip_dok='NA' and sifra in ('" & sises & "') order by datum", myConection, adOpenDynamic, adLockOptimistic
End If
Else
ratt.Open "select datum from nabasif where tip_dok='NA' and sifra='" & Me.txtProdID.Text & "' order by datum", myConection, adOpenDynamic, adLockOptimistic
End If
If Not ratt.EOF Then
ratt.MoveFirst
Xvs = 1
Do While Not ratt.EOF
Xvs = Xvs + 1
ratt.MoveNext
Loop
ratt.MoveFirst


End If
If Not ratt.EOF Then
Dim RSS As New ADODB.Recordset
myConection.Execute ("delete from cenik where sifra='" & Me.txtProdID.Text & "'")
RSS.Open "select * from cenik", myConection, adOpenDynamic, adLockOptimistic
ratt.MoveFirst
'Dim klk, ii As Integer

Yvs = 1
Me.UserControl21.opentime
Me.UserControl21.Visible = True
Do While Not ratt.EOF
DoEvents
RSS.AddNew
RSS.Fields("datum") = ratt.Fields("datum")
RSS.Fields("cena") = Getcena(Me.txtProdID.Text, ratt.Fields("datum"))
RSS.Fields("sifra") = Me.txtProdID.Text
RSS.Update
Dim das As String
  das = Format(RSS.Fields("datum"), "dd.mm.yyyy")
  dod = novast(LTrim(LTrim(str(Month(das)))), 2) & "/" & novast(LTrim(LTrim(str(Day(das)))), 2) & "/" & LTrim(LTrim(str(Year(das))))
'MsgBox ("update nabasif set mpc=" & RSS.Fields("cena") & " where tip_dok='PA' and datum>=#" & dod & "#")
myConection.Execute ("update nabasif set mpc=" & Replace(RSS.Fields("cena"), ",", ".") & " where sifra='" & Me.txtProdID.Text & "' and tip_dok='PA' and datum>=#" & dod & "#")
Yvs = Yvs + 1
ratt.MoveNext
Loop
Me.UserControl21.Visible = False
Me.UserControl21.closetime
End If
MsgBox ("Urejeno")
End Sub

Private Sub poprav_Click()
Dim sql2 As String
 sql2 = "select * from MADA order by madagrup,dobavit_id"
    If ConnectRS(myConection, rs, sql2) = False Then

    End If
    If Not rs.EOF Then
    rs.MoveFirst
    End If
    Dim XX As Integer
    XX = 1
    Do While Not rs.EOF
   ' RS.Fields("madasifr") = LTrim(Str(10000 + XX))
   ' RS.Update
    rs.MoveNext
    XX = XX + 1
    Loop
End Sub

Private Sub LaVolpeButton7_Click()
Dim vsia As New ADODB.Recordset
vsia.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
vsia.MoveFirst
Xvs = 1
Do While Not vsia.EOF
Xvs = Xvs + 1
vsia.MoveNext
Loop
vsia.MoveFirst
Yvs = 1
Me.UserControl21.opentime
Me.UserControl21.Visible = True

Do While Not vsia.EOF
DoEvents

Dim ratt As New ADODB.Recordset
If ratt.State = 1 Then ratt.Close
'MsgBox (Getnazi("select sifra from nabasif where tip_dok='NA' and sifra='" & Me.txtProdID.Text & "' order by datum"))
If Getnazi("select sifras from sestavi where sifra=" & Val(vsia.Fields("madasifr"))) <> "" Then
Dim sises As String
Dim rrr As New ADODB.Recordset
If rrr.State = 1 Then rrr.Close
'MsgBox ("select sifras from sestavi where sifra=" & Val(vsia.fields("madasifr")))
rrr.Open "select sifras from sestavi where sifra=" & Val(vsia.Fields("madasifr")), myConection, adOpenDynamic, adLockOptimistic
If Not rrr.EOF Then
rrr.MoveFirst
Do While Not rrr.EOF
If sises = "" Then
sises = Trim(str(rrr.Fields("sifras")))
Else
sises = sises & "','" & Trim(str(rrr.Fields("sifras")))
End If
rrr.MoveNext
Loop
'MsgBox ("select datum from nabasif where tip_dok='NA' and sifra in ('" & sises & "') order by datum")
ratt.Open "select datum from nabasif where tip_dok='NA' and sifra in ('" & sises & "') order by datum", myConection, adOpenDynamic, adLockOptimistic
End If
Else
ratt.Open "select datum from nabasif where tip_dok='NA' and sifra='" & vsia.Fields("madasifr") & "' order by datum", myConection, adOpenDynamic, adLockOptimistic
End If
If Not ratt.EOF Then

ratt.MoveFirst


End If
If Not ratt.EOF Then

Dim RSS As New ADODB.Recordset
myConection.Execute ("delete from cenik where sifra='" & vsia.Fields("madasifr") & "'")
If RSS.State = 1 Then RSS.Close
RSS.Open "select * from cenik", myConection, adOpenDynamic, adLockOptimistic
ratt.MoveFirst
'Dim klk, ii As Integer


Do While Not ratt.EOF

RSS.AddNew
RSS.Fields("datum") = ratt.Fields("datum")
RSS.Fields("cena") = Getcena(vsia.Fields("madasifr"), ratt.Fields("datum"))
RSS.Fields("sifra") = vsia.Fields("madasifr")
RSS.Update
Dim das As String
  das = Format(RSS.Fields("datum"), "dd.mm.yyyy")
  dod = novast(LTrim(LTrim(str(Month(das)))), 2) & "/" & novast(LTrim(LTrim(str(Day(das)))), 2) & "/" & LTrim(LTrim(str(Year(das))))
'MsgBox ("update nabasif set mpc=" & RSS.Fields("cena") & " where tip_dok='PA' and datum>=#" & dod & "#")
myConection.Execute ("update nabasif set mpc=" & Replace(RSS.Fields("cena"), ",", ".") & " where sifra='" & vsia.Fields("madasifr") & "' and tip_dok='PA' and datum>=#" & dod & "#")
ratt.MoveNext
Loop

End If

Yvs = Yvs + 1

vsia.MoveNext
Loop
Me.UserControl21.Visible = False
Me.UserControl21.closetime

MsgBox ("Urejeno")

End Sub

Private Sub LaVolpeButton8_Click()
Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select
    
    Unload Me
End Sub

Private Sub LaVolpeButton9_Click()
If Me.grupa.BoundDatax = "" Then
MsgBox "Grupa je obvezen podatek!"
Exit Sub
End If
If Me.eme.BoundDatax = "" Then
MsgBox "EM je obvezen podatek!"
Exit Sub
End If
If Me.tip_a.BoundDatax = "" Then
MsgBox "Tip artikla je obvezen podatek!"
Exit Sub
End If
If mpcc.Text = "" Then
mpcc.Text = "0"
End If
If Text4.Text = "" Then
Text4.Text = "0"
End If
If Text5.Text = "" Then
Text5.Text = "0"
End If
If Text6.Text = "" Then
Text6.Text = "0"
End If
If Text7.Text = "" Then
Text7.Text = "0"
End If

If nabc.Text = "" Then
 nabc.Text = "0"
 End If
Dim xzalo As Double
xzalo = 0
If Me.zaloga.Caption <> "" Then
'xzalo = Me.zaloga.Caption
End If
Dim sifr, naz, fakto, dob_id, ean, kat, gru, doz, min, ti_a, KMM, vs_g, iz_g, na_ce, pr_ce, gg, dob_ide, post, barv, xod, xdo, xhap
sifr = Trim(str(Val(txtProdID.Text)))
naz = naziv.Text
dob_id = "/" & Trim(Text3.Text) & "/"
dob_ide = Replace(dob_id, ",", "/,/")
ean = ean_C.Text

gru = Val(Getnazi("select sifra from grupa where grupa='" & grupa.BoundDatax & "'"))
kat = eme.BoundDatax
'MsgBox kat
If Text1.Text = "" Then
Text1.Text = "1"
End If
doz = FormatNumber(Text1.Text, 2)
If Text2.Text = "" Then Text2.Text = "0"
min = FormatNumber(Text2.Text, 2)
ti_a = tip_a.BoundDatax
vs_g = FormatNumber(Combo2.Text, 2)
na_ce = FormatNumber(nabc.Text, 4)
iz_g = FormatNumber(Combo1.Text, 2)
pr_ce = FormatNumber(mpcc.Text, 4)
post = Trim(Combo4.Text)
barv = Me.Picture1.BackColor
xod = Text4.Text
xdo = Text5.Text
KMM = Text7.Text
If fakx.Text = "" Then fakx.Text = "1"
fakto = FormatNumber(fakx.Text, 4)

If Text6.Text = "" Then
Text6.Text = 0
End If
xhap = FormatNumber(Text6.Text, 2)


myConection.Execute ("delete from mada where madasifr='" & sifr & "'")

Dim sfw As String
If rs.State = 1 Then rs.Close
rs.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("madasifr") = sifr
rs.Fields("postava") = post
rs.Fields("madanazi") = naz
rs.Fields("madanaz1") = naziv1.Text
rs.Fields("dobavit_id") = dob_ide
rs.Fields("madaean") = ean
rs.Fields("madaenme") = kat
rs.Fields("madagrup") = gru
rs.Fields("madadoza") = doz
rs.Fields("madaminz") = min
rs.Fields("tip_art") = ti_a
rs.Fields("madapdv") = vs_g
rs.Fields("madapd") = iz_g
rs.Fields("madanabc") = na_ce
rs.Fields("madampcd") = pr_ce
rs.Fields("kontrola") = barv
rs.Fields("odure") = xod
rs.Fields("doure") = xdo
rs.Fields("happy") = xhap
rs.Fields("madazalo") = xzalogg
rs.Fields("madaemba") = fakto
rs.Fields("madaZACS") = KMM
rs.Update
touch

frmMAIN.beno_os (27)
Unload Me

End Sub

Private Sub Pros_Click()
If Me.grupa.BoundDatax = "" Then
MsgBox "Izberi si grupo artikla"
Else
 CreateH_Page "Select madasifr,madanazi,dobavit_id from mada where madagrup=" & Getnazi("select sifra from grupa where grupa='" & Me.grupa.BoundDatax & "'") & " order by dobavit_id desc", " Proste dobaviteljeve šifre "
End If
End Sub

Private Sub Text1_GotFocus()
 If Me.Text1.Text = "" Then
' If Me.Combo3.text = "l" Or Me.Combo3.text = "L" Then
 '   Me.Text1.text = "0.03"
 '   Else
 '   Me.Text1.text = "1"
 '   End If
End If
End Sub

Private Sub txtProdID_DblClick()
Me.txtProdID.Enabled = True
End Sub

Private Sub txtProdID_LostFocus()
If Getnazi("select madasifr from mada where madasirf='" & Trim(Me.txtProdID.Text) & "'") <> "" Then
MsgBox ("Širfa je že vnesena!! Vnesi drugo!")
Me.txtProdID.SetFocus
End If

End Sub
