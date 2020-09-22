VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{A2CA1CF5-D01D-4EF8-BF60-A5862EB99E1A}#1.0#0"; "EYEDRO~1.OCX"
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
   Begin VB.PictureBox Picture1x 
      Align           =   3  'Align Left
      BackColor       =   &H00FF8080&
      FillColor       =   &H00E0E0E0&
      Height          =   8625
      Left            =   0
      ScaleHeight     =   8565
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin LVbuttons.LaVolpeButton Vrzi 
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   8280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
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
         BCOL            =   15790320
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "MDIForm1xs.frx":0000
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   8625
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   8655
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   0
            Left            =   0
            Picture         =   "MDIForm1xs.frx":001C
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   0
            Left            =   2160
            Picture         =   "MDIForm1xs.frx":0266
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgPreferences 
            Height          =   195
            Index           =   0
            Left            =   6960
            Picture         =   "MDIForm1xs.frx":04B0
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgHelp 
            Height          =   195
            Index           =   0
            Left            =   8400
            Picture         =   "MDIForm1xs.frx":06FA
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   1
            Left            =   2640
            Picture         =   "MDIForm1xs.frx":0944
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   2
            Left            =   3120
            Picture         =   "MDIForm1xs.frx":0B8E
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   3
            Left            =   3600
            Picture         =   "MDIForm1xs.frx":0DD8
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   4
            Left            =   4080
            Picture         =   "MDIForm1xs.frx":1022
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   5
            Left            =   5280
            Picture         =   "MDIForm1xs.frx":126C
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   6
            Left            =   5760
            Picture         =   "MDIForm1xs.frx":14B6
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   7
            Left            =   6240
            Picture         =   "MDIForm1xs.frx":1700
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   8
            Left            =   4680
            Picture         =   "MDIForm1xs.frx":194A
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgTree 
            Height          =   195
            Index           =   9
            Left            =   6600
            Picture         =   "MDIForm1xs.frx":1B94
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgPreferences 
            Height          =   195
            Index           =   1
            Left            =   7200
            Picture         =   "MDIForm1xs.frx":1DDE
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgPreferences 
            Height          =   195
            Index           =   2
            Left            =   7560
            Picture         =   "MDIForm1xs.frx":2028
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   1
            Left            =   360
            Picture         =   "MDIForm1xs.frx":2272
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   2
            Left            =   720
            Picture         =   "MDIForm1xs.frx":24BC
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   3
            Left            =   1080
            Picture         =   "MDIForm1xs.frx":2706
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   4
            Left            =   1320
            Picture         =   "MDIForm1xs.frx":2950
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgFile 
            Height          =   195
            Index           =   5
            Left            =   1560
            Picture         =   "MDIForm1xs.frx":2B9A
            Top             =   0
            Width           =   195
         End
         Begin VB.Image imgPreferences 
            Height          =   195
            Index           =   3
            Left            =   7800
            Picture         =   "MDIForm1xs.frx":2DE4
            Top             =   0
            Width           =   195
         End
      End
      Begin EyeDropperTab.EyeDropper EyeDropper1 
         Height          =   2055
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3625
         DefaultItemHeight=   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         VisibleItems    =   0
         DisplayIconsInMenu=   -1  'True
         BackColor       =   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   5
            Left            =   1200
            TabIndex        =   17
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1x"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   4
            Left            =   2040
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1x"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   2
            Left            =   2280
            TabIndex        =   14
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1x"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList1x"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox SmallImages 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   10560
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   4
         Top             =   150
         Width           =   1200
      End
      Begin VB.PictureBox i32x32 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   6240
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   3
         Top             =   0
         Width           =   1200
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
      TabIndex        =   0
      Top             =   0
      Width           =   14700
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
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
            Picture         =   "MDIForm1xs.frx":302E
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
            Picture         =   "MDIForm1xs.frx":3E82
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
            TextSave        =   "3.11.2011"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":45DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":48F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":4D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":AFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":B2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":B750
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":BBA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":11394
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1762E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":17948
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":17C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1DEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1E996
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1EDE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1F23A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1F554
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1F9A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":1FDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":2024A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":2069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":209B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":20CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":20E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":20F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":22486
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":23190
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1x 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":234AA
            Key             =   "category"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":238FC
            Key             =   "font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":23B56
            Key             =   "reset"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":23DB0
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":2400A
            Key             =   "compact"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":24264
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":246B6
            Key             =   "search"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":24B08
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":24C1A
            Key             =   "locate"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":24E74
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":250CE
            Key             =   "main"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25520
            Key             =   "add"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25632
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25744
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25856
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25B70
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25CCA
            Key             =   "component"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1xs.frx":25FE4
            Key             =   "equipment"
         EndProperty
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
      Begin VB.Menu mnuViewall 
         Caption         =   "Pregled / urejanje vseh uporabnikov"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Orodja"
      Begin VB.Menu mnucommandd 
         Caption         =   "COMMANDNO"
      End
      Begin VB.Menu mnuvoz 
         Caption         =   "UVOZ"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu cc 
         Caption         =   "Nastavitve"
         Begin VB.Menu mnunastav 
            Caption         =   "Nastavitve"
         End
         Begin VB.Menu mnuotvoritev 
            Caption         =   "OTVORITEV (FIFO)"
         End
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetMenu Lib "user32" _
                          (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
                          (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
                          (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As _
                          Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked _
                          As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" _
        (ByVal hMenu As Long, ByVal nPos As Long) As Long ':( Missing Scope

Private Const MF_BYPOSITION = &H400&
Private mHandle As Long
Private lRet As Long
Private sHandle As Long

Dim CdlgEx1 As New CdlgEx
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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
Private Sub Form_Initialize()

   Call ManifestWrite

End Sub
Private Sub MenuHeader_Click(Index As Integer)

       
        
           


    'If not currently resizeing allow menu to be resized
    If doresize = False Then
        
        'Works out if the menu needs to expand or contract
        'If FrameXPMenu(Index).Height = MenuHeader(Index).Height Then 'minimised
         '   expand = True
        'Else
        '    expand = False
        'End If
        
        'Tell the timer to do the resizing to what frame
        doresize = True
        frame = Index
        
    End If


End Sub

Public Sub Size(Frm As Form)
    Frm.Width = Me.ScaleWidth
    Frm.Height = Me.ScaleHeight
    EyeDropper1.Top = Picture1x.Top + 10
    EyeDropper1.Height = Frm.Height - (Frm.Height / 60)
    Vrzi.Top = EyeDropper1.Top + EyeDropper1.Height
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim tempSql As String
    frmControlMain.MSHFlexGrid1.Visible = True
    frmControlMain.Wbrow.Visible = False
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
frmControlMain.Wbrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
'tip_dok = Left(Me.PREG(Index).Caption, 2)
If dejpre = 1 Then
        'SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from storno where [datum] between #" & dod & "# AND #" & ddo & "# and stw<>'A' group by st order by st"
          'SQL = "Select tip_dok,id_dok,min(stdok) as stdok,min(datum) as datum,sum(cena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif  where [datum] between #" & dod & "# AND #" & ddo & "# and tip_dok='" & Left(PREG(Index).Caption, 2) & "' group by tip_dok,id_dok order by tip_dok,id_dok"
        Else
          ' SQL = "Select tip_dok,id_dok,min(stdok) as stdok,min(datum) as datum,sum(cena*kol) as nabcena, min(sifrapart) as sifrapart,max(poknj) as poknj from nabasif where tip_dok='" & Left(PREG(Index).Caption, 2) & "' group by tip_dok,id_dok order by tip_dok,id_dok"
   End If
        'CatalogueName = "Purchase Registry"
'Form7.Show
Call GetNewConnection2
Set Rs1 = New Recordset
'If CatalogueName <> "" Then

'Set Rs1 = DCON.Execute(SQL)
'If Rs1.RecordCount <= 0 Then
'    frmControlMain.MSHFlexGrid1.Visible = False
'Else
'    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
    frmControlMain.osv_Click
'End If
'End If
Set Rs1 = Nothing
Set DCON = Nothing

End Sub

Private Sub Label1_Click(Index As Integer)
cst = Index
    frmControlMain.Wbrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
    Select Case Index
    Case 14
         SQL = "Select * from dokm where atribut='POST' "
        CatalogueName = "DOKMI"
    
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
        SQL = "Select datum,  format(sum(znesek),'fixed') as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by datum order by datum"
        Else
        SQL = "Select datum,  format(sum(znesek),'fixed') as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif group by datum order by datum"
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
        SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from racusif group by st order by st"
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

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
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
Private Sub AssignMenuBitmaps(ByRef Frm As Form, ByRef img As Image, ByVal Menu_Position As Integer, ByVal Sub_Menu_Position As Integer)
   mHandle = GetMenu(Frm.hwnd)
   sHandle = GetSubMenu(mHandle, Menu_Position)
   lRet = SetMenuItemBitmaps(sHandle, Sub_Menu_Position, MF_BYPOSITION, img.Picture, img.Picture)
End Sub
'Sub to enumerate menus to be painted
Private Sub PaintMenuBitmaps()
On Error Resume Next
  'For File Menu
  AssignMenuBitmaps Me, imgFile(1), 0, 0 'Exit
  AssignMenuBitmaps Me, imgFile(5), 0, 2 'Exit
  'For Tree Menu
  AssignMenuBitmaps Me, imgTree(0), 1, 0 'Expand
  AssignMenuBitmaps Me, imgTree(1), 1, 1 'Search
  AssignMenuBitmaps Me, imgTree(2), 1, 3 'Add
  AssignMenuBitmaps Me, imgTree(3), 1, 4 'Edit
  AssignMenuBitmaps Me, imgTree(4), 1, 5 'Delete
  AssignMenuBitmaps Me, imgTree(8), 1, 7 'Cut
  AssignMenuBitmaps Me, imgTree(5), 1, 8 'Copy
  AssignMenuBitmaps Me, imgTree(6), 1, 9 'Paste
  AssignMenuBitmaps Me, imgTree(7), 1, 11 'Refresh
  AssignMenuBitmaps Me, imgTree(9), 1, 13 'Search Database
  'For Options - Preferences Menu
  AssignMenuBitmaps Me, imgPreferences(0), 2, 0 'Font Settings
  AssignMenuBitmaps Me, imgPreferences(1), 2, 2 'Compact Database
  AssignMenuBitmaps Me, imgPreferences(3), 2, 3 'Repair Database
  AssignMenuBitmaps Me, imgPreferences(2), 2, 5 'Reset Node Color
  'For Help Menu
  AssignMenuBitmaps Me, imgHelp(0), 3, 0 'Help
End Sub
Private Sub lblTask_Click()
    Load frmControlMain
    frmControlMain.MSHFlexGrid1.Visible = False
    frmControlMain.Wbrow.Visible = True
    
    
    Dim SqLargs As String
    SqLargs = "SELECT madasifr,madanazi,madazalo,madampcd From mada WHERE ((madazalo)<=0) Order by madazalo DESC"
   ' Call frmControlMain.CreateStartPage(SqLargs)
End Sub

Private Sub MDIForm_Resize()
    Call Size(frmControlMain)
   
    
End Sub
Private Sub mnuCashBook_Click()
    Form1.Show
End Sub
Private Sub mnuDaybook_Click()
    GetNewConnection2
   Call frmControlMain.CreateDataPage("Select  TOP 2 * From V1", "Day book")
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
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
Private Sub mnucommandd_click()
Dim sqq As String
sqq = InputBox("", "Vnesi sql", "")
myConection.Execute (sqq)
MsgBox "Konèano"
End Sub

Private Sub mnuChangePassword_Click()
Load frmChange
frmChange.Show vbModal

End Sub
Private Sub mnunastav_Click()
Load nastavitve
nastavitve.Show vbModal

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
          frmControlMain.MSHFlexGrid1.Col = UREJAJ
            SQL = "Delete From partner Where sifra=" & frmControlMain.MSHFlexGrid1.Text
            End If
        
       
  Case "Supplier"
        
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
          frmControlMain.MSHFlexGrid1.Col = UREJAJ
            SQL = "Delete From partner Where sifra=" & frmControlMain.MSHFlexGrid1.Text
         End If
      Case "em"
        
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
          frmControlMain.MSHFlexGrid1.Col = UREJAJ
            SQL = "Delete From em Where sifra='" & frmControlMain.MSHFlexGrid1.Text & "'"
         End If
         
      Case "tipa"
        
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
          frmControlMain.MSHFlexGrid1.Col = UREJAJ
            SQL = "Delete From TIP_ART Where sifra='" & frmControlMain.MSHFlexGrid1.Text & "'"
         End If
        
  Case "Category"
     ' Set Rs1 = myConection.Execute("Select * From mada where madasifr=" & frmControlMain.DataGrid1.Columns(0).text)
      '  If Rs1.RecordCount = 0 Then
        If MsgBox("Zbrišem Artikel?", vbInformation + vbYesNo) = vbYes Then
        
        frmControlMain.MSHFlexGrid1.Col = UREJAJ
        If Getnazi("select id_dok from nabasif where sifra='" & frmControlMain.MSHFlexGrid1.Text & "'") = "" Then
        SQL = "Delete from mada Where madasifr='" & (frmControlMain.MSHFlexGrid1.Text) & "'"
        Else
        MsgBox "Tega artikla ne morem izbrisati ker je že bil uporabljen"
        SQL = ""
        End If
        End If
       'Else
        '    MsgBox "Neuspešno!", vbInformation
           
        'End If
        
  Case "Location"
    If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
    frmControlMain.MSHFlexGrid1.Col = UREJAJ
        SQL = "Delete from grupa Where sifra=" & Val(frmControlMain.MSHFlexGrid1.Text)
    End If
  Case "Purchase Order"
    Set Rs1 = DCON.Execute("Select * From nabasif where stdok='" & frmControlMain.MSHFlexGrid1.Text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.Text))
        'If Rs1.RecordCount = 0 Then
        If MsgBox("Zbrišem podatek?", vbInformation + vbYesNo) = vbYes Then
        SQL = "Delete from nabasif Where stdok='" & frmControlMain.MSHFlexGrid1.Text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.Text)
        End If
       
  Case "Purchase Return"
     MsgBox "Neuspešno!", vbInformation
         
        'Delete from PurchaseReturnDetail Where PurchaseReturnID='" & "'"
        'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "Purchase Registry"
     MsgBox "Neuspešno!", vbInformation
         
       'Delete from PurchaseOrderDetail Where PurchaseOrderID='" & "'"
       'Delete from PurchaseOrderHeader Where PurchaseOrderID='" & "'"
  Case "MATE"
        If MsgBox("Naredim storno raèuna " & frmControlMain.MSHFlexGrid1.Text & " ? ", vbInformation + vbYesNo) = vbYes Then
'        DCON.Execute "insert into storno select * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Val(frmControlMain.MSHFlexGrid1.text) & "'"
 myConection.Execute "Delete from glavna where tip_dok='" & tip_dok & "' and id_dok='" & (frmControlMain.MSHFlexGrid1.Text) & "'"
        myConection.Execute "Delete from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & (frmControlMain.MSHFlexGrid1.Text) & "'"
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
    CreateH_Page "Select *  from nabasif where  stdok='" & frmControlMain.MSHFlexGrid1.Text & "' and sifrapart=" & Val(frmControlMain.MSHFlexGrid1.Text) & "", " Prevzemni list št.: " & frmControlMain.MSHFlexGrid1.Text
  Case "Purchase Return"
        CreateH_Page "Select datum,min(st) as zac,max(st) as kon,sum(znesek) as znesek from racusif  where datum=#" & (frmControlMain.MSHFlexGrid1.Text) & "# group by datum", " Rekapitulacija "
  Case "Purchase Registry"
    CreateH_Page "Select sifra,naziv,kol,znesek from storno where st=" & Val(frmControlMain.MSHFlexGrid1.Text), " Raèun št: " & frmControlMain.MSHFlexGrid1.Text & ", z dne : " & frmControlMain.MSHFlexGrid1.Text
   Case "Sales Return"
        CreateH_Page "Select sifra,naziv,kol,znesek from racusif where st=" & Val(frmControlMain.MSHFlexGrid1.Text), " Raèun št: " & frmControlMain.MSHFlexGrid1.Text & ", z dne : " & frmControlMain.MSHFlexGrid1.Text
   Case "Sales Registry"
        CreateH_Page "Select * from tdr", " Details "
  Case "MATE"
    CreateH_Page "Select sifra,naziv,kol,znes from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", " Dokument št.: " & tip_dok & frmControlMain.MSHFlexGrid1.Text
   
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
   PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
   
'little bit of blabla this method is very widely used in Scripting Lanuguage very powerful code but very gentle
adder:
Exit Sub
End Sub
Private Sub mnuvsedni_Click()
Dim iskn As String
Dim aaa As Integer
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PLDNI'") = "" Then
aaa = 0
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PLDNI'")
End If
iskn = InputBox("Vnesi plaèilo dni za vse partnerje", "Vnesi plaèilo dni za vse partnerje", aaa)
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='PLDNI'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','PLDNI','" & iskn & "')")

End If
End Sub
Private Sub mnugend_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='GEND'") = "" Then
aaa = ""
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='GEND'")
End If
iskn = InputBox("Vnesi pgeneralni datum oblika DD.MM.LLLL", "Vnesi gen. datum", aaa)
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='GEND'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','GEND','" & iskn & "')")

End If
End Sub
Private Sub mnuavtoopis_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='AVTOOP'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='AVTOOP'")
End If
iskn = UCase(InputBox("Avto opis na pozicijah? D/N", "Avto opis na pozicijah? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='AVTOOP'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','AVTOOP','" & iskn & "')")

End If
End Sub
Private Sub mnucezpol_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CEZPO'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CEZPO'")
End If
iskn = UCase(InputBox("Delovni èas cez polnoè? D/N", "Delovni èas cez polnoè? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='CEZPO'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','CEZPO','" & iskn & "')")

End If
End Sub
Private Sub mnucenapa_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CENAPA'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CENAPA'")
End If
iskn = UCase(InputBox("Cena na Gumbih PA? D/N", "Cena na Gumbih PA? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='CENAPA '")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','CENAPA','" & iskn & "')")

End If
End Sub
Private Sub mnuzaklj_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='ZAKPA'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='ZAKPA'")
End If
iskn = UCase(InputBox("Zakljuèim PA? D/N", "Zakljuèim PA? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='ZAKPA '")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','ZAKPA','" & iskn & "')")

End If
End Sub
Private Sub mnupoppa_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='POPPA'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='POPPA'")
End If
iskn = UCase(InputBox("Izpišem popust na pozicijah PA? D/N", "Popust na pozicijah? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='POPPA'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','POPPA','" & iskn & "')")

End If
End Sub

Private Sub mnuotvoritev_Click()
If MsgBox("Ali želiš narediti otvoritev???", vbOKCancel) = vbOK Then
Dim fso As New FileSystemObject
    
    
    If fso.FolderExists("c:\arhiv" & LTrim(str(Year(Date) - 1))) = False Then
        fso.CreateFolder ("c:\arhiv" & LTrim(str(Year(Date) - 1)))
     Call Shell("xcopy " & RTrim(App.path) & " c:\arhiv" & LTrim(str(Year(Date) - 1)) & " /E", vbNormalFocus)
    Else

myConection.Execute ("delete from dokm where tip_dok='NA'")
myConection.Execute ("delete from dokm where tip_dok='IZ'")
myConection.Execute ("delete from glavna where tip_dok='NA'")
myConection.Execute ("delete from glavna where tip_dok='IZ'")
myConection.Execute ("delete from nabasif where tip_dok='NA'")
myConection.Execute ("delete from nabasif where tip_dok='IZ'")
MsgBox ("Konèano")
End If
End If
End Sub
Private Sub mnuopispo_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='OPISPO'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='OPISPO'")
End If
iskn = UCase(InputBox("Opis poz. v Šifr.artiklov na pozicijah? D/N", "Opis na pozicijah? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='OPISPO'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','OPISPO','" & iskn & "')")

End If
End Sub
Private Sub mnuskrsto_Click()
Dim iskn As String
Dim aaa As String
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'") = "" Then
aaa = "N"
Else
aaa = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='SKRIST'")
End If
iskn = UCase(InputBox("Skrijm Storno gumbe na PA? D/N", "Skrij storno gumbe? D/N", aaa))
If iskn <> "D" Then
iskn = "N"
End If
If iskn <> "" Then
myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='SKRIST'")
myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','SKRIST','" & iskn & "')")

End If
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
MODIFYID = frmControlMain.MSHFlexGrid1.Text


Select Case CatalogueName

 Case "MATE"
 ma_ured = 1
 blag.Show
 
  Case "Customer"
    Load C_frmCustomer
    C_frmCustomer.Show vbModal
  Case "Supplier"
 
  std = Trim(frmControlMain.MSHFlexGrid1.Text) & Trim(frmControlMain.MSHFlexGrid1.Text)
  'MsgBox std
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
    If frmControlMain.Wbrow.Visible = True Then
        frmControlMain.Wbrow.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPrintPrv_Click()
On Error GoTo adder:
    If frmControlMain.Wbrow.Visible = True Then
        frmControlMain.Wbrow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
    End If
Exit Sub
adder:
    Exit Sub
End Sub
Private Sub mnuPageSetup_Click()
On Error GoTo adder:
If frmControlMain.Wbrow.Visible = True Then
        frmControlMain.Wbrow.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
End If
    Exit Sub
adder:
    Exit Sub
End Sub

Private Sub mnuPurchaseRegister_Click()
 frmControlMain.Wbrow.Visible = True
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
 frmControlMain.Wbrow.Visible = True
    frmControlMain.MSHFlexGrid1.Visible = False
rptState = "SalesRegistry"

Load Form1
Form1.Show vbModal



End Sub

Private Sub mnuSave_Click()
'    On Error GoTo adder:
'If frmControlMain.wbrow.Visible = True Then
        'frmControlMain.WBrow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER, 1
        
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

 PaintMenuBitmaps
Dim i As Integer
 'For I = 0 To MenuHeader.Count - 1
 '       'Decide if the menu is already contracted and if so display the expand header
 '       If FrameXPMenu(I).Height = imgUp.Height Then
 '           MenuHeader(I).Picture = imgDown.Picture
 '       Else
 '           MenuHeader(I).Picture = imgUp.Picture
 '       End If
 '       MenuHeader(I).Height = 375
 '       MenuHeader(I).Width = FrameXPMenu(I).Width
 '   Next
    
 '   doresize = False
'lblHeader_Click (1)
'lblHeader_Click (2)

  'Me.EyeDropper1
        
    '-- Set The Custom Colour Properties
    '-- I've Opted For A Purple Look
    '-- You Can Set The Colours To What ever You Want
    Me.EyeDropper1.SetCustomProperties &HFF80FF, &H800080, vbWhite, _
                &H800080, &HFF80FF, &HC000C0, vbWhite, _
                &HFFC0FF, &H800080, &H800080, &HFFC0FF
    
    '-- Load Sample Items
    Call LoadItems
    Me.StatusBar1.Panels(2).Text = UPORABNIK
     Me.StatusBar1.Panels(4).Text = Now
 Me.EyeDropper1.VisibleItems = 4
If Getnazi("select naziv from izpisi where tip_dok='VS' and naziv='OSNOVNI IZPIS'") = "" Then
If rs.State = 1 Then rs.Close
 rs.Open "select * from izpisi", myConection, adOpenDynamic, adLockOptimistic
 rs.AddNew
 rs.Fields("tip_dok") = "VS"
 rs.Fields("naziv") = "OSNOVNI IZPIS"
 rs.Fields("pozicija") = 4
 rs.Update
 rs.AddNew
 rs.Fields("tip_dok") = "VS"
 rs.Fields("naziv") = "PREGLED ZALOG"
 rs.Fields("pozicija") = 3
 rs.Update
 End If
 If Getnazi("select naziv from izpisi where tip_dok='VS' and naziv='PREGLED ZALOG - FIFO'") = "" Then
 If rs.State = 1 Then rs.Close
 rs.Open "select * from izpisi", myConection, adOpenDynamic, adLockOptimistic
 rs.AddNew
 rs.Fields("tip_dok") = "VS"
 rs.Fields("naziv") = "PREGLED ZALOG - FIFO"
 rs.Fields("pozicija") = 5
 rs.Update

 End If
 ime_form

End Sub
Private Sub LoadItems()
  Dim i     As Long
  Dim iItems As Long
  Dim Rsa As New ADODB.Recordset
  Dim oItem As cEDItem
  Dim onode As node
 
    With Me.EyeDropper1
        'Me.vbalImageList1
        .ImageList = Me.ImageList1
        .Redraw = False
        .EyeDropperItems.clear
        .Redraw = False
        '-- Add Some Sample nodes
        
     For i = 0 To 4
        TreeView1(i).Nodes.clear
          
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM MENU_I WHERE ID=2  and kje=" & i + 1 & " order by zaporedna", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
       rs.MoveFirst
       Dim v, xxc As Integer
       
       v = 1
       
       Do While Not rs.EOF
        Dim IM As String
        IM = Trim(rs.Fields("naziv"))
            Set onode = TreeView1(i).Nodes.Add(, , "Main" & v, "(" & presled(LTrim(str(rs.Fields("zaporedna"))), 3) & ")  " & IM)
            
            If Rsa.State = 1 Then Rsa.Close
             Rsa.Open "SELECT * FROM MENU_I WHERE ID=3 and kje=" & rs.Fields("zaporedna") & " order by zaporedna ", myConection, adOpenDynamic, adLockOptimistic
If Not Rsa.EOF Then
                Rsa.MoveFirst
            xxc = 1
            Do While Not Rsa.EOF
                Set onode = TreeView1(i).Nodes.Add("Main" & v, tvwChild, "MainSub" & xxc & "Main" & v, "(" & presled(LTrim(str(Rsa.Fields("zaporedna"))), 3) & ")  " & Rsa.Fields("naziv"))
            
            Rsa.MoveNext
            xxc = xxc + 1
            Loop
    End If
            v = v + 1
            rs.MoveNext
        Loop
        
   End If
     Next i
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM MENU_I WHERE ID=1 order by zaporedna", myConection, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
        '-- Add Some Sample Panels
        Dim XD As Integer
        XD = 1
        Do While Not rs.EOF
        i = XD
        
            Select Case i
                Case Is = 1
                Set oItem = .EyeDropperItems.Add("Item: " & i, , rs.Fields("naziv"))
                Case Is = 2
                Set oItem = .EyeDropperItems.Add("Item: " & i, , rs.Fields("naziv"))
                Case Is = 3
                Set oItem = .EyeDropperItems.Add("Item: " & i, , rs.Fields("naziv"))
                Case Is = 4
                Set oItem = .EyeDropperItems.Add("Item: " & i, , rs.Fields("naziv"))
                Case Is = 5
                Set oItem = .EyeDropperItems.Add("Item: " & i, , rs.Fields("naziv"))
            End Select
            
        
            If i = 1 Then
                oItem.Panel = Me.TreeView1(i - 1)
            End If
            If i = 2 Then
                oItem.Panel = Me.TreeView1(i - 1)
            End If
            If i = 3 Then
                oItem.Panel = Me.TreeView1(i - 1)
            End If
             If i = 4 Then
                oItem.Panel = Me.TreeView1(i - 1)
            End If
             If i = 5 Then
                oItem.Panel = Me.TreeView1(i - 1)
            End If
            oItem.IconIndex = i
            If i = 1 Then
                oItem.Selected = True
            End If
           ' If i = 4 Or i = 9 Then
           '     oItem.Visible = False
           '  Else
           '     oItem.Visible = True
           ' End If
           XD = XD + 1
            rs.MoveNext
        Loop
        .VisibleItems = 5
        .Redraw = True
        
    End With

End Sub

Private Sub Timer1_Timer()

 
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

'frmDelUser.Caption = "View Users"
'frmDelUser.Command1.Visible = False
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
pro = pro + Round(RS2.Fields("znesek") / (1 + (Val(Getnazi("select madapd from mada where madasifr='" & RS2.Fields("sifra") & "'")) / 100)), 2)
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
    frmControlMain.Wbrow.Visible = False
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
        SQL = "Select datum,  format(sum(znesek),'fixed') as znesek,min(st) as zacstrac,max(st) as konstarc  from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by datum order by datum"
        Else
        SQL = "Select datum,  format(sum(znesek) as znesek,'fixed') ,min(st) as zacstrac,max(st) as konstarc  from racusif group by datum order by datum"
        End If
        CatalogueName = "Purchase Return"
    Case 6
        If dejpre = 1 Then
        SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from storno where [datum] between #" & dod & "# AND #" & ddo & "# and stw<>'A' group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from storno where stw<>'A' group by st order by st"
         End If
        CatalogueName = "Purchase Registry"
    Case 7
     If dejpre = 1 Then
        SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by st order by st"
        Else
         SQL = "Select st,min(datum) as datum,  format(sum(znesek),'fixed') as znesek,min(oseba) as oseba from racusif group by st order by st"
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
    'frmMAIN.ProgressBar.Value = counter
   ' frmMAIN.Label5.Caption = Str(counter) & " %"
    SQL = "select embalaza,kol from nabasif where sifra=" & Round(rs3.Fields("madasifr"), 0)
    Set Rs1 = myConection.Execute(SQL)
    If Not Rs1.EOF Then
    Rs1.MoveFirst
    End If
    Do While Not Rs1.EOF
    na = na + (Round(Rs1.Fields("embalaza"), 2) * Round(Rs1.Fields("kol"), 2))
    Rs1.MoveNext
    Loop

    
     sql55 = "select * from mada where madasifr='" & (rs3.Fields("madasifr")) & "'"
    If ConnectRS(myConection, rs55, sql55) = False Then
            
        End If
        
    
    rs55.MoveFirst
    rs55.Fields("madazalo") = Round(na, 2)
    rs55.Fields("madasest") = "N"
    rs55.Update
  
  Set rs55 = Nothing
  
    
    
    'myConection.Execute "Update mada set madazalo=" & Round(na, 2) & " where madasifr=" & rs3.Fields("madasifr")
    'MsgBox ("nabava " & na)
    opa = Getnazi("select madaenme from mada where madasifr='" & (rs3.Fields("madasifr")) & "'")
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
    sql55 = "select * from mada where madasifr='" & (rs3.Fields("madasifr")) & "'"
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
    sql55 = "select * from mada where madasifr='" & (rs3.Fields("madasifr")) & "'"
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
    myConection.Execute "Update mada set madazalo=0 where madasifr='" & (rs3.Fields("madasifr")) & "'"
    myConection.Execute "Update mada set madasest='D' where madasifr='" & (rs3.Fields("madasifr")) & "'"
    
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
    opa = Getnazi("select madaenme from mada where madasifr='" & ss & "'")
    ops = rs3.Fields("kol")
    SQL = "select sum(kol) as kkoo from racusif where sifra=" & ssa
       If rs.State = 1 Then rs.Close
       
    rs.Open SQL, myConection, adOpenStatic, adLockOptimistic
    If IsNull(rs.Fields("kkoo")) Then
    kk = 0
    Else
    kk = rs.Fields("kkoo") * ops
    End If
    'sql = "select * from mada where madasifr=" & ss
    '   If Rs.State = 1 Then Rs.Close
       
    'Rs.Open sql, myConection, adOpenStatic, adLockOptimistic
    myConection.Execute "Update mada set madazalo=madazalo-" & kk & " where madasifr='" & ss & "'"
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
Set rs = myConection.Execute(sql1)
If rs.Fields("stma") - Rs1.Fields("st") > 0 Then
If MsgBox("Ali želiš uvoziti podatke???", vbOKCancel) = vbOK Then
myConection.Execute ("delete from racusif where st<=" & Rs1.Fields("st"))

SQL = "select * from racusif"
Set Rs1 = dcon1.Execute(SQL)

If rs.State = 1 Then rs.Close

rs.Open "select * from racusif ", myConection, adOpenStatic, adLockOptimistic

Rs1.MoveFirst
Do While Not Rs1.EOF
rs.AddNew
rs.Fields("st") = Rs1.Fields("st")
rs.Fields("datum") = Rs1.Fields("datum")
rs.Fields("oseba") = Rs1.Fields("oseba")
rs.Fields("kol") = Rs1.Fields("kol")
rs.Fields("znesek") = Rs1.Fields("znesek")
rs.Fields("doza") = Rs1.Fields("doza")
rs.Fields("placilo") = Rs1.Fields("placilo")
rs.Fields("vst") = Rs1.Fields("vst")
rs.Fields("ura") = Rs1.Fields("ura")
rs.Fields("sifra") = Rs1.Fields("sifra")

rs.Update

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
    
End If
dex = 2
End Sub

Public Sub TreeView1_NodeClick(Index As Integer, ByVal node As MSComctlLib.node)
myConection.Execute ("update dokm set poz=poz+1 where atribut='STAP'")
myConection.Execute ("update dokm set poz=1 where poz=6 and atribut='STAP'")
myConection.Execute ("update dokm set id_dok='" & Mid(node.Text, 6, 20) & "' where poz=1 and atribut='STAP'")
myConection.Execute ("update dokm set tekst='" & Mid(node.Text, 2, 3) & "' where poz=1 and atribut='STAP'")


frmControlMain.Wbrow.Visible = False
tip_dok = ""
zapore = Val(Mid(node.Text, 2, 3))
vmessql = Getnazi("select ukaz from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3)))
SQL = Replace(Getnazi("select ukaz from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3))), "<and>", "")
CatalogueName = Getnazi("select cat from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3)))
If Getnazi("select cat from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3))) = "MATE" Then
tip_dok = Mid(node.Text, 8, 2)
End If
If Getnazi("select cat from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3))) = "FAXX" Then
tip_dok = Mid(node.Text, 8, 2)
End If
If Getnazi("select cat from menu_i where zaporedna=" & Val(Mid(node.Text, 2, 3))) = "KOMP" Then
tip_dok = Mid(node.Text, 8, 2)
End If

Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "MATE" Then
If SQL = "" Then
MsgBox "Nimaš vnesenega SQL ukaza!"
erro = "1"
Else
erro = ""
'MsgBox SQL
SQL = Replace(SQL, "<where>", "")
SQL = Replace(SQL, "<and>", "")

Set Rs1 = DCON.Execute(SQL)
ssqq = SQL
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
  
End If
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing
  frmControlMain.osv_Click
End Sub
Public Sub beno_os(ii As Integer)
frmControlMain.Wbrow.Visible = False
tip_dok = ""
zapore = ii
vmessql = Getnazi("select ukaz from menu_i where zaporedna=" & ii)
SQL = Replace(Getnazi("select ukaz from menu_i where zaporedna=" & ii), "<and>", "")
CatalogueName = Getnazi("select cat from menu_i where zaporedna=" & ii)
If Getnazi("select cat from menu_i where zaporedna=" & ii) = "MATE" Then
tip_dok = Left(Getnazi("select naziv from menu_i where zaporedna=" & ii), 2)
End If
If Getnazi("select cat from menu_i where zaporedna=" & ii) = "FAXX" Then
tip_dok = Left(Getnazi("select naziv from menu_i where zaporedna=" & ii), 2)
End If
If Getnazi("select cat from menu_i where zaporedna=" & ii) = "KOMP" Then
tip_dok = Left(Getnazi("select naziv from menu_i where zaporedna=" & ii), 2)
End If

Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "MATE" Then
If SQL = "" Then
MsgBox "Nimaš vnesenega SQL ukaza!"
erro = "1"
Else
erro = ""
'MsgBox SQL
SQL = Replace(SQL, "<where>", "")
SQL = Replace(SQL, "<and>", "")

Set Rs1 = DCON.Execute(SQL)
ssqq = SQL
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
  
End If
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing
  frmControlMain.osv_Click

End Sub
Private Sub Vrzi_Click()
If Me.EyeDropper1.Width = Me.Vrzi.Width Then
Me.Picture1x.Width = 4000
Me.EyeDropper1.Width = 3920
Me.Vrzi.Caption = "<=="
Call Size(frmControlMain)
Else
Me.Picture1x.Width = Me.Vrzi.Width + 30
Me.EyeDropper1.Width = Me.Vrzi.Width
Me.Vrzi.Caption = "==>"
Call Size(frmControlMain)
End If
End Sub
