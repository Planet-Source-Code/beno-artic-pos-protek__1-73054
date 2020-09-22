VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmControlMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   ForeColor       =   &H00404040&
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin ProsVent.UserControl2 UserControl21 
      Height          =   975
      Left            =   4080
      Top             =   3360
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   195
      Left            =   11760
      TabIndex        =   36
      Top             =   120
      Width           =   135
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   7995
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command2"
      Height          =   135
      Left            =   5040
      TabIndex        =   31
      Top             =   0
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6600
      TabIndex        =   30
      Text            =   "100"
      Top             =   1200
      Width           =   615
   End
   Begin LVbuttons.LaVolpeButton kontr 
      Height          =   255
      Left            =   10560
      TabIndex        =   29
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "KONTROLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   6.75
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
      MICON           =   "frmControlMains.frx":0000
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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   11520
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   11400
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   8640
      Top             =   1680
   End
   Begin LVbuttons.LaVolpeButton tvor 
      Height          =   375
      Left            =   10080
      TabIndex        =   20
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "=>"
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
      MICON           =   "frmControlMains.frx":001C
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmControlMains.frx":0038
      Left            =   9480
      List            =   "frmControlMains.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   0
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMains.frx":003C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMains.frx":5F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMains.frx":6BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMains.frx":6F04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton izpi 
      Height          =   495
      Left            =   2220
      TabIndex        =   14
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Pregled"
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
      MICON           =   "frmControlMains.frx":700E
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.CommandButton osv 
      Caption         =   "Command2"
      Height          =   135
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Height          =   270
      Left            =   0
      MaskColor       =   &H8000000F&
      Picture         =   "frmControlMains.frx":702A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   270
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   180
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   75
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing... Please wait"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   3405
      End
   End
   Begin LVbuttons.LaVolpeButton zakl 
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Izpis zaklj"
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
      MICON           =   "frmControlMains.frx":7124
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
   Begin VB.ComboBox text1 
      Height          =   315
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Po grupi"
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
      MICON           =   "frmControlMains.frx":7140
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
   Begin MSComCtl2.DTPicker datdo 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   39075
   End
   Begin MSComCtl2.DTPicker datod 
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   39075
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7440
      Top             =   1680
   End
   Begin SHDocVwCtl.WebBrowser Wbrow 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   4260
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin LVbuttons.LaVolpeButton zalog 
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Izpis zalog"
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
      MICON           =   "frmControlMains.frx":715C
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
   Begin LVbuttons.LaVolpeButton NOVA 
      Height          =   975
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Nov"
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
      MICON           =   "frmControlMains.frx":7178
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton UREDI 
      Height          =   975
      Left            =   1125
      TabIndex        =   16
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "UREDI"
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
      MICON           =   "frmControlMains.frx":7194
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton knj 
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Faktura"
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
      MICON           =   "frmControlMains.frx":71B0
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmControlMains.frx":71CC
      Height          =   3600
      Left            =   0
      TabIndex        =   18
      Top             =   1440
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   6350
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
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
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin LVbuttons.LaVolpeButton lansiraj 
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Lansiraj DN"
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
      MICON           =   "frmControlMains.frx":74D6
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
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
      COLTYPE         =   2
      BCOL            =   16777215
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmControlMains.frx":74F2
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "POKNJIŽENI"
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
      COLTYPE         =   2
      BCOL            =   12632319
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmControlMains.frx":750E
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "NEPOKNJIŽENI"
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
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmControlMains.frx":752A
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton6 
      Height          =   735
      Left            =   10560
      TabIndex        =   25
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "KNJIŽI"
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
      MICON           =   "frmControlMains.frx":7546
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton izopi 
      Height          =   495
      Left            =   2220
      TabIndex        =   28
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Izpis"
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
      MICON           =   "frmControlMains.frx":7562
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   270
      Left            =   7440
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin LVbuttons.LaVolpeButton n2 
      Height          =   495
      Left            =   7560
      TabIndex        =   33
      Top             =   360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "FIFO"
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
      MICON           =   "frmControlMains.frx":757E
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
   Begin LVbuttons.LaVolpeButton analize 
      Height          =   375
      Left            =   8520
      TabIndex        =   34
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Analize"
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
      MICON           =   "frmControlMains.frx":759A
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
   Begin LVbuttons.LaVolpeButton zalna 
      Height          =   495
      Left            =   8520
      TabIndex        =   37
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Zaloga na dan"
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
      MICON           =   "frmControlMains.frx":75B6
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
      Height          =   495
      Left            =   9480
      TabIndex        =   38
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "TDR"
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
      MICON           =   "frmControlMains.frx":75D2
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
   Begin MSForms.CheckBox vklj 
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   255
      BackColor       =   -2147483633
      ForeColor       =   4210752
      DisplayStyle    =   4
      Size            =   "450;450"
      Value           =   "0"
      FontHeight      =   165
      FontCharSet     =   238
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "od"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "do"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   0
      Picture         =   "frmControlMains.frx":75EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11850
   End
End
Attribute VB_Name = "frmControlMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' variables for ta binding
Private datPrimaryRS As ADODB.Recordset

' variables for enabling column sort
Private m_iSortCol As Integer
Private m_iSortType As Integer

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub analize_Click()
Analiza.Show
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton Then
Call FG_AutosizeCols(MSHFlexGrid1, Me, , , False)

 'For i = MSHFlexGrid1.FixedCols To MSHFlexGrid1.Cols - 1
 'MSHFlexGrid1.Col = I
' MSHFlexGrid1.ColWidth(i) = 200
' Next i
End If
End Sub
 
  
Private Sub Command1_Click()

SaveFlexGridColumnWidths MSHFlexGrid1, tip_dok & CatalogueName

End Sub

Private Sub Command2_Click()
'frmblag.Show
Dim ber As String
ber = InputBox("Sql", "sql", Getnazi("select polja from dokumenti where tip_dok='" & tip_dok & "'"))
myConection.Execute ("update dokumenti set polja='" & ber & "' where tip_dok='" & tip_dok & "'")
End Sub

Private Sub Command3_Click()
'If MSHFlexGrid1.Visible = True Then
'    MSHFlexGrid1.Visible = False
'    Call PrintFlexi("", Me.MSHFlexGrid1)
'    MSHFlexGrid1.Visible = True
'End If
Dim ber As String
ber = InputBox("Sql", "sql", Getnazi("select ukaz from menu_i where zaporedna=" & zapore))
If rs.State = 1 Then rs.Close
rs.Open "select * from menu_i where zaporedna=" & zapore, myConection, adOpenDynamic, adLockOptimistic
rs.Fields("ukaz") = ber
'myConection.Execute ("update menu_i set ukaz='" & ber & "' where zaporedna=" & zapore)
rs.Update
End Sub



Private Sub Command4_Click()
'Dim ber As String
'ber =ox("Sql", "sql", Getnazi("select polja from dokumenti where tip_dok='NT'"))
'myConection.Execute ("update dokumenti set polja='" & ber & "' where tip_dok='NT'")
'If RS.State = 1 Then RS.Close
'RS.Open "SELECT SIFRA,EM FROM ARTIKLI", myConection, adOpenDynamic, adLockOptimistic
'If Not RS.EOF Then
'RS.MoveFirst
'End If
'Dim rst As New ADODB.Recordset
'Do While Not RS.EOF
'If rst.State = 1 Then rst.Close
'rst.Open "SELECT MADASIFR,MADAENME FROM MADA WHERE MADASIFR='" & LTrim(Str(Val(RS.Fields("SIFRA")))) & "'", myConection, adOpenDynamic, adLockOptimistic
'If Not rst.EOF Then
'rst.MoveFirst
'End If
'rst.Fields("MADAENME") = RS.Fields("EM")
'rst.Update

'RS.MoveNext
'Loop
'myConection.Execute ("delete from menu_i where naziv='FIFO PO ARTIKLIH'")
'If RS.State = 1 Then RS.Close
'RS.Open "select * from menu_i", myConection, adOpenDynamic, adLockOptimistic
'RS.AddNew
'RS.Fields("id") = 3
'RS.Fields("naziv") = "FIFO PO ARTIKLIH"
'RS.Fields("zaporedna") = 30
'RS.Fields("ukaz") = "SELECT tip_dok, id_dok,format(datum,'dd.mm.yyyy') as datum,  sifra, kol as skkol, cena as madampcd, veza_td, veza_id,format(prosta,'standard') as prosta FROM zaloga <where> ORDER BY sifra, datum, veza_td,id_dok"
'RS.Fields("kje") = 27
'RS.Fields("cat") = "Category"
'RS.Update
'MsgBox "Konèano!"
'myConection.Execute ("insert into menu_i (id,naziv,zaporedna,ukaz,kje,cat) values (3,'FIFO PO ARTIKLIH',30,'SELECT tip_dok, id_dok,format(datum,'dd.mm.yyyy') as datum,  sifra, kol as skkol, cena as madampcd, veza_td, veza_id FROM zaloga <where> ORDER BY sifra, datum, veza_td,id_dok',27,'Category')")
Placa.Show vbModal
End Sub

Private Sub Command5_Click()
Xpl.Show
End Sub

Private Sub datdo_change()
osv_Click
End Sub

Private Sub datod_change()
osv_Click
End Sub

Private Sub izopi_Click()
Dim reporx As String
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
If CatalogueName = "MATE" Or CatalogueName = "FAXX" Or CatalogueName = "KOMP" Then
frmControlMain.MSHFlexGrid1.Col = UR_id
xid_dok = frmControlMain.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, frmControlMain.MSHFlexGrid1.Col)
If Getnazi("select naziv from izpisi where naziv like '* %' and tip_dok='" & tip_dok & "'") = "" Then
MsgBox "Nimate izbranega default poroèila! Odprite predogled ter ga izberite!"
Exit Sub
End If
reporx = Getnazi("select naziv from izpisi where naziv like '* %' and tip_dok='" & tip_dok & "'")
If Left(Getnazi("select naziv from izpisi where naziv like '* %' and tip_dok='" & tip_dok & "'"), 1) = "*" Then
repor = Mid(Getnazi("select naziv from izpisi where naziv like '* %' and tip_dok='" & tip_dok & "'"), 3)

Else
repor = Getnazi("select naziv from izpisi where naziv like '* %' and tip_dok='" & tip_dok & "'")

'MsgBox repor
End If
'MsgBox "'" & repor & "'"
End If
 If Val(Getnazi("select pozicija from izpisi where naziv='" & reporx & "'")) = 1 Then
      Call Print_dob(repor)
     End If
     If Val(Getnazi("select pozicija from izpisi where naziv='" & reporx & "'")) = 2 Then
      Call Print_preg(repor)
     End If
    If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 3 Then
      zalo.Show
     End If
      If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 4 Then
      Call Print_osn(repor, frmControlMain.MSHFlexGrid1)
     End If
       If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 5 Then
      Call Print_zal_fifo(repor)
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 7 Then
      Call PrintFlexix(repor)
     End If
      If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 10 Then
     

myConection.Execute ("delete from bbe")

myConection.Execute ("Insert into bbe SELECT * from zaloga where sifra='" & MODIFYID & "' order by datum,id_dok,poz")
Dim rst As New ADODB.Recordset
Set rst = myConection.OpenRecordset("bbe")
If Not rst.EOF() Then
rst.MoveFirst
Dim aha As Double
Dim vre As Double
aha = 0
vre = 0
Do While Not rst.EOF
'rst.Edit
If rst.Fields("tip_dok") = "NA" Then
rst.Fields("kon") = rst.Fields("kol")

rst.Fields("vri") = 0
rst.Fields("vrn") = rst.Fields("vrednost")

rst.Fields("koi") = 0

End If
If rst.Fields("tip_dok") = "IZ" Then

rst.Fields("vri") = rst.Fields("vrednost")
rst.Fields("vrn") = 0

rst.Fields("koi") = rst.Fields("kol")
rst.Fields("kon") = 0
End If
aha = aha + Round(rst.Fields("kon") + rst.Fields("koi"), 3)
vre = vre + Round(rst.Fields("vrn") + rst.Fields("vri"), 3)

rst.Fields("prosta") = aha
rst.Fields("prostav") = vre
rst.Update
rst.MoveNext
Loop
End If

     PRINTSNAP repor, ""
     'sgBox (MODIFYID)
     End If
     
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 8 Then
     PRINTSNAP repor, "tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 9 Then
     
     'MsgBox "datum>=#" & Replace(frmControlMain.datod.Value, ".", "-") & "# and datum<=#" & Replace(frmControlMain.datdo.Value, ".", "-") & "#"
     'MsgBox "datum>=#" & Format(frmControlMain.datod.Value, "mm/dd/yyyy") & "# and datum<=#" & Format(frmControlMain.datdo.Value, "mm/dd/yyyy") & "#"
     If repor = "DNEVNIK" Then
     PRINTSNAP repor, "datum>=#" & Replace(Format(frmControlMain.DATOD.Value, "mm/dd/yyyy"), ".", "/") & "# and datum<=#" & Replace(Format(frmControlMain.DATDO.Value, "mm/dd/yyyy"), ".", "/") & "#"
     Else
     PRINTSNAP repor, "tip_dok='" & tip_dok & "' and datum>=#" & Replace(Format(frmControlMain.DATOD.Value, "mm/dd/yyyy"), ".", "/") & "# and datum<=#" & Replace(Format(frmControlMain.DATDO.Value, "mm/dd/yyyy"), ".", "/") & "#"
     End If
     End If
     
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 6 Then
      Call Print_dob_les(repor)
     End If
End Sub

Private Sub kontr_Click()
If intCtrlDown = 2 Then
If MsgBox("Ali naredim preraèun FIFO?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then

pre_fifo_vse ""
End If
intCtrlDown = 0
Else
If rs.State = 1 Then rs.Close
rs.Open "select tip_dok,id_dok,sum(FAKTOR*kol*(cena*(1-(pop/100)))) as znesek from nabasif where poknj='K' group by tip_dok,id_dok order by tip_dok,id_dok", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Dim kirr As String
kirr = ""
Do While Not rs.EOF
tti = rs.Fields("tip_dok")
iid = rs.Fields("id_dok")
zne = Round(rs.Fields("znesek"), 2)
'MsgBox Getnazi("select sum(vrednost) as vr from zaloga where tip_dok='" & tti & "' and id_dok='" & iid & "'")
'MsgBox zne
If Getnazi("select sum(vrednost) as vr from zaloga where tip_dok='" & tti & "' and id_dok='" & iid & "'") <> "" Then
If Round(Getnazi("select sum(vrednost) as vr from zaloga where tip_dok='" & tti & "' and id_dok='" & iid & "'"), 2) <> zne Then
kirr = kirr & tti & iid & ","
End If
End If
rs.MoveNext
Loop
If kirr = "" Then
MsgBox "Vse ok!!"
Else
MsgBox kirr
End If
End If
End If
End Sub

Private Sub Label1_Click()
MsgBox Getnazi("select sum(prosta*cena) as xx from zaloga")
End Sub

Private Sub lansiraj_Click()
frmControlMain.MSHFlexGrid1.Col = UR_id
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
MsgBox "Ta dokument je že LANSIRAN!"
Exit Sub
End If
Dim norma, stek, lesi As String
Dim koli, xfxt As Long
Dim xrsn As New ADODB.Recordset
Dim xox, yoy, zapp, dkr As Integer
Dim kol, XX, YY As Long
dkr = 1
xfxt = 1

imedn = frmControlMain.MSHFlexGrid1.Text
If xrsn.State = 1 Then xrsn.Close
xrsn.Open "select * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", myConection, adOpenDynamic, adLockOptimistic
If Not xrsn.EOF Then
xrsn.MoveFirst
End If
myConection.Execute "delete from normati"
myConection.Execute "delete from xnorm"
Do While Not xrsn.EOF

norma = Trim(xrsn.Fields("sifra"))
If Getnazi("select madanazi from mada where madasifr='" & norma & "'") Like "% DK%" Then
dkr = 2.05
End If
If Getnazi("select madaemba from mada where madasifr='" & norma & "'") = "" Then
xfxt = 1
Else
xfxt = Getnazi("select madaemba from mada where madasifr='" & norma & "'")
End If

stek = Trim(xrsn.Fields("stdok"))
lesi = Trim(xrsn.Fields("kopija"))
xox = (Trim(xrsn.Fields("x")) / dkr)
yoy = (Trim(xrsn.Fields("y")) / dkr)
koli = xrsn.Fields("kol")
XX = xrsn.Fields("x")
YY = xrsn.Fields("y")
Dim rst As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If rst.State = 1 Then rst.Close

rst.Open "select * from nabasif where tip_dok='NT' and id_dok='" & norma & "'", myConection, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
rst.MoveFirst
End If
Dim sii, nazii As String

Dim fixx, ss As Integer
If rsta.State = 1 Then rsta.Close
rsta.Open "select * from xnorm", myConection, adOpenDynamic, adLockOptimistic
ss = 1
Do While Not rst.EOF
sii = rst.Fields("sifra")
nazii = rst.Fields("naziv")
kol = rst.Fields("kol")

fixx = IIf(rst.Fields("chk_fix") = "b", 1, 0)
rsta.AddNew
rsta.Fields("sifr") = sii
rsta.Fields("naz") = nazii
rsta.Fields("poz") = levi_pres(LTrim(str(ss)), 4)
If fixx = 0 Then
If Getnazi("select madaenme from mada where madasifr='" & sii & "'") = "KOM" Then
rsta.Fields("kol") = FormatNumber(kol * koli * (((XX / 100) * (YY / 100)) / xfxt), 0)
Else
rsta.Fields("kol") = FormatNumber(kol * koli * (((XX / 100) * (YY / 100)) / xfxt), 4)
End If
Else
rsta.Fields("kol") = kol * koli

End If
rsta.Update
rst.MoveNext
ss = ss + 1
Loop
If rst.State = 1 Then rst.Close
rst.Open "select * from nabasif where tip_dok='NT' and id_dok like 'X" & LTrim(norma) & "%' and x<=" & xox * 10 & " and y>=" & xox * 10, myConection, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
rst.MoveFirst
End If
'rsta.Open "select * from normati", myConection, adOpenDynamic, adLockOptimistic
ss = 1
Do While Not rst.EOF
sii = rst.Fields("sifra")
nazii = rst.Fields("naziv")
kol = rst.Fields("kol")
zapp = rst.Fields("placilo")
rsta.AddNew
rsta.Fields("sifr") = sii
rsta.Fields("naz") = nazii
rsta.Fields("poz") = levi_pres(LTrim(str(ss)), 4)
rsta.Fields("kol") = kol * koli
rsta.Fields("zap") = zapp
rsta.Update
rst.MoveNext
Loop
Dim ssqll As String
ssqll = "select * from nabasif where tip_dok='NT' and id_dok like 'Y" & LTrim(norma) & "%' and x<=" & yoy * 10 & " and y>=" & yoy * 10
'MsgBox ssqll
If rst.State = 1 Then rst.Close
rst.Open ssqll, myConection, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
rst.MoveFirst
End If

'rsta.Open "select * from normati", myConection, adOpenDynamic, adLockOptimistic
ss = 1
Do While Not rst.EOF
sii = rst.Fields("sifra")
nazii = rst.Fields("naziv")
kol = rst.Fields("kol")
zapp = rst.Fields("placilo")
rsta.AddNew
rsta.Fields("sifr") = sii
rsta.Fields("naz") = nazii
rsta.Fields("poz") = levi_pres(LTrim(str(ss)), 4)
rsta.Fields("kol") = kol * koli
rsta.Fields("zap") = zapp
rsta.Update
rst.MoveNext
Loop
'steklo
rsta.AddNew
rsta.Fields("sifr") = stek
rsta.Fields("naz") = Getnazi("select madanazi from mada where madasifr='" & stek & "'")
rsta.Fields("poz") = levi_pres(LTrim(str(ss)), 4)
rsta.Fields("kol") = XX * YY * koli / 10000 * 0.76
rsta.Update
'les
If rs.State = 1 Then rs.Close
If lesi <> "" Then
rs.Open "select * from sestavi where sifra=" & lesi, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Dim faktx As Double
faktx = Getnazi("select madaemba from mada where madasifr='" & lesi & "'")
Do While Not rs.EOF

rsta.AddNew
rsta.Fields("sifr") = rs.Fields("sifras")
rsta.Fields("naz") = Getnazi("select madanazi from mada where madasifr='" & rs.Fields("sifras") & "'")
rsta.Fields("poz") = levi_pres(LTrim(str(ss)), 4)

rsta.Fields("kol") = FormatNumber(((XX * YY * koli / 10000) / faktx) * rs.Fields("kol"), 3)
rsta.Update
rs.MoveNext
Loop
End If
xrsn.MoveNext
Loop
If rs.State = 1 Then rs.Close
rs.Open "select sifr,min(naz) as naz,sum(kol) as kol,sum(zap) as zap from xnorm group by sifr", myConection, adOpenDynamic, adLockOptimistic
Dim Rsa As New ADODB.Recordset
If Rsa.State = 1 Then Rsa.Close
Rsa.Open "select * from normati", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If

Do While Not rs.EOF
Rsa.AddNew
Rsa.Fields("sifr") = rs.Fields("sifr")
Rsa.Fields("naz") = rs.Fields("naz")
Rsa.Fields("kol") = rs.Fields("kol")
Rsa.Fields("zap") = rs.Fields("zap")
Rsa.Update
rs.MoveNext
Loop
If Rsa.State = 1 Then Rsa.Close
Rsa.Open "select * from normati where sifr='10217'", myConection, adOpenDynamic, adLockOptimistic
If Not Rsa.EOF Then
'Rsa.Fields("kol") = Getnazi("select sum(zap) as x from normati")
'Rsa.Update

End If
kosovni = 1
tip_dok = "IZ"
MsgBox "DN je uspešno lansiran"
myConection.Execute ("update nabasif set poknj='K' where tip_dok='DN' and id_dok='" & imedn & "'")
NOVA_Click
End Sub

Private Sub LaVolpeButton3_Click()
dodajwh = " "
osve = frmControlMain.MSHFlexGrid1.Row
End Sub

Private Sub LaVolpeButton4_Click()
Dim iLoop As Integer
       If CatalogueName = "MATE" Then
       Dim jkl As Integer
       jkl = 0
       MSHFlexGrid1.Redraw = False
       For iLoop = MSHFlexGrid1.FixedRows To MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.Row = iLoop
            If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
            MSHFlexGrid1.RowHeight(iLoop) = ve_ro
            If jkl = 0 Then
                 jkl = iLoop
                 End If
            
            Else
            MSHFlexGrid1.RowHeight(iLoop) = 0
            End If
          Next
          MSHFlexGrid1.Row = jkl
           MSHFlexGrid1.Redraw = True
      End If
End Sub

Private Sub LaVolpeButton5_Click()
Dim iLoop As Integer
       If CatalogueName = "MATE" Then
       Dim jkl As Integer
       jkl = 0
       MSHFlexGrid1.Redraw = False
        For iLoop = MSHFlexGrid1.FixedRows To MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.Row = iLoop
            If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
                 MSHFlexGrid1.RowHeight(iLoop) = 0
                 Else
                 MSHFlexGrid1.RowHeight(iLoop) = ve_ro
                 If jkl = 0 Then
                 jkl = iLoop
                 End If
            End If
          Next
          MSHFlexGrid1.Row = jkl
          MSHFlexGrid1.Redraw = True
      End If
End Sub
Private Sub odknjiz()
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
    'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
     If CatalogueName = "MATE" Then
        frmControlMain.MSHFlexGrid1.Col = UR_id
        If Getnazi("select id_dok from zaloga where veza_td='" & tip_dok & "' and veza_id='" & frmControlMain.MSHFlexGrid1.Text & "'") <> "" Then
        Dim poknd As String
        poknd = ""
        
        MsgBox "Dokument je že imel izdajo zato ga ne morem odknjižit!! " & Getdo("select veza_td,veza_id from zaloga where veza_td='" & tip_dok & "' and veza_id='" & frmControlMain.MSHFlexGrid1.Text & "'")
        Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from zaloga where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", myConection, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
        rs.MoveFirst
        End If
        Do While Not rs.EOF
        
       myConection.Execute ("update zaloga set prosta=prosta-" & Replace(rs.Fields("kol"), ",", ".") & " where tip_dok='" & rs.Fields("veza_td") & "' and id_dok='" & rs.Fields("veza_id") & "' and poz=" & rs.Fields("poz"))
        
        rs.MoveNext
        Loop
        myConection.Execute ("DELETE FROM zaloga  where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
       myConection.Execute ("update nabasif set poknj='' where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
       myConection.Execute ("update nabasif set dat_k=#01/01/1899# where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
      '  If RS.State = 1 Then RS.Close
      '  RS.Open "select sifra from nabasif where  tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'", myConection, adOpenDynamic, adLockOptimistic
      '  RS.MoveFirst
      '  Do While Not RS.EOF
      '  Dim zall As Double
      '  zall = 0
      '  zall = Val(Getnazi("select sum(kol*faktor) as ss from nabasif where poknj='K' and sifra='" & RS.Fields("sifra") & "'"))
      '  myConection.Execute ("update mada set madazalo=" & Replace(zall, ",", ".") & " where madasifr='" & RS.Fields("sifra") & "'")
      '  RS.MoveNext
      '  Loop
        End If
    End If
Else
    MsgBox "Dokument še ni poknjižen!"
End If
End Sub
Private Sub LaVolpeButton6_Click()
frmControlMain.MSHFlexGrid1.Col = UR_id
If tip_dok = "NK" Then
Dim dass
Dim datum As String

dass = Format(Now, "dd.mm.yyyy")
datum = Left(dass, 2) & "/" & Mid(dass, 4, 2) & "/" & Mid(dass, 7, 4)
myConection.Execute ("update nabasif set poknj='K' where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
myConection.Execute ("update nabasif set dat_k=" & datum & " where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")

Exit Sub
End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='GEND'") <> "" Then
Dim adyda, adxda As Date
adyda = Getnazi("select datum from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
adxda = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='GEND'")
If adyda < adxda Then
MsgBox ("Generalni datum je nastavljen!")
Exit Sub
End If
End If

If intCtrlDown = 2 Then
odknjiz
intCtrlDown = 0
Else
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
MsgBox "Ta dokument je že poknjižen!"
Else
'preverim zalogo
If rs.State = 1 Then rs.Close
rs.Open "select sifra,sum(kol) as kol from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "' group by sifra", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Dim xarti As String
Dim fakkt As Long
xarti = ""
fakkt = 0
Dim Rsa As New ADODB.Recordset
Dim das, dodx
das = Format(Getnazi("select datum from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'"), "dd.mm.yyyy")

dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
  
Do While Not rs.EOF
If Rsa.State = 1 Then Rsa.Close
fakkt = Getnazi("select faktor from dokumenti where tip_dok='" & tip_dok & "'")
If fakkt <> 0 Then
Rsa.Open "select sum(prosta) as prosta from  zaloga where sifra='" & rs.Fields("sifra") & "' and datum<=#" & dodx & "#", myConection, adOpenDynamic, adLockOptimistic
End If
If Getnazi("select faktor from dokumenti where tip_dok='" & tip_dok & "'") < 0 Then
If IsNull(Rsa.Fields("prosta")) Then
xarti = xarti & rs.Fields("sifra") & ","

Else
If Rsa.Fields("prosta") < rs.Fields("kol") Then

xarti = xarti & rs.Fields("sifra") & ","
End If
End If
End If
If Getnazi("select faktor from dokumenti where tip_dok='" & tip_dok & "'") > 0 Then
If rs.Fields("kol") < 0 Then
If IsNull(Rsa.Fields("prosta")) Then
xarti = xarti & rs.Fields("sifra") & ","

Else
If Rsa.Fields("prosta") < rs.Fields("kol") * -1 Then

xarti = xarti & rs.Fields("sifra") & ","
End If
End If
End If
End If

rs.MoveNext
Loop
If xarti <> "" Then
MsgBox "Artikli " & xarti & " nimajo zaloge zato ne moreš poknjižiti!!!"
Exit Sub
End If
If tip_dok = "IN" Then
Else
 pre_fifo_vse tip_dok & frmControlMain.MSHFlexGrid1.Text
 End If
End If
End If
If tip_dok = "IN" Then
If MsgBox("Ali kreiram VIŠKE in MANJKE?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then
tip_dok = "NA"
id_inv = frmControlMain.MSHFlexGrid1.Text
NOVA_Click
frmblag.mviski
NOVA_Click
frmblag.mmanjki
End If
End If
osve = frmControlMain.MSHFlexGrid1.Row
End Sub
Sub bb()
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
MsgBox "Ta dokument je že poknjižen!"
Else
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
If CatalogueName = "MATE" Then
frmControlMain.MSHFlexGrid1.Col = UR_id
''preverimo èe ma zalogo
If Getnazi("select faktor from dokumenti where tip_dok='" & tip_dok & "'") = "-1" Then
If rs.State = 1 Then rs.Close
rs.Open "select * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", myConection, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Dim artikl As String
artikl = ""
Do While Not rs.EOF
If Getnazi("select prosta from zaloga where sifra='" & rs.Fields("sifra") & "' and prosta>0 order by datum") = "" Then
artikl = artikl & RTrim(rs.Fields("sifra")) & ","
End If
rs.MoveNext
Loop
If artikl <> "" Then
MsgBox artikl & " Ta artikel nima zaloge ne moreš poknjižiti!"
Exit Sub
End If
End If
myConection.Execute ("update nabasif set poknj='K' where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
Dim dass
Dim datum As String
Dim kirf As Long
kirf = Val(Getnazi("select faktor from dokumenti where tip_dok='" & tip_dok & "'"))

dass = Format(Now, "dd.mm.yyyy hh:mm:ss")
datum = Left(dass, 2) & "." & Mid(dass, 4, 2) & "." & Mid(dass, 7, 4) & " " & Mid(dass, 12, 2) & ":" & Mid(dass, 15, 2) & ":" & Mid(dass, 18, 2)
myConection.Execute ("update nabasif set dat_k='" & datum & "' where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'")
If kirf <> 0 Then
Dim rsz As New ADODB.Recordset
rsz.Open "select * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "' order by pozicija", myConection, adOpenDynamic, adLockOptimistic
Dim rsza As New ADODB.Recordset
rsza.Open "select * from zaloga where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", myConection, adOpenDynamic, adLockOptimistic

If Not rsz.EOF Then
Dim kolkse As Long
kolkse = 0
Dim td_dokum, id_dokum As String
rsz.MoveFirst

Do While Not rsz.EOF

If kirf = 1 Then
rsza.AddNew
rsza.Fields("sifra") = rsz.Fields("sifra")
rsza.Fields("naziv") = Left(rsz.Fields("naziv"), 50)
rsza.Fields("skl") = rsz.Fields("skl")
rsza.Fields("datum") = rsz.Fields("dat_k")
rsza.Fields("tip_dok") = tip_dok
rsza.Fields("id_dok") = frmControlMain.MSHFlexGrid1.Text
rsza.Fields("kol") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("prosta") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("cena") = FormatNumber(rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100)), 4)
rsza.Fields("vrednost") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor") * (rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100))), 4)
rsza.Update
Else

Dim kolkje, kolkbo, TOTA As Long
kolkje = 0
kolkbo = 0
TOTA = 0
'MsgBox rsz.Fields("kol")
Do While Not kolkje = rsz.Fields("kol")
If TOTA = 0 Then
If Val(Getnazi("select prosta from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum")) >= rsz.Fields("kol") Then

kolkje = rsz.Fields("kol")
kolkbo = kolkje

Else
kolkse = Val(Getnazi("select prosta from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum"))
kolkje = kolkje + kolkse
kolkbo = kolkse
End If
TOTA = 1
End If
If TOTA <> 0 Then
If Val(Getnazi("select prosta from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum")) >= rsz.Fields("kol") - kolkje Then

kolkje = rsz.Fields("kol")
kolkbo = kolkje

Else
kolkse = Val(Getnazi("select prosta from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum"))
kolkje = kolkje + kolkse
kolkbo = kolkse
End If

End If


rsza.AddNew
td_dokum = (Getnazi("select tip_dok from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum"))
id_dokum = (Getnazi("select id_dok from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum"))
rsza.Fields("sifra") = rsz.Fields("sifra")
rsza.Fields("naziv") = rsz.Fields("naziv")
rsza.Fields("skl") = rsz.Fields("skl")
rsza.Fields("datum") = rsz.Fields("dat_k")
rsza.Fields("tip_dok") = tip_dok
rsza.Fields("id_dok") = frmControlMain.MSHFlexGrid1.Text
rsza.Fields("kol") = FormatNumber(kolkbo * -1, 3)
rsza.Fields("prosta") = 0
rsza.Fields("cena") = FormatNumber((Getnazi("select cena from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum")), 4)
rsza.Fields("vrednost") = FormatNumber((Getnazi("select cena from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum")) * kolkbo * -1, 4)
rsza.Fields("veza_td") = td_dokum
rsza.Fields("veza_id") = id_dokum
rsza.Update
'MsgBox "update zaloga set poraba=poraba-" & kolkbo & " where tip_dok='" & td_dokum & "' and id_dok ='" & id_dokum & "'"
myConection.Execute ("update zaloga set prosta=prosta-" & kolkbo & " where tip_dok='" & td_dokum & "' and id_dok ='" & id_dokum & "' and sifra='" & rsz.Fields("sifra") & "' and cena=" & Replace(FormatNumber((Getnazi("select cena from zaloga where sifra='" & rsz.Fields("sifra") & "' and prosta>0 order by datum")), 4), ",", "."))

Loop

End If
'rsza.Update
If tip_dok = "NA" Then
Dim ssaa As String
ssaa = "update mada set madanabc=" & Replace(rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100)), ",", ".") & " where madasifr='" & rsz.Fields("sifra") & "'"
'MsgBox ssaa
myConection.Execute (Replace(ssaa, ",", "."))
ssaa = "update mada set madazalo=madazalo+" & Replace(rsz.Fields("kol"), ",", ".") & " where madasifr='" & rsz.Fields("sifra") & "'"
'MsgBox ssaa


myConection.Execute (ssaa)
End If
rsz.MoveNext
Loop
End If
End If

End If
End If
End If
If rs.State = 1 Then rs.Close
        rs.Open "select sifra from nabasif where  tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'", myConection, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
        rs.MoveFirst
        End If
        Do While Not rs.EOF
        Dim zall As Double
        zall = 0
        If Getnazi("select sum(kol*faktor) as ss from nabasif where poknj='K' and sifra='" & rs.Fields("sifra") & "'") <> "" Then
        zall = Getnazi("select sum(kol*faktor) as ss from nabasif where poknj='K' and sifra='" & rs.Fields("sifra") & "'")
        End If
        Dim sss As String
        sss = "update mada set madazalo=" & Replace(strVal(zall), ",", ".") & " where madasifr='" & rs.Fields("sifra") & "'"
      '  MsgBox SSS
        myConection.Execute (sss)
        rs.MoveNext
        Loop
        'End If
If rs.State = 1 Then rs.Close
rs.Open "select avg(cena) as cena,sifra from zaloga where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.Text & "' group by sifra"
If Not rs.EOF Then
'RS.MoveFirst
'Dim Rsa As New ADODB.Recordset
'Do While Not RS.EOF
'myConection.Execute ("update nabasif set cena=" & FormatNumber(RS.Fields("cena"), 4) & " where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "' and sifra='" & RS.Fields("sifra") & "'")
'If Rsa.State = 1 Then Rsa.Close
'Rsa.Open "select cena,znes,kol from nabasif  where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "' and sifra='" & RS.Fields("sifra") & "'", myConection, adOpenDynamic, adLockOptimistic
'Rsa.Fields("cena") = FormatNumber(RS.Fields("cena"), 4)
'Rsa.Fields("znes") = FormatNumber(Rsa.Fields("kol") * RS.Fields("cena"), 4)
'Rsa.Update
'myConection.Execute ("update nabasif set znes=cena*kol  where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "' and sifra='" & RS.Fields("sifra") & "'")

'RS.MoveNext
'Loop
End If
osve = frmControlMain.MSHFlexGrid1.Row
End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' we weren't dragging
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------
On Error Resume Next
'If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If Button = 2 Then
        
          MSHFlexGrid1.SetFocus
            frmMAIN.mnuEdit.Enabled = True
            frmMAIN.mnuModify.Enabled = True
            PopupMenu frmMAIN.mnuEdit
      Else
         
    
        xdn = X
        ydn = y
        m_iDragCol = -1     ' clear drag flag
        m_bDragOK = True
       
    End If
    

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    ' test to see if we should start drag
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' wrong button
    If m_iDragCol <> -1 Then Exit Sub                   ' already dragging
    If Abs(xdn - X) + Abs(ydn - y) < 50 Then Exit Sub   ' didn't move enough yet
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' must drag header

    ' if got to here then start the drag
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, _
     Shift As Integer)
      
    intCtrlDown = Shift
     If Shift = 2 Then
     LaVolpeButton6.Caption = "ODKNJIŽI"
     NOVA.Caption = "Kopiraj"
     kontr.Caption = "Prerac. FIFO"
     zalog.Caption = "Preraèun zalog"
     End If
     
    ' Select Case KeyCode

Select Case KeyCode

 Case vbKeyA To vbKeyZ
Dim sFind As String
'MsgBox (vmessql)
sFind = InputBox("Najdi zapis v " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)), "NAJDI ISKALNI NIZ", Chr(KeyCode))
sFind = Replace(sFind, "'", "", 1, Len(sFind), vbTextCompare)
Dim vmess As String
Dim whg, whgw As String
If InStr(vmessql, "where") <> 0 Then
vmessql = Replace(vmessql, "where", "where " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & " like '%" & sFind & "%' and ")
Else
vmessql = Replace(vmessql, "order ", "where " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & " like '%" & sFind & "%' order ")
End If
whg = " and " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & " like '" & sFind & "%' "
whgw = " where " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & " like '" & sFind & "%' "
If sFind <> "" Then
If Getnazi("select tekst from dokm where tip_dok='XX' and atribut='IMGL' and id_dok='" & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & "'") <> "" Then
sqt = Replace(whg, Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)), Getnazi("select tekst from dokm where tip_dok='XX' and atribut='IMGL' and id_dok='" & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & "'"))
Else
sqt = whg
End If
vmess = Replace(vmessql, "<and>", whg)
vmess = Replace(vmess, "<where>", whgw)
 If CatalogueName <> "MATE" Then
     
frmControlMain.Wbrow.Visible = False
'MsgBox vmessql
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(vmessql)
ssqq = vmess
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
  
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing
  frmControlMain.osv_Click

 End If
 Case Else
 End Select
End Sub
Private Sub MSHFlexGrid1_KeyUP(KeyCode As Integer, _
     Shift As Integer)

    intCtrlDown = 0
If Shift <> 2 Then
     LaVolpeButton6.Caption = "KNJIŽI"
     NOVA.Caption = "Nov"
      kontr.Caption = "KONTROLA"
      zalog.Caption = "Izpis zalog"
     End If

End Sub

Private Sub MSHFlexGrid1_dblClick()
'-------------------------------------------------------------------------------------------
' code in grid's DblClick event enables column sorting
'-------------------------------------------------------------------------------------------

    Dim i As Integer

    ' sort only when a fixed row is clicked
    If MSHFlexGrid1.MouseRow < MSHFlexGrid1.FixedRows Then

    i = m_iSortCol                  ' save old column
    m_iSortCol = MSHFlexGrid1.Col   ' set new column

    ' increment sort type
    If i <> m_iSortCol Then
        ' if clicking on a new column, start with ascending sort
        m_iSortType = 1
    Else
        ' if clicking on the same column, toggle between ascending and descending sort
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort
    Else
    Call UREDI_Click
    End If
End Sub


Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------

    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType

        .FillStyle = flexFillRepeat
        .Col = 0
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
      '  .CellBackColor = &HFFFFFF
        ' grey every other row
        Dim iLoop As Integer
      
       If CatalogueName = "MATE" Then
     
        For iLoop = .FixedRows To .Rows - 1
        Dim asx As String
        asx = MSHFlexGrid1.TextMatrix(iLoop, 1)
       
         .Row = iLoop
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
           ' MsgBox asx
            'MsgBox (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'"))
        If (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'")) = "K" Then
       If Getnazi("select placilo from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'") <> "" Then
        If (Getnazi("select placilo from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'")) = 1 Then
          
            .CellBackColor = &HFF00FF
        Else
            .CellBackColor = &HC0C0FF
        End If
        Else
            .CellBackColor = &HC0C0FF
        End If
            Else
            If Getnazi("select placilo from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'") <> "" Then
            If (Getnazi("select placilo from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'")) = 1 Then
            
             .CellBackColor = &HFF00&
            Else
             .CellBackColor = &HC0FFC0
            End If
            Else
             .CellBackColor = &HC0FFC0
             End If
       End If
        Next iLoop
        .FillStyle = flexFillSingle

        End If
        .Redraw = True
        
    End With

End Sub




Private Sub Command22_Click()
'dob.Show

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")
Select Case KeyCode

 Case vbKeyA To vbKeyZ
Dim sFind As String
'MsgBox (SQL)
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
        Call GRIDBIND("PurchaseOrderHeader", frmControlMain.MSHFlexGrid1, " Where PurchaseOrderID like'" & sFind & "%'")

  Case "Purchase Return"
            Call GRIDBIND("PurchaseReturnHeader", frmControlMain.MSHFlexGrid1, " Where PurchaseReturnID like'" & sFind & "%'")
  Case "Purchase Registry"
      Call GRIDBIND("PurchaseRegistryHeader", frmControlMain.MSHFlexGrid1, " Where PurchaseRegistryID like'" & sFind & "%'")
  Case "Sales Return"
            Call GRIDBIND("racusif", frmControlMain.MSHFlexGrid1, " Where st=" & sFind)
  Case "Sales Registry"
             Call GRIDBIND("tdr", frmControlMain.MSHFlexGrid1, "")
    
End Select

End If
Case Else
    End Select
End Sub




Private Sub Form_Unload(Cancel As Integer)
'UnHook
 Call WheelUnHook(Me.hwnd)
End Sub

Private Sub LoadFlexGridColumnWidths(ByVal flx As MSHFlexGrid, kira As String)
Dim i As Integer

If Not tip_dok = "" Then
    For i = 0 To flx.Cols - 1
        ' Get the column width. Use its current
        ' width as the default value.
        
        flx.ColWidth(i) = GetSetting( _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.ColWidth(i))
    Next i
    End If
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
Me.DATDO.Value = Date
Me.DATOD.Value = Date

 Hook Me.hwnd
 Call CMB1("users", "username1", Text1)
FillC_ Combo1, "select tvorba from dokumenti where tip_dok='" & tip_dok & "'"
Dim SqLargs As String

        SqLargs = "SELECT madasifr,madanazi,madazalo,madampcd From mada WHERE ((madazalo)<=0) and tip_art='xxx' Order by madazalo DESC"
    Call CreateStartPage(SqLargs)
  Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hwnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
    If TypeOf ctl Is MSHFlexGrid Then
      If IsOver(ctl.hwnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
     If TypeOf ctl Is DataGrid Then
      If IsOver(ctl.hwnd, Xpos, Ypos) Then DataGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Image1.Width = Me.ScaleWidth
    If MSHFlexGrid1.Visible = True Then
        MSHFlexGrid1.Move 0, Image1.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (Image1.Height + 300)
    End If
        Wbrow.Move 0, Image1.Height, Me.ScaleWidth, Me.ScaleHeight - (Image1.Height + 150) '- 100
End Sub

    
  

Sub CreateStartPage(strSqry As String)
On Error GoTo adder:
Dim Rs1 As New ADODB.Recordset
    Rs1.CursorLocation = adUseClient
    GetNewConnection2
    Call Rs1.Open(strSqry, DCON, adOpenForwardOnly, adLockReadOnly)
Dim i As Integer
Dim Data2 As Variant

Wbrow.Navigate2 "about:blank"
        Do While Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        With Wbrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write ("Dobrodošli na POS sistemu")
        .Write ("<table border=0 Width=100% height=80%>")
        .Write ("<tr><td valign=TOP width=80%><table Width=100% border=0>")
        'FIRST TITLE
        .Write ("<tr><td bgcolor=#B4C0DC Height=20>" & "Naziv")
        .Write ("<td bgcolor=#B15C0DC Height=20>" & "Zaloga")
        .Write ("<td bgcolor=#B15C0D0 Height=20>" & "Maloprodajna cena")
                ''DATA COLUMN
        While Rs1.EOF <> True
                .Write ("<tr><td><li><A href='ID?" & Rs1.Collect(0) & "'>")  ''' this thing here is so bull shit
                .Write (Rs1.Collect(1)) ''Product Status
                .Write ("<td> <font color=Red>**" & Rs1.Collect(2) & "</td>")
                If Rs1.Collect(0) <= 0 Then
                         .Write ("<td></td>")
                    Else
                                .Write ("<td>" & (Rs1.Collect(3)) & "</td>")
                End If
                .Write ("</a></li></td></tr>")
                Rs1.MoveNext
        Wend
                .Write ("</td></tr></table></td><td><td valign=TOP>")
       .Write ("</td></table></BODY></HTML>")
       Wbrow.Document.Script.Document.clear
        Wbrow.Document.Script.Document.Close
End With
adder:
Exit Sub
End Sub

Private Sub izpi_Click()
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
If CatalogueName = "MATE" Or CatalogueName = "FAXX" Or CatalogueName = "KOMP" Then
frmControlMain.MSHFlexGrid1.Col = UR_id
xid_dok = frmControlMain.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, frmControlMain.MSHFlexGrid1.Col)

End If
izpisi.Show
End Sub

Private Sub knj_Click()
'knjiz = frmControlMain.DataGrid1.Columns("tip_dok").text & frmControlMain.DataGrid1.Columns("sifrapart").text
'knji.Show
End Sub

Private Sub LaVolpeButton1_Click()
Dim sq As String
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
If Me.Text1.Text = "" Then
sq = "SELECT DISTINCTROW mada.MADAGRUP,mada.MADANAZI,Sum(nabasif.KOL* nabasif.mpc) as nabv,Sum(nabasif.KOL) AS [KOL], format(Sum(nabasif.ZNES),'fixed') AS [znesek]" _
& " FROM mada RIGHT JOIN nabasif ON mada.MADASIFR = nabasif.SIFRA" _
& " Where nabasif.tip_dok='PA' and nabasif.DATUM between #" & dod & "# And  #" & ddo & "#" _
& " GROUP BY  nabasif.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
& " order by mada.madagrup"
CreateH_Page sq, "X Pregled prodaje po grupah"
Else
sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(nabasif.KOL* nabasif.mpc) as nabv,Sum(nabasif.KOL) AS [KOL], format(Sum(nabasif.ZNES),'fixed') AS [znesek]" _
& " FROM mada RIGHT JOIN nabasif ON mada.MADASIFR = nabasif.SIFRA" _
& " Where nabasif.tip_dok='PA' and nabasif.uporabnik='" & Getnazi("select up from users where username1='" & Me.Text1.Text & "'") & "' and nabasif.DATUM between #" & dod & "# And #" & ddo & "#" _
& " GROUP BY  nabasif.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
& " order by mada.madagrup"
CreateH_Page sq, "X Pregled prodaje po zaposlenem: " & Me.Text1.Text
End If

 

End Sub

Private Sub LaVolpeButton2_Click()
Dim sq As String
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
sq = "select * from tdr "
'sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(nabasif.KOL) AS [KOL], Sum(nabasif.ZNES) AS [znesek]" _
'& " FROM mada RIGHT JOIN nabasif ON mada.MADASIFR = nabasif.SIFRA" _
'& " Where nabasif.DATUM between #" & dod & "# And #" & ddo & "#" _
'& " GROUP BY  nabasif.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
'& " order by mada.madagrup"
tdrr
 CreateH_Page sq, "TDR"
End Sub
Private Sub tdrr()
Dim da, datt As Date
Dim dod As String
Dim c As Integer
Me.UserControl21.opentime
Me.UserControl21.Visible = True
datt = Getnazi("select datum from nabasif where tip_dok='NA' order by datum")
c = 1
Dim rst As New ADODB.Recordset
Dim RSt1 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
myConection.Execute ("delete from tdr")
rst3.Open "select * from tdr", myConection, adOpenDynamic, adLockOptimistic
Yvs = 1
Xvs = Date - datt
Do While Not Date = datt
DoEvents
Yvs = Yvs + 1
das = Format(datt, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
'Set rst = dbs.OpenRecordset("select znes from proda where datum=#" & dod & "#")
If rst.State = 1 Then rst.Close
rst.Open "select sum(kol*mpc)as znes from nabasif where tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenDynamic, adLockOptimistic


If rst.EOF Then
'rst3.prodaja = 0
Else
rst3.AddNew
rst3.Fields("zap") = c
rst3.Fields("datum") = datt
rst3.Fields("prodaja") = Round(rst.Fields("znes"), 2)
rst3.Fields("opis") = "Dnevni iztržek"
rst3.Update
End If

If RSt1.State = 1 Then RSt1.Close
RSt1.Open "select tip_dok,id_dok,datum, sum(znes) as znesek from nabasif where tip_dok='NA' and datum=#" & dod & "# group by tip_dok,id_dok,datum order by datum,id_dok", myConection, adOpenDynamic, adLockOptimistic

If RSt1.EOF Then
Else
c = c + 1
RSt1.MoveFirst
Do While Not RSt1.EOF
rst3.AddNew
rst3.Fields("zap") = c
rst3.Fields("prodaja") = 0
rst3.Fields("datum") = datt
rst3.Fields("nabava") = Round(RSt1.Fields("znesek"), 2)
rst3.Fields("opis") = RSt1.Fields("tip_dok") + RTrim(RSt1.Fields("id_dok")) + " " + Getnazi("select dod2 from glavna where tip_dok='" & RSt1.Fields("tip_dok") & "' and id_dok='" & RSt1.Fields("id_dok") & "'")
c = c + 1
RSt1.MoveNext
rst3.Update
Loop

End If
datt = datt + 1
c = c + 1

Loop
rst3.MoveFirst
Dim skupi As Double
skupi = 0
Do While Not rst3.EOF
'rst3.Edit
If IsNull(rst3.Fields("prodaja")) Then
rst3.Fields("prodaja") = 0
End If
If IsNull(rst3.Fields("nabava")) Then
rst3.Fields("nabava") = 0
End If
skupi = skupi + Round(rst3.Fields("nabava"), 2) - Round(rst3.Fields("prodaja"), 2)
rst3.Fields("zaloga") = skupi
rst3.Update
rst3.MoveNext
Loop
Me.UserControl21.closetime
Me.UserControl21.Visible = False
End Sub
Private Sub n2_Click()
Dim sq As String
Dim das, des
Dim skupi As Double
Dim poglej As String
Dim kli As String
poglej = "SELECT nabasif.SIFRA FROM nabasif INNER JOIN mada ON nabasif.SIFRA = mada.MADASIFR Where (((mada.tip_art) = 'IZD') And ((nabasif.tip_dok) = 'NA')) Or (((mada.tip_art) = 'IZD') And ((nabasif.tip_dok) = 'IZ'))"
If Getnazi(poglej) <> "" Then
MsgBox "Imaš izdelke v materialnem delu (NA,IZ),Možnost napake!) Artikel:" & Getnazi(poglej)
End If
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
frmControlMain.MSHFlexGrid1.Col = UREJAJ
If MsgBox("Ali naredim preraèun FIFO samo za aritkel" & frmControlMain.MSHFlexGrid1.Text & "?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then
sq = "SELECT left(sifra,7) as sifra,tip_dok, id_dok as id_dokumenta,format(datum,'dd.mm.yyyy') as datum,   kol as skkol, format(cena,'#####.####') as madampcd, veza_td, veza_id,kol*0 as stanje,format(prosta,'fixed') as prosta,cena*prosta as vrednostz FROM zaloga where sifra='" & frmControlMain.MSHFlexGrid1.Text & "' and DATUM between #" & dod & "# And  #" & ddo & "# ORDER BY sifra, datum, veza_td,id_dok"
skupi = Getnumb("select sum(cena*IIf([veza_td]=[tip_dok],kol,prosta)) from FROM zaloga where sifra='" & frmControlMain.MSHFlexGrid1.Text & "' and DATUM between #" & dod & "# And  #" & ddo & "# ORDER BY sifra, datum, veza_td,id_dok")
kli = "0"
Else
sq = "SELECT left(sifra,7) as sifra,tip_dok, id_dok as id_dokumenta,format(datum,'dd.mm.yyyy') as datum,   kol as skkol, format(cena,'#####.####') as madampcd, veza_td, veza_id,kol*0 as stanje,format(prosta,'fixed') as prosta,cena*IIf([veza_td]=[tip_dok],kol,prosta) as vrednostz FROM zaloga where DATUM between #" & dod & "# And  #" & ddo & "# ORDER BY sifra, datum, veza_td,id_dok"
skupi = Getnumb("select sum(cena*IIf([veza_td]=[tip_dok],kol,prosta)) from FROM zaloga where  DATUM between #" & dod & "# And  #" & ddo & "# ORDER BY sifra, datum, veza_td,id_dok")
kli = "1"
End If
'& " FROM mada RIGHT JOIN nabasif ON mada.MADASIFR = nabasif.SIFRA" _
'& " Where nabasif.DATUM between #" & dod & "# And  #" & ddo & "#" _
'& " GROUP BY  nabasif.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
'& " order by mada.madagrup"
Cre_Page sq, " PREGLED FIFO KARTICE", skupi, kli

End Sub

Private Sub NOVA_Click()
MODIFYID = ""
If intCtrlDown = 2 Then
ma_ured = "1"
ma_ko = 1
dtip_dok = tip_dok
frmblag.Show

intCtrlDown = 0
Else
If CatalogueName = "Location" Then

    C_frmLocation.Show
End If
If CatalogueName = "posta" Then
idpo = ""
    VRPO.Show
End If
If CatalogueName = "FAXX" Then
idfx = ""
    FAX.Show
End If
If CatalogueName = "KOMP" Then
idko = ""
    Kompen.Show
End If

If CatalogueName = "DOKMI" Then

    Vr_po.Show
End If
If CatalogueName = "KUPEC" Then

    C_frmCustomer.Show
End If
If CatalogueName = "DOBAVITELJ" Then

    C_frmCustomer.Show
End If

If CatalogueName = "potni" Then
potni.Show
End If
If CatalogueName = "em" Then
xEM = ""
eme.Show
End If
If CatalogueName = "sklad" Then
xskladd = ""
skladisce.Show
End If
If CatalogueName = "tipa" Then
tipa = ""
tip_art.Show
End If

If CatalogueName = "Category" Then
frmProdEntry.ShowAdd
End If
If CatalogueName = "zaposleni" Then
zaposle = ""
zaposleni.Show
End If
If CatalogueName = "relacija" Then
relacij = ""
relacije.Show
End If
If CatalogueName = "avtom" Then
avtomob = ""
avtomobil.Show
End If
If CatalogueName = "MATE" Then
ma_ured = 0
If tip_dok = "PA" Then
frmsalesbill.Show
Else
frmblag.Show
End If
End If
End If
End Sub

Public Function i_dod(ux As Integer) As String
i_dod = ""
i_dod = UCase(Trim(Getnazi("select dod" & LTrim(str(ux)) & " from dokumenti where tip_dok='" & tip_dok & "'")))
i_dod = Replace(i_dod, "È", "C")
i_dod = Replace(i_dod, "Š", "S")
i_dod = Replace(i_dod, "Ž", "Z")
i_dod = Replace(i_dod, "Ð", "DJ")
i_dod = Replace(i_dod, "Æ", "C")

If i_dod = "" Then
i_dod = "DODATEK_" & LTrim(str(ux))
End If
End Function
Public Sub osv_Click()
Dim fds As Integer
fds = 0
For fds = 1 To 4
Me.StatusBar1.Panels(fds).Text = Getnazi("select id_dok from dokm where atribut='STAP' and poz=" & fds)
Next
frmControlMain.MSHFlexGrid1.Visible = True
'MsgBox tip_dok
Dim dods As String
dods = ""

Dim zz As Integer
If dejpre = 1 Then
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
dodajwh = "and nabasif.datum between #" & dod & "# AND #" & ddo & "# "

Else
dodajwh = ""
End If
For zz = 0 To 7
dods = dods & "glavna.dod" & LTrim(str(zz)) & " as " & i_dod(zz) & ","

If Getnazi("select id_dok from dokm where tip_dok='XX' and atribut='IMGL' and id_dok='" & i_dod(zz) & "'") = "" Then
myConection.Execute ("insert into dokm (tip_dok,id_dok,atribut,tekst) values ('XX','" & Left(i_dod(zz), 10) & "','IMGL','" & "glavna.dod" & LTrim(str(zz)) & "')")

End If

Next
If Getnazi("select id_dok from dokm where tip_dok='XX' and atribut='IMGL'") = "" Then
myConection.Execute ("insert into dokm (tip_dok,id_dok,atribut,tekst) values ('XX','Tujina','IMGL','placilo')")

End If
If tip_dok <> "PA" Then
SQL = "SELECT top " & Me.Text2.Text & " glavna.tip_dok, glavna.id_dok, " & dods & "glavna.opis, glavna.skl, Max(format(nabasif.datum,'dd.mm.yyyy')) AS datum,Sum(nabasif.kol) AS skkol, Max(nabasif.uporabnik) AS uporabnik, Max(nabasif.placilo) AS Tujina, Max(nabasif.poknj) AS poknj, format(sum(kol*(cena*(1-(pop/100)))),'fixed') as znesek " & _
      " FROM glavna LEFT JOIN nabasif ON (glavna.tip_dok = nabasif.tip_dok) AND (glavna.id_dok = nabasif.id_dok)" & _
    " Where glavna.tip_dok='" & tip_dok & "' <and> " & dodajwh & _
      " GROUP BY glavna.tip_dok, glavna.id_dok, glavna.faktor, glavna.dod0, glavna.dod1, glavna.dod2, glavna.dod3, glavna.dod4, glavna.dod5, glavna.dod6, glavna.dod7, glavna.opis, glavna.skl " & _
      " order by glavna.id_dok desc"
Else
SQL = "SELECT top " & Me.Text2.Text & " glavna.tip_dok, glavna.id_dok, glavna.opis, glavna.skl, Max(format(nabasif.datum,'dd.mm.yyyy')) AS datum,Sum(nabasif.kol) AS skkol, Max(nabasif.uporabnik) AS uporabnik, Max(nabasif.placilo) AS Tujina, Max(nabasif.poknj) AS poknj, format(sum(kol*(cena*(1-(pop/100)))),'fixed')  as znesek " & _
      " FROM glavna LEFT JOIN nabasif ON (glavna.tip_dok = nabasif.tip_dok) AND (glavna.id_dok = nabasif.id_dok)" & _
    " Where glavna.tip_dok='" & tip_dok & "' <and> " & dodajwh & _
      " GROUP BY glavna.tip_dok, glavna.id_dok, glavna.faktor, glavna.dod0, glavna.dod1, glavna.dod2, glavna.dod3, glavna.dod4, glavna.dod5, glavna.dod6, glavna.dod7, glavna.opis, glavna.skl " & _
      " order by glavna.id_dok desc"
End If
'SQL = "SELECT glavna.tip_dok, glavna.id_dok, glavna.faktor, glavna.dod0, glavna.dod1, glavna.dod2, glavna.dod3, glavna.dod4, glavna.dod5, glavna.dod6, glavna.dod7, glavna.opis, glavna.skl, Max(nabasif.uporabnik) AS uporabnik, Max(nabasif.poknj) AS poknj, Sum(nabasif.ZNES) AS ZNESEK " & _
'    " FROM glavna LEFT JOIN nabasif ON (glavna.tip_dok = nabasif.tip_dok) AND (glavna.id_dok = nabasif.id_dok) " & _
'    " GROUP BY glavna.tip_dok, glavna.id_dok, glavna.faktor, glavna.dod0, glavna.dod1, glavna.dod2, glavna.dod3, glavna.dod4, glavna.dod5, glavna.dod6, glavna.dod7, glavna.opis, glavna.skl "

'sqt = " and " & Trim(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Col)) & " like '" & sFind & "%' "

'If sFind <> "" Then
SQL = Replace(SQL, "<and>", sqt)
sqt = ""
'End If
Dim rsa1 As New ADODB.Recordset
'MsgBox SQL
If CatalogueName = "MATE" Then
'MsgBox SQL
rsa1.Open SQL, myConection, adOpenDynamic, adLockOptimistic

'Set rs1 = myConection.Execute(SQL)
'MsgBox rsa1.Fields("tip_dok")

'MsgBox SQL
'Call GetNewConnection2

If rsa1.EOF Then
    frmControlMain.MSHFlexGrid1.Visible = False
    
    erro = "1"
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = rsa1
     erro = ""
       ve_ro = MSHFlexGrid1.RowHeight(1)
      
End If
End If


Set rsa1 = Nothing
Set DCON = Nothing
cooznes = 0
COOZALO = 0
cooskkol = 0
coolldat = 0
UREJAJ = 0
UR_id = 0
If frmControlMain.MSHFlexGrid1.Visible = True Then
For i = MSHFlexGrid1.Col To MSHFlexGrid1.Cols - 1
        Dim asx As String
        
        asx = Trim(MSHFlexGrid1.TextMatrix(0, i))
        If UCase(asx) = "MADAZALO" Then
        COOZALO = i
        End If
         If UCase(asx) = "SKKOL" Then
        cooskkol = i
        End If
        If UCase(asx) = "DATUM" Then
        coolldat = i
        End If
        
        If UCase(asx) = "ZNESEK" Or UCase(asx) = "MADAMPCD" Or UCase(asx) = "MADANABC" Then
        cooznes = i
        End If
        If UCase(asx) = "SIFRA" Or UCase(asx) = "MADASIFR" Then
        UREJAJ = i
        End If
        If UCase(asx) = "ID_DOK" Then
        UR_id = i
        End If
Next
'If cooznes <> 0 Then
  Dim cenn As Double
  Dim ZAL, skkk As Double
        With MSHFlexGrid1
       ' MsgBox fgtrial.TextMatrix(lCount, coollce)
        .Redraw = False ' makes it about 10x faster !
       
        For lcount = .FixedRows To .Rows - 1
           'cena
          If MSHFlexGrid1.Rows > 1 Then
          ' MsgBox MSHFlexGrid1.TextMatrix(lCount, cooznes)
          If cooznes <> 0 Then
          If MSHFlexGrid1.TextMatrix(lcount, cooznes) = "" Then
          cenn = 0
          Else
         'cenn = MSHFlexGrid1.TextMatrix(lCount, cooznes)
       '  MsgBox cooznes
         '  cenn = Replace(IIf(MSHFlexGrid1.TextMatrix(lCount, cooznes) = "", "0", MSHFlexGrid1.TextMatrix(lCount, cooznes)), ".", ",")
          ' MsgBox MSHFlexGrid1.TextMatrix(lCount, cooznes)
           cenn = Replace(Replace(MSHFlexGrid1.TextMatrix(lcount, cooznes), "", ""), ".", ",")
           End If
             MSHFlexGrid1.TextMatrix(lcount, cooznes) = FormatNumber(cenn, 2)
             End If
             If cooskkol <> 0 Then
             If MSHFlexGrid1.TextMatrix(lcount, cooskkol) <> "" Then
             skkk = Replace(Replace(MSHFlexGrid1.TextMatrix(lcount, cooskkol), ",", ""), ".", ",")
             MSHFlexGrid1.TextMatrix(lcount, cooskkol) = FormatNumber(skkk, 2)
             End If
             End If
           
           If COOZALO <> 0 Then
'           ZAL = Replace(IIf(MSHFlexGrid1.TextMatrix(lCount, COOZALO) = "", "0", MSHFlexGrid1.TextMatrix(lCount, COOZALO)), ".", ",")
          ZAL = Replace(Replace(MSHFlexGrid1.TextMatrix(lcount, COOZALO), ",", ""), ".", ",")
            If ZAL <> "" Then
             MSHFlexGrid1.TextMatrix(lcount, COOZALO) = FormatNumber(ZAL, 2)
             End If
            End If
         If coolldat <> 0 Then
             If MSHFlexGrid1.TextMatrix(lcount, coolldat) <> "" Then
              'MSHFlexGrid1.TextMatrix(lCount, coolldat) = Format(MSHFlexGrid1.TextMatrix(lCount, coolldat), "long date")
             
             End If
             End If
              
            End If
        Next
       
        .ColAlignment(cooznes) = flexAlignRightCenter
        .ColAlignment(COOZALO) = flexAlignRightCenter
        .Redraw = True ' dont forget to do this !
        End With
'End If

FillC_ Combo1, "select tvorba from dokumenti where tip_dok='" & tip_dok & "'"
If erro = "" Then
 DoColumnSort

   Call FG_AutosizeCols(MSHFlexGrid1, Me, , , False)

 LoadFlexGridColumnWidths MSHFlexGrid1, tip_dok & CatalogueName
 MSHFlexGrid1.Row = 1
End If
End If
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
Dim iinn As Integer
'MSComctlLib
Dim nodde As MSComctlLib.node
'MsgBox Panel
'nodde = Getnazi("select tekst from dokm where atribut='STAP' and id_dok='" & Panel & "'")
iinn = Getnazi("select tekst from dokm where atribut='STAP' and id_dok='" & Panel & "'")
Call frmMAIN.beno_os(iinn)
End Sub

Private Sub Timer1_Timer()
If Me.MSHFlexGrid1.Visible = True Then
Me.Command1.Visible = True
Else
Me.Command1.Visible = False
End If
If ber = 1 Then
ber = 0
'ref
End If

If CatalogueName = "MATE" And tip_dok = "PA" Then
Me.LaVolpeButton1.Visible = True

Me.Text1.Visible = True

Else
Me.LaVolpeButton1.Visible = False

Me.Text1.Visible = False
End If
If CatalogueName = "Category" Then
Me.n2.Visible = True

'Me.text1.Visible = True

Else
Me.n2.Visible = False

'Me.text1.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
If intCtrlDown <> 2 Then
     LaVolpeButton6.Caption = "KNJIŽI"
     NOVA.Caption = "Nov"
      kontr.Caption = "KONTROLA"
      zalog.Caption = "Izpis zalog"
     End If

If tip_dok = "DN" Then
Me.lansiraj.Visible = True
Else
Me.lansiraj.Visible = False
End If
If osve <> 0 Then
osv_Click
If osve <> 1 Then
frmControlMain.MSHFlexGrid1.Row = osve
End If
osve = 0
End If
End Sub

Private Sub tvor_Click()
If tip_dok = "PA" Then
If MsgBox("Klasièni naèin ali Vojè naèin?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then
Form6.Show
Else
VOJKO.Show
End If
Else
frmControlMain.MSHFlexGrid1.Col = UR_id
knjiz = frmControlMain.MSHFlexGrid1.Text
knji.Show
End If

End Sub

Private Sub UREDI_Click()
If CatalogueName = "KUPEC" Or CatalogueName = "DOBAVITELJ" Then
frmControlMain.MSHFlexGrid1.Col = UREJAJ
End If
If CatalogueName = "MATE" Then
frmControlMain.MSHFlexGrid1.Col = UR_id
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
If MSHFlexGrid1.CellBackColor = &HC0C0FF Then
MsgBox "Ta dokument je že poknjižen!"
Else
ma_ured = "1"
frmblag.Show
End If
End If
If CatalogueName = "Location" Then

MODIFYID = frmControlMain.MSHFlexGrid1.Text
'Load C_frmLocation
C_frmLocation.Show
    
End If
If CatalogueName = "posta" Then

idpo = frmControlMain.MSHFlexGrid1.Text
'Load C_frmLocation
VRPO.Show
    
End If
If CatalogueName = "FAXX" Then
frmControlMain.MSHFlexGrid1.Col = UR_id

idfx = frmControlMain.MSHFlexGrid1.Text
'Load C_frmLocation
FAX.Show
    
End If
If CatalogueName = "KOMP" Then
frmControlMain.MSHFlexGrid1.Col = UR_id

idko = frmControlMain.MSHFlexGrid1.Text
'Load C_frmLocation
Kompen.Show
    
End If
If CatalogueName = "DOKMI" Then
MODIFYID = frmControlMain.MSHFlexGrid1.Text
'Load C_frmLocation
   Vr_po.Show
End If
If CatalogueName = "Category" Then
frmControlMain.MSHFlexGrid1.Col = UREJAJ
MODIFYID = frmControlMain.MSHFlexGrid1.Text
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "MADASIFR" Then
If Not MODIFYID = "" Then
Load frmProdEntry
    frmProdEntry.ShowEdit MODIFYID
End If

End If
If CatalogueName = "KUPEC" Then
MODIFYID = frmControlMain.MSHFlexGrid1.Text

    C_frmCustomer.Show
End If
If CatalogueName = "em" Then
xEM = frmControlMain.MSHFlexGrid1.Text
eme.Show
End If
If CatalogueName = "sklad" Then
xskladd = frmControlMain.MSHFlexGrid1.Text
skladisce.Show
End If
If CatalogueName = "tipa" Then
tipa = frmControlMain.MSHFlexGrid1.Text
tip_art.Show
End If
If CatalogueName = "DOBAVITELJ" Then
MODIFYID = frmControlMain.MSHFlexGrid1.Text

    C_frmCustomer.Show
End If
If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "sifra" Or frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "madasifr" Then
If CatalogueName = "Customer" Then
Load frmProdEntry
frmProdEntry.ShowEdit frmControlMain.MSHFlexGrid1.Text
End If
If CatalogueName = "zaposleni" Then
zaposle = frmControlMain.MSHFlexGrid1.Text
zaposleni.Show
End If
If CatalogueName = "avtom" Then
avtomob = frmControlMain.MSHFlexGrid1.Text
avtomobil.Show
End If
If CatalogueName = "relacija" Then
relacij = frmControlMain.MSHFlexGrid1.Text
relacije.Show
End If
End If

End Sub

Private Sub vklj_Click()
If dejpre = 1 Then
Me.vklj.Value = False
dejpre = 0
Else
dejpre = 1
Me.vklj.Value = True

End If
osv_Click
End Sub

Private Sub wbrow_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error GoTo adder:
    Dim pos As Integer
    Dim newString As String
    pos = InStr(URL, "?")
    
     If pos > 0 Then
        Cancel = True
        newString = (URL)
        
        newString = Replace(newString, "%20", " ", 1, Len(URL), vbTextCompare)
            SQL = "SELECT  madazacs,madazalo,madanazi from mada where madasifr='" & (newString)
        '    rptState = "Product Details "
    
        '"Select  *  from Product where ProductID='" & newString & "'"
        Select Case rptState
           
        Case "SalesRegistry"
                SQL = "Select * from nabasif where stdok='" & newString & "'"
        Case "PurchaseRegistry"
          SQL = "Select * , Quantity * rate as Amount from PurchaseRegistryDetail where PurchaseregistryID='" & newString & "'"
    '    Case Else
    '        Call CreateSubPage(SQL, rptState)
        End Select
       Call CreateSubPage(SQL, rptState)
        
    End If
    Exit Sub
adder:
     Exit Sub
End Sub

Sub CreateSubPage(strSqry As String, Title As String)
On Error Resume Next
Dim tempRs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim i As Integer
Dim Data2 As Variant

Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute(strSqry)
        
        Wbrow.Navigate2 "about:blank"
        Do While Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
     With Wbrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write (Title)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        For Each fld In Rs1.Fields
            .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
        Next fld
        'First row
        .Write ("</tr>")
        'Make Data Cells and Loop to Another Row
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
            If i Mod 2 <> 0 Then
                .Write ("<td>" & fld.Value & "</td>")
            Else
                .Write ("<td bgcolor=#CCCCC2>" & fld.Value & "</td>")
            End If
            Next fld
            .Write ("</tr>")
            Rs1.MoveNext
        Wend
            .Write ("</td></tr></table></BODY></HTML>")

        Wbrow.Document.Script.Document.clear
        Wbrow.Document.Script.Document.Close
End With
'adder:
End Sub
'

Private Sub wbrow_NavigateError(ByVal pDisp As Object, URL As Variant, frame As Variant, StatusCode As Variant, Cancel As Boolean)
    MsgBox URL
End Sub
Sub CreateDataPage(strSqry As String, titiles As String)
On Error Resume Next
Dim fld As ADODB.Field
Dim i As Integer
Dim J As Integer
Dim Data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    
    Wbrow.Navigate2 "about:blank"
    'Wbrow.Navigate2 "about:blank"
        Do While Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
    With Wbrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        For Each fld In Rs1.Fields
            .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
        Next fld
        'First row
            .Write ("<tr>")
        'Make Data Cells and Loop to Another Row
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
            
            If J = 0 Then
                .Write ("<td><A href='ID?" & fld.Value & "'>")  ''' this thing here is so bull shit
            Else
                .Write ("<td>")  ''' this thing here is so bull shit
            End If
            'If i Mod 2 <> 0 Then '' making facncy look here
                .Write ("" & fld.Value & "</a></td>")
                J = J + 1
            'Else
            '    .Write ("<td bgcolor=#CCCCC2>" & fld.Value & "</td>")
           ' End If
            Next fld
            .Write ("</tr>")
            J = 0
            Rs1.MoveNext
            
        Wend
            .Write ("</td></tr></table></BODY></HTML>")

        Wbrow.Document.Script.Document.clear
        Wbrow.Document.Script.Document.Close
End With
'adder:

Set DCON = Nothing

End Sub
'

Private Sub ref()
frmControlMain.Wbrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
   
     SQL = "Select madasifr,madanazi,madampcd,madazalo from mada"
     CatalogueName = "Category"
    

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
End If

Set Rs1 = Nothing
Set DCON = Nothing

End Sub

Private Sub XX_Click()
If rs.State = 1 Then rs.Close
 SQL = "SELECT nabasif.DATUM, nabasif.STDOK, nabasif.SIFRAPART, nabasif.SIFRA, nabasif.EMBALAZA, nabasif.KOL, nabasif.CENA, " & _
    " nabasif.ZNES, nabasif.pop, nabasif.tip_dok, nabasif.id_dok, nabasif.poknj, nabasif.faktor, nabasif.naziv, nabasif.SIFRAPLAC, " & _
    " nabasif.mpc, nabasif.x, nabasif.y, nabasif.uporabnik, nabasif.pozicija, nabasif.chk_fix, dokm.tekst " & _
    " FROM nabasif LEFT JOIN dokm ON (nabasif.id_dok = dokm.id_dok) AND (nabasif.tip_dok = dokm.tip_dok) and (nabasif.pozicija = dokm.atribut) " & _
    " where nabasif.tip_dok='" & tip_dok & "' and nabasif.id_dok='" & frmControlMain.MSHFlexGrid1.Text & "'"
   ' MsgBox SQL
  rs.Open SQL, myConection, adOpenDynamic, adLockOptimistic
  
  
  ' While RS.EOF = False
Set izpis.DataSource = rs
'izpis.Sections("Section2").Controls.Item("lbldaily").Caption = tip_dok
'izpis.Sections("Section2").Controls.Item("txtFamily").DataField = RS("id_dok").Name
izpis.Show
'RS.MoveNext
'Wend


End Sub

Private Sub zakl_Click()
Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
   ' frmPrinterSetUp.Show vbModal
   ' If QuitCommand Then
   '     Set cPrint = Nothing
   '     Exit Sub
   ' End If

    
SendToPrinter:
   picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    If Me.DATOD.Value = Me.DATDO.Value Then
    Me.DATDO.Value = Me.DATDO.Value + 1
    End If
 Do While Not Me.DATOD.Value = Me.DATDO.Value
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Rekapitulacija za dan:", 0.1, True
    cPrint.pPrint Format(Me.DATOD.Value, "dd/mm/yyyy"), 2.5, True
    cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
    
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If rs.State = 1 Then rs.Close
   
 Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where tip_dok='PA' and datum = #" & dod & "#", myConection, adOpenStatic, adLockOptimistic
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
Dim vsto, kart, gotov As Double
zne = 0
kart = 0
gotov = 0
ddva = 0
ddvb = 0
hrana = 0
pijaca = 0
cig = 0
storitve = 0
vsto = 0
Dim davek As Double
Dim vrsta As Integer
Do While Not rs.EOF
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

If rs.Fields("sifrapart") <> 0 Then
orr = orr + rs.Fields(0)
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
If Getnumb("select vr from grupa where sifra=" & vrsta) = 3 Then
storitve = storitve + rs.Fields(0)
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
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where tip_dok='PA' and datum=#" & dod & "#", myConection, adOpenStatic, adLockOptimistic
cPrint.pPrint "", 0.1, False
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



End If
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Zacetna st.rac. : " & ee, 0.1, False
    cPrint.pPrint "Konèna st.rac.  : " & ff, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Skupaj izdano raèunov : " & ff - ee + 1, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Skupaj znesek prodaje : " & zne, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
'   If pijaca <> 0 Then
'   cPrint.pPrint "Skupaj znesek storitev : " & pijaca, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If hrana <> 0 Then
'   cPrint.pPrint "Skupaj znesek hrane : " & hrana, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If cig <> 0 Then
'   cPrint.pPrint "Skupaj znesek cigaretov : " & cig, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If storitve <> 0 Then
'   cPrint.pPrint "Skupaj znesek storitev : " & storitve, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If kart <> 0 Then
'   cPrint.pPrint "Skupaj znesek kartic : " & kart, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If gotov <> 0 Then
'   cPrint.pPrint "Skupaj znesek gotovine : " & gotov, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
   
   If vsto <> 0 Then
  ' cPrint.pPrint "Skupaj znesek vstopnic : " & vsto, 0.1, False
  '  cPrint.pPrint "=======================================", 0.1, False
   End If
   
   
 '  If orr <> 0 Then
 '  cPrint.pPrint "Skupaj znesek Orginalov : " & orr, 0.1, False
 '   cPrint.pPrint "=======================================", 0.1, False
 '  End If
   
  cPrint.pPrint
  
  
  If rs.State = 1 Then rs.Close
   
rs.Open "select DISTINCTROW nabasif.sifra,sum(nabasif.kol) as [koli],mada.madagrup,mada.madanazi" _
& " FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR" _
& " where nabasif.tip_dok='PA' and mada.madagrup=10" _
& " group by mada.madagrup,nabasif.sifra,mada.madanazi", myConection, adOpenStatic, adLockOptimistic
'& " and racusif.org=0 and racusif.oseba='" & Me.Label3.Caption & "'"
If Not rs.EOF Then
rs.MoveFirst
End If
Do While Not rs.EOF
'cPrint.pPrint RS.Fields("madanazi"), 0.1, True
'cPrint.pRightJust RS.Fields("koli"), 3, True
  rs.MoveNext
'  cPrint.pPrint
Loop
  
      If ddva <> 0 Or ddvb <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddva <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddva / 1.2, "standard"), 1.2 * tiskdol, True
    cPrint.pRightJust " 20 %", 1.9 * tiskdol, True
    cPrint.pRightJust Format(ddva - (ddva / 1.2), "standard"), 3 * tiskdol, True
    cPrint.pRightJust Format(ddva, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
     If ddvb <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddvb / 1.085, "standard"), 1.2 * tiskdol, True
    cPrint.pRightJust "8.5 %", 1.9 * tiskdol, True
    cPrint.pRightJust Format(ddvb - (ddvb / 1.085), "standard"), 3 * tiskdol, True
    cPrint.pRightJust Format(ddvb, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
    End If
cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    
    
 'odrez
    cPrint.pPrint Chr(27) & Chr(105), 0.1, False
  'Print #1, Chr(27) & Chr(105)
  
    Me.DATOD.Value = Me.DATOD.Value + 1
    
    cPrint.pPrint
  Loop
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
End Sub

Private Sub zalnaa_Click()
Dim tString  As String
Dim rstx As New ADODB.Recordset
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
   ' frmPrinterSetUp.Show vbModal
   ' If QuitCommand Then
   '     Set cPrint = Nothing
   '     Exit Sub
   ' End If

   Dim skuu, zz As Integer
skuu = 0
skuu = Getnazi("select count(madasifr) as xx from mada ")
zz = 0
 
SendToPrinter:
   picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Pregled zalog za dan:", 0.1, True
    cPrint.pPrint Format(Me.DATDO.Value, "dd/mm/yyyy"), 2.5, True
    cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
    
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If rs.State = 1 Then rs.Close
   
Dim des

Dim dat_ixx As String
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dat_ixx = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
rs.Open "select madasifr,madanazi,madazalo,madadoza,madagrup,madasest,madanabc,madadoza from mada order by madagrup,madasifr", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Dim dat_i As String
Dim zal_i As Double
Dim dat_ii As Date
 Dim naa, proo As Double

dat_ii = Date - 30000
Dim zalo As Double
    cPrint.pPrint "=======================================", 0.1, False
  Dim grpa As Integer
  grpa = 0
  Dim VREDZ, vree, dozz, nab, prod As Double
  vree = 0
  Do While Not rs.EOF
  zz = zz + 1
    Me.ProgressBar.Visible = True
        If zz < skuu Then
        Me.ProgressBar.Value = zz / skuu * 100
        End If
  zal_i = 0
  dat_ii = Date - 30000
   If Getnazi("select datum from nabasif where sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K'") = "" Then
   Else
   dat_ii = Format(Getnazi("select datum  from nabasif where sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K' order by datum desc"), "dd/mm/yyyy")
   zal_i = Getnazi("select kol  from nabasif where sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K' order by datum desc")
   End If
 dat_i = RTrim(LTrim(str(Month(dat_ii)))) & "/" & RTrim(LTrim(str(Day(dat_ii)))) & "/" & RTrim(LTrim(str(Year(dat_ii))))
  
  VREDZ = 0

  If Round(rs.Fields("madagrup"), 0) <> grpa Then
      cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Grupa : " & rs.Fields("madagrup"), 0.1, True
    cPrint.pPrint Getnazi("select grupa from grupa where sifra=" & Round(rs.Fields("madagrup"), 0)), 1.5, True
        cPrint.pPrint "", 0.1, False
    End If
    cPrint.pPrint Round(LTrim(RTrim(rs.Fields("madasifr"))), 0), 0.1, True
    cPrint.pPrint Left(rs.Fields("madanazi"), 20), 1, True
    If Getnazi("select sum(kol*faktor) as xx from nabasif where sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#  and datum<=#" & dat_ixx & "#") = "" Then
    VREZD = 0
    Else
    dozz = (Getnazi("select madadoza from mada where madasifr='" & LTrim(RTrim(rs.Fields("madasifr"))) & "'"))
    'MsgBox dozz
    If Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "'  and datum>#" & dat_i & "#  and datum<=#" & dat_ixx & "#") = "" Then
    nab = 0
    Else
    
    nab = (Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#  and datum<=#" & dat_ixx & "#"))
    End If
    If Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "# and datum<=#" & dat_ixx & "#") = "" Then
    prod = 0
    Else
    prod = (Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#  and datum<=#" & dat_ixx & "#")) * dozz
    End If
    'MsgBox nab
    VREDZ = FormatNumber(nab + prod, 2)
    End If
     If Getnazi("select sifra from sestavi where sifra=" & LTrim(RTrim(rs.Fields("madasifr"))) & "") <> "" Then
     VREDZ = 0
     rs.Fields("madasest") = "D"
     Else
     rs.Fields("madasest") = ""
     End If
     
     If rstx.State = 1 Then rstx.Close
  rstx.Open "select * from sestavi where sifras=" & LTrim(RTrim(rs.Fields("madasifr"))) & "", myConection, adOpenDynamic, adLockOptimistic
  If Not rstx.EOF Then
  rstx.MoveFirst
  Do While Not rstx.EOF
   
  VREDZ = zal_i + VREDZ + ((Val(Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & rstx.Fields("sifra") & "' and datum>#" & dat_i & "# and datum<=#" & dat_ixx & "#")) * rstx.Fields("kol")))
  rstx.MoveNext
  Loop
  End If
  'If rs.Fields("MADADOZA") > 1 Then
  
 '  vree = vree + (VREDZ * rs.Fields("madANABC"))
 ' Else
  ' vree = vree + (VREDZ * rs.Fields("madANABC") * rs.Fields("MADADOZA"))
 'End If
  
  'vree = vree + Getnumb()
  '  rs.Fields("MADAZALO") = VREDZ
  '  MsgBox ("select sum(kol*mpc) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum<=#" & dat_ixx & "#")
  '  rs.Fields("MADAnabc") = FormatNumber(Getnumb("select cena from nabasif where tip_dok='NA' and sifra='" & ltrim(rtrim(rs.Fields("madasifr"))) & "' order by datum desc"), 2)
 proo = Getnumb("select sum(kol*mpc) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum<=#" & dat_ixx & "#")
 naa = Getnumb("select sum(kol*cena) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum<=#" & dat_ixx & "#")
    
    cPrint.pRightJust VREDZ, 3.5, True
    cPrint.pPrint "Nab.vr: " & FormatNumber(naa, 2), 3.7, True
    cPrint.pPrint "Pro.vr: " & FormatNumber(proo, 2), 5.1, True
    cPrint.pPrint "Vr.ZAL: " & FormatNumber(naa - proo, 2), 6.5, True
    
    cPrint.pPrint "", 0.1, False
 grpa = Round(rs.Fields("madagrup"), 0)
 rs.Update
  rs.MoveNext
  Loop
 proo = Getnumb("select sum(kol*mpc) as xx from nabasif where tip_dok='PA' and datum<=#" & dat_ixx & "#")
 naa = Getnumb("select sum(kol*cena) as xx from nabasif where tip_dok='NA' and datum<=#" & dat_ixx & "#")
 
  cPrint.pPrint "", 0.1, False
     cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "NABAVA =>Nab.vred:" & FormatNumber(naa, 2), 0.1, False
    cPrint.pPrint "PRODAJA=>Nab.vred:" & FormatNumber(proo, 2), 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "ZALOGA=>Nab.vred:" & FormatNumber(naa - proo, 2), 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    
    
 odrez
    
    cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing

End Sub

Private Sub zalna_Click()
Dim sq As String
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
sq = "select sifra,naziv,kol,vrednost from zaloga "
'sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(nabasif.KOL) AS [KOL], Sum(nabasif.ZNES) AS [znesek]" _
'& " FROM mada RIGHT JOIN nabasif ON mada.MADASIFR = nabasif.SIFRA" _
'& " Where nabasif.DATUM between #" & dod & "# And #" & ddo & "#" _
'& " GROUP BY  nabasif.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
'& " order by mada.madagrup"
xzallo
 CreateH_PageZAL sq, "ZALOGA"
End Sub
Sub xzallo()
Dim da, datt As Date
'Dim dod As String
Me.UserControl21.opentime
Me.UserControl21.Visible = True
Dim c As Integer
c = 1
Dim kkk As Double
Dim znex As Double
Dim kkkna As Double
Dim znexna As Double
Dim rst As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
Dim RSt1 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
myConection.Execute ("delete from zaloga")
rst3.Open "select * from zaloga", myConection, adOpenDynamic, adLockOptimistic
RSt1.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
RSt1.MoveFirst
Yvs = 1
Xvs = Getnumb("select count(madasifr) from mada")
Do While Not RSt1.EOF
Yvs = Yvs + 1
DoEvents
If rst.State = 1 Then rst.Close
If rsta.State = 1 Then rsta.Close

rst.Open "select sum(kol*mpc)as znes,sum(kol*embalaza) as koli from nabasif where tip_dok='PA' and  datum<=#" & ddo & "# and sifra='" & RSt1.Fields("madasifr") & "'", myConection, adOpenDynamic, adLockOptimistic
rsta.Open "select sum(kol*cena)as znes,sum(kol) as koli from nabasif where tip_dok='NA' and  datum<=#" & ddo & "# and sifra='" & RSt1.Fields("madasifr") & "'", myConection, adOpenDynamic, adLockOptimistic
kkk = 0
znex = 0
kkkna = 0
znexna = 0
rst3.AddNew
rst3.Fields("sifra") = RSt1.Fields("madasifr")
rst3.Fields("datum") = ddo
rst3.Fields("naziv") = RSt1.Fields("madanazi")
If rst.EOF Then
kkk = 0
znex = 0
Else

kkk = IIf(IsNull(rst.Fields("koli")), 0, rst.Fields("koli"))
znex = IIf(IsNull(rst.Fields("znes")), 0, rst.Fields("znes"))
End If
If rsta.EOF Then
kkkna = 0
znexna = 0
Else

kkkna = IIf(IsNull(rsta.Fields("koli")), 0, rsta.Fields("koli"))
znexna = IIf(IsNull(rsta.Fields("znes")), 0, rsta.Fields("znes"))
End If

rst3.Fields("kol") = kkkna - kkk
rst3.Fields("vrednost") = znexna - znex

rst3.Update
RSt1.MoveNext
Loop
rst3.MoveFirst
Do While Not rst3.EOF
If rst4.State = 1 Then rst4.Close
rst4.Open "select * from sestavi where sifra=" & Val(rst3.Fields("sifra")), myConection, adOpenDynamic, adLockOptimistic
If Not rst4.EOF Then
'myConection.Execute ("update sestavi set emba=" & Replace(rst4.Fields("kol") * rst3.Fields("kol"), ",", ".") & ",nabc=" & Replace(rst3.Fields("vrednost"), ",", ".") & " where sifra=" & (rst4.Fields("SIFRAS")) & "")
rst4.MoveFirst
Do While Not rst4.EOF
rst4.Fields("emba") = rst3.Fields("kol")
rst4.Fields("nabc") = rst3.Fields("vrednost")
rst4.fileds("enme") = str(Getnumb("select cena from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rst4.Fields("SIFRAS"))) & "' order by datum desc") * rst4.Fields("KOL"))
rst4.Update
rst4.MoveNext
Loop
rst3.Fields("kol") = 0
rst3.Fields("vrednost") = 0
rst3.Update
End If
rst3.MoveNext

Loop
If rst3.State = 1 Then rst3.Close
rst3.Open "select * from sestavi", myConection, adOpenDynamic, adLockOptimistic
If Not rst3.EOF Then
Do While Not rst3.EOF
If rst4.State = 1 Then rst4.Close
rst4.Open "select * from zaloga where sifra='" & Trim(str(rst3.Fields("sifras"))) & "'", myConection, adOpenDynamic, adLockOptimistic
If Not rst4.EOF Then
'myConection.Execute ("update sestavi set emba=" & Replace(rst4.Fields("kol") * rst3.Fields("kol"), ",", ".") & ",nabc=" & Replace(rst3.Fields("vrednost"), ",", ".") & " where sifra=" & (rst4.Fields("SIFRAS")) & "")
rst4.Fields("kol") = rst4.Fields("kol") + (rst3.Fields("kol") * rst3.Fields("emba"))
rst4.Fields("vrednost") = rst4.Fields("vrednost") + rst3.Fields("nabc")
rst4.Update
rst3.Fields("nabc") = 0
rst3.Fields("emba") = 0
rst3.Update
End If
rst3.MoveNext

Loop

End If
Me.UserControl21.closetime
Me.UserControl21.Visible = False
End Sub

Private Sub zalog_Click()
 Dim skuu, skuup, zz As Long
skuup = 0
skuu = 0
zz = 0
Xvs = 1
Me.UserControl21.opentime
Me.UserControl21.Visible = True
Dim rsta As New ADODB.Recordset
If rsta.State = 1 Then rsta.Close
rsta.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
rsta.MoveFirst

Do While Not rsta.EOF
myConection.Execute ("update nabasif set embalaza=" & Replace(rsta.Fields("madadoza"), ",", ".") & " where tip_dok='PA' and sifra='" & rsta.Fields("madasifr") & "' and embalaza=0")
Xvs = Xvs + 1
rsta.MoveNext
Loop
End If

 Dim printam As String
 printam = ""
 If MsgBox("Ali izpišem na mali tiskalnik?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then
   
printam = "D"
End If

If intCtrlDown = 2 Then
    Me.ProgressBar.Visible = True

Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)

skuup = Getnumb("select count(id_dok) as xxe from nabasif where tip_dok='PA' and DATUM between #" & dod & "# And  #" & ddo & "# ")
If rsta.State = 1 Then rsta.Close
rsta.Open "select sifra,datum,mpc from nabasif where  tip_dok='PA' and DATUM between #" & dod & "# And  #" & ddo & "# ", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
rsta.MoveFirst

Do While Not rsta.EOF
'If rsta.Fields("sifra") = "534" Then
'MsgBox (Getcena(rsta.Fields("sifra"), rsta.Fields("datum")))
'End If

rsta.Fields("mpc") = Getcena(rsta.Fields("sifra"), rsta.Fields("datum"))
rsta.Update
rsta.MoveNext
 zz = zz + 1
        If zz < skuup Then
        Me.ProgressBar.Value = zz / skuup * 100
        Me.StatusBar1.Panels(5).Text = Round(zz / skuup * 100, 2)
        End If
Loop




Yvs = 1

End If
Me.ProgressBar.Visible = False
Else
    Me.ProgressBar.Visible = True

Dim tString  As String
Dim rstx As New ADODB.Recordset
      
   ' frmPrinterSetUp.Show vbModal
   ' If QuitCommand Then
   '     Set cPrint = Nothing
   '     Exit Sub
   ' End If

skuu = Getnazi("select count(madasifr) as xx from mada ")
  If printam <> "" Then
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview

SendToPrinter:
   picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Pregled zalog za dan:", 0.1, True
    cPrint.pPrint Format(Date, "dd/mm/yyyy"), 2, True
    cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
   cPrint.pPrint "=======================================", 0.1, False
End If
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If rs.State = 1 Then rs.Close
   
 
rs.Open "select madasifr,madanazi,madazalo,madadoza,madagrup,madasest,madanabc from mada order by madagrup,madasifr", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Dim dat_i As String
Dim zal_i As Double
Dim dat_ii As Date
dat_ii = Date - 30000
Dim zalo As Double
 
  Dim grpa As Integer
  grpa = 0
  Dim VREDZ, dozz, nab, prod As Double
Yvs = 1
  Do While Not rs.EOF

  Yvs = Yvs + 1

  DoEvents
  

  zz = zz + 1
   Me.ProgressBar.Visible = True
        If zz < skuu Then
        Me.ProgressBar.Value = zz / skuu * 100
        End If
  zal_i = 0
  dat_ii = Date - 30000
   
 dat_i = RTrim(LTrim(str(Month(dat_ii)))) & "/" & RTrim(LTrim(str(Day(dat_ii)))) & "/" & RTrim(LTrim(str(Year(dat_ii))))
  
  VREDZ = 0
  If printam <> "" Then
  If Round(rs.Fields("madagrup"), 0) <> grpa Then
      cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Grupa : " & rs.Fields("madagrup"), 0.1, True
    cPrint.pPrint Getnazi("select grupa from grupa where sifra=" & Round(rs.Fields("madagrup"), 0)), 1.5, True
        cPrint.pPrint "", 0.1, False
    End If
    cPrint.pPrint Round(LTrim(RTrim(rs.Fields("madasifr"))), 0), 0.1, True
    cPrint.pPrint Left(rs.Fields("madanazi"), 16), 0.5, True
  End If
    If Getnazi("select sum(kol*faktor) as xx from nabasif where sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#") = "" Then
    VREZD = 0
    Else
    '*BEN dozz = (Getnazi("select madadoza from mada where madasifr='" & LTrim(RTrim(rs.Fields("madasifr"))) & "'"))
    dozz = rs.Fields("madadoza")
    'MsgBox dozz
   ' If Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "'  and datum>#" & dat_i & "#") = "" Then
   ' nab = 0
   ' Else
    
    nab = (Getnumb("select sum(kol*faktor) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#"))
   ' End If
    'If Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#") = "" Then
    'prod = 0
    'Else
    prod = (Getnumb("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' and datum>#" & dat_i & "#")) * dozz
    'End If
    'MsgBox nab
    VREDZ = FormatNumber(nab + prod, 2)
    End If
     If Getnazi("select sifra from sestavi where sifra=" & LTrim(RTrim(rs.Fields("madasifr"))) & "") <> "" Then
     VREDZ = 0
     rs.Fields("madasest") = "D"
     Else
     rs.Fields("madasest") = ""
     End If
     
     If rstx.State = 1 Then rstx.Close
  rstx.Open "select * from sestavi where sifras=" & LTrim(RTrim(rs.Fields("madasifr"))) & "", myConection, adOpenDynamic, adLockOptimistic
  If Not rstx.EOF Then
  rstx.MoveFirst
  
  
  Do While Not rstx.EOF
   
   VREDZ = zal_i + VREDZ + ((Val(Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & rstx.Fields("sifra") & "' and datum>#" & dat_i & "#")) * rstx.Fields("kol")))
  rstx.MoveNext
  Loop
  End If
     
    rs.Fields("MADAZALO") = VREDZ
    
    rs.Fields("MADAnabc") = FormatNumber(Getnumb("select cena from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(rs.Fields("madasifr"))) & "' order by datum desc"), 2)
    If printam <> "" Then
    cPrint.pRightJust VREDZ, 2.5 * tiskdol, True
    cPrint.pPrint "", 0.1, False
    End If
 grpa = Round(rs.Fields("madagrup"), 0)
 rs.Update
  rs.MoveNext
  
  
  Loop
 If printam <> "" Then
  cPrint.pPrint "", 0.1, False
     cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    
     cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)

    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
    End If
' odrez
If printam <> "" Then
    Else
Dim sq As String
sq = "select madasifr,madanazi,madagrup, madazalo, space(10) as popis from mada order by madagrup,madanazi"
 CreateH_Page sq, "ZALOGA"
End If
   
End If
Me.UserControl21.Visible = False
Me.UserControl21.closetime
End Sub

Private Sub odrez()
Open "be1.txt" For Output As #1
Print #1, Chr(27) & Chr(105)
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
Close #1
Call Shell("print /d:LPT1 be1.txt", 6)
   
End Sub
Private Sub pre_fifo_vse(katera As String)
Dim rsz As New ADODB.Recordset
Dim rsza As New ADODB.Recordset
Dim rrr As New ADODB.Recordset
Dim pozzz As String
Dim xcex As Double
Dim xffx As Integer
Dim bremk As Double
Open "c:\fifo.txt" For Output As #1
Dim skuu, zz As Integer
skuu = 0
bremk = 0
xffx = 1
xcex = 0
zz = 0
myConection.Execute ("delete from nabasif where kol=0")

If katera = "" Then
myConection.Execute ("update nabasif set cena=0 where faktor<0")
myConection.Execute ("update nabasif set znes=0 where faktor<0")
myConection.Execute ("delete from zaloga")
rsz.Open "select * from nabasif where faktor<>0  and poknj='K' order by datum,faktor*-1,tip_dok,id_dok,pozicija*-1 desc", myConection, adOpenDynamic, adLockOptimistic
rsza.Open "select * from zaloga", myConection, adOpenDynamic, adLockOptimistic
skuu = Getnazi("select count(sifra) as xx from nabasif where faktor<>0 and sifra<>'' and poknj='K' ")
Else
rsz.Open "select * from nabasif where tip_dok='" & Left(katera, 2) & "'  and sifra<>'' and id_dok='" & Mid(katera, 3) & "' order by pozicija*-1 desc", myConection, adOpenDynamic, adLockOptimistic
rsza.Open "select * from zaloga where tip_dok='" & Left(katera, 2) & "' and id_dok='" & Mid(katera, 3) & "'", myConection, adOpenDynamic, adLockOptimistic
myConection.Execute ("update nabasif set poknj='K' where tip_dok='" & Left(katera, 2) & "' and id_dok='" & Mid(katera, 3) & "'")
skuu = Getnazi("select count(sifra) as xx from nabasif where tip_dok='" & Left(katera, 2) & "'  and sifra<>'' and id_dok='" & Mid(katera, 3) & "'")
End If


If Not rsz.EOF Then
Dim kolkse, kolkje, kicen, kolko As Double
kolkse = 0
kolkje = 0
kolko = 0

Dim pozix As Integer
Dim rsta As New ADODB.Recordset
Dim td_dokum, id_dokum As String
rsz.MoveFirst

Do While Not rsz.EOF
pozzz = rsz.Fields("pozicija")
xffx = 1
xcex = 0
Print #1, ""
Print #1, " ______________________________________________________________________________"
zz = zz + 1
Me.ProgressBar.Visible = True
If zz < skuu Then
Me.ProgressBar.Value = zz / skuu * 100
End If
rsza.AddNew
rsza.Fields("sifra") = rsz.Fields("sifra")
rsza.Fields("naziv") = Left(rsz.Fields("naziv"), 50)
rsza.Fields("skl") = rsz.Fields("skl")
rsza.Fields("datum") = rsz.Fields("datum")
rsza.Fields("tip_dok") = rsz.Fields("tip_dok")
rsza.Fields("id_dok") = rsz.Fields("id_dok")

If rsz.Fields("faktor") > 0 Then
rsza.Fields("poz") = rsz.Fields("pozicija")
rsza.Fields("kol") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("prosta") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("cena") = FormatNumber(rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100)), 4)
rsza.Fields("vrednost") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor") * (rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100))), 4)


'BREMEPISI
If rsz.Fields("kol") * rsz.Fields("faktor") < 0 Then

If rsta.State = 1 Then rsta.Close
rsta.Open "select * from zaloga where prosta>=" & Replace(rsz.Fields("kol") * -1, ",", ".") & " and sifra='" & rsz.Fields("sifra") & "' AND CENA=" & Replace(Replace(rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100)), ".", ""), ",", ".") & " order by datum,tip_dok,id_dok,poz*-1 desc", myConection, adOpenDynamic, adLockOptimistic

If Not rsta.EOF Then
rsza.Fields("PROSTA") = 0
rsza.Fields("veza_td") = rsta.Fields("TIP_DOK")
rsza.Fields("veza_id") = rsta.Fields("ID_DOK")
rsza.Fields("poz") = rsta.Fields("POZ")
'rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") + rsz.Fields("KOL"), 3)

rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") + rsz.Fields("KOL"), 3)
rsta.Update

Else
If rsta.State = 1 Then rsta.Close
rsta.Open "select * from zaloga where prosta>0 and sifra='" & rsz.Fields("sifra") & "'  order by datum,tip_dok,id_dok,poz*-1 desc", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
rsza.Fields("PROSTA") = 0
rsza.Fields("veza_td") = rsta.Fields("TIP_DOK")
rsza.Fields("veza_id") = rsta.Fields("ID_DOK")
rsza.Fields("poz") = rsta.Fields("POZ")
'rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") + rsz.Fields("KOL"), 3)
If rsta.Fields("PROSTA") > rsz.Fields("KOL") * -1 Then
rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") + rsz.Fields("KOL"), 3)
rsta.Update
Else
bremk = (rsz.Fields("KOL") * -1) - rsta.Fields("PROSTA")

rsza.Fields("KOL") = rsta.Fields("PROSTA") * -1
rsza.Update
rsta.Fields("PROSTA") = 0
rsta.Update
Do While Not bremk = 0
rsza.AddNew
rsza.Fields("sifra") = rsz.Fields("sifra")
rsza.Fields("naziv") = Left(rsz.Fields("naziv"), 50)
rsza.Fields("skl") = rsz.Fields("skl")
rsza.Fields("datum") = rsz.Fields("datum")
rsza.Fields("tip_dok") = rsz.Fields("tip_dok")
rsza.Fields("id_dok") = rsz.Fields("id_dok")
rsza.Fields("poz") = rsz.Fields("pozicija")
rsza.Fields("kol") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("prosta") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor"), 3)
rsza.Fields("cena") = FormatNumber(rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100)), 4)
rsza.Fields("vrednost") = FormatNumber(rsz.Fields("kol") * rsz.Fields("faktor") * (rsz.Fields("cena") * (1 - (rsz.Fields("pop") / 100))), 4)

If rsta.State = 1 Then rsta.Close
rsta.Open "select * from zaloga where prosta>0 and sifra='" & rsz.Fields("sifra") & "'  order by datum,tip_dok,id_dok,poz*-1 desc", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
rsza.Fields("PROSTA") = 0
rsza.Fields("veza_td") = rsta.Fields("TIP_DOK")
rsza.Fields("veza_id") = rsta.Fields("ID_DOK")
rsza.Fields("poz") = rsta.Fields("POZ")
'rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") + rsz.Fields("KOL"), 3)
If rsta.Fields("PROSTA") > bremk Then
rsza.Fields("KOL") = rsta.Fields("PROSTA") * -1
rsza.Update
rsta.Fields("PROSTA") = FormatNumber(rsta.Fields("PROSTA") - bremk, 3)
bremk = 0
rsta.Update

Else
bremk = bremk - rsta.Fields("PROSTA")
rsza.Fields("KOL") = rsta.Fields("PROSTA") * -1
rsza.Update
rsta.Fields("PROSTA") = 0
rsta.Update

End If
End If
'MsgBox (bremk)
Loop

End If
End If

End If


End If
rsza.Update

Else
If rsta.State = 1 Then rsta.Close
rsta.Open "select * from zaloga where prosta>0 and sifra='" & rsz.Fields("sifra") & "' order by datum,tip_dok,id_dok,poz desc", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
kolkje = FormatNumber(rsta.Fields("prosta"), 3)
Else

MsgBox rsz.Fields("sifra") & " nima zaloge"
Close #1
Exit Sub
End If

kolkse = FormatNumber(rsz.Fields("kol"), 3)
If rsta.Fields("prosta") > rsz.Fields("kol") Then
kolkje = kolkse
End If
td_dokum = rsta.Fields("tip_dok")
id_dokum = rsta.Fields("id_dok")

kicen = FormatNumber(rsta.Fields("cena"), 4)
pozix = rsta.Fields("poz")
rsza.Fields("poz") = pozix
rsza.Fields("kol") = FormatNumber(kolkje * -1, 3)
rsza.Fields("prosta") = 0
rsza.Fields("cena") = FormatNumber(kicen * -1, 4)
rsza.Fields("vrednost") = FormatNumber(kolkje * kicen * -1, 4)
rsza.Fields("veza_td") = td_dokum
rsza.Fields("veza_id") = id_dokum
If rsta.Fields("prosta") < rsz.Fields("kol") Then
kolkse = kolkse - kolkje
Else
kolkse = 0
End If
rsza.Update
xcex = (rsza.Fields("cena"))
Dim sdsd As String
Print #1, rsz.Fields("sifra") & "   " & rsz.Fields("tip_dok") & rsz.Fields("id_dok") & "   " & str(rsz.Fields("kol"))

If rrr.State = 1 Then rrr.Close
rrr.Open "select * from zaloga where tip_dok='" & td_dokum & "' and id_dok='" & id_dokum & "' and poz=" & pozix, myConection, adOpenDynamic, adLockOptimistic

'rrr.Fields("prosta") = FormatNumber(rrr.Fields("prosta") - kolkje, 3)
rrr.Fields("prosta") = FormatNumber(rrr.Fields("prosta") - kolkje, 3)
rrr.Update
Print #1, td_dokum & id_dokum & "  " & pozix & "   " & str(rsza.Fields("kol"))
'sdsd = "update zaloga set prosta=(prosta-" & LTrim(Replace(Replace(kolkje, ".", ""), ",", ".")) & ") where tip_dok='" & td_dokum & "' and id_dok='" & id_dokum & "' and poz=" & pozix
'myConection.Execute (sdsd)

If kolkse <> 0 Then
Do While Not kolkse <= 0

If rsta.State = 1 Then rsta.Close
rsta.Open "select * from zaloga where prosta>0 and sifra='" & rsz.Fields("sifra") & "' order by datum,tip_dok,id_dok,poz desc", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
kolkje = rsta.Fields("prosta")
Else

MsgBox rsz.Fields("sifra") & " nima zaloge"
Close #1
Exit Sub
End If

If rsta.Fields("prosta") > kolkse Then

kolkje = kolkse
End If
kicen = FormatNumber(rsta.Fields("cena"), 4)
pozix = rsta.Fields("poz")
td_dokum = rsta.Fields("tip_dok")
id_dokum = rsta.Fields("id_dok")

rsza.AddNew

rsza.Fields("sifra") = rsz.Fields("sifra")
rsza.Fields("naziv") = Left(rsz.Fields("naziv"), 50)
rsza.Fields("skl") = rsz.Fields("skl")
rsza.Fields("datum") = rsz.Fields("datum")
rsza.Fields("tip_dok") = rsz.Fields("tip_dok")
rsza.Fields("id_dok") = rsz.Fields("id_dok")
'rsza.Fields("poz") = rsz.Fields("pozicija")
rsza.Fields("poz") = pozix
rsza.Fields("kol") = FormatNumber(kolkje * -1, 3)
rsza.Fields("prosta") = 0
rsza.Fields("cena") = FormatNumber(kicen * -1, 4)
rsza.Fields("vrednost") = FormatNumber(kolkje * kicen * -1, 4)
rsza.Fields("veza_td") = td_dokum
rsza.Fields("veza_id") = id_dokum
rsza.Update
xffx = xffx + 1
xcex = xcex + rsza.Fields("cena")
Dim sdsdd As String
If rrr.State = 1 Then rrr.Close
rrr.Open "select * from zaloga where tip_dok='" & td_dokum & "' and id_dok='" & id_dokum & "' and poz=" & pozix, myConection, adOpenDynamic, adLockOptimistic

'rrr.Fields("prosta") = FormatNumber(rrr.Fields("prosta") - kolkje, 3)
rrr.Fields("prosta") = FormatNumber(rrr.Fields("prosta") - kolkje, 3)
rrr.Update
'MsgBox (rrr.Fields("prosta"))
Print #1, td_dokum & id_dokum & "  " & pozix & "   " & str(rsza.Fields("kol"))

'sdsdd = "update zaloga set prosta=(prosta+" & LTrim(Replace(Replace(kolkje * -1, ".", ""), ",", ".")) & ") where tip_dok='" & td_dokum & "' and id_dok='" & id_dokum & "' and poz=" & pozix
'myConection.Execute (sdsdd)
If rsta.Fields("prosta") < kolkse Then
kolkse = kolkse - kolkje
Else
kolkse = 0
End If
'kolkse = 0

Loop
End If
End If

Print #1, " ______________________________________________________________________________"
Print #1, str(xcex) & "       " & str(xffx)
Print #1, ""

If rsz.Fields("faktor") < 0 Then
End If
rsz.MoveNext
Loop
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from NABASIF where FAKTOR<0", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Do While Not rs.EOF
If rrr.State = 1 Then rrr.Close
rrr.Open "select SUM(VREDNOST) AS CENA from ZALOGA where tip_dok='" & rs.Fields("tip_dok") & "' and id_dok='" & rs.Fields("id_dok") & "' and sifra='" & rs.Fields("sifra") & "' GROUP BY TIP_DOK,ID_DOK,SIFRA", myConection, adOpenDynamic, adLockOptimistic
If Not rrr.EOF Then
rrr.MoveFirst
rs.Fields("CENA") = FormatNumber(rrr.Fields("CENA") / Getnumb("SELECT SUM(KOL) AS XX FROM NABASIF where tip_dok='" & rs.Fields("tip_dok") & "' and id_dok='" & rs.Fields("id_dok") & "' and sifra='" & rs.Fields("sifra") & "' GROUP BY TIP_DOK,ID_DOK,SIFRA"), 4)
rs.Update
End If
'MsgBox "select avg(cena) as cc from zaloga where tip_dok='" & RS.Fields("tip_dok") & "' and id_dok='" & RS.Fields("id_dok") & "' and sifra='" & RS.Fields("sifra") & "'"
'myConection.Execute ("update nabasif set cena=" & Replace(FormatNumber(Getnazi("select avg(cena) as cc from zaloga where tip_dok='" & RS.Fields("tip_dok") & "' and id_dok='" & RS.Fields("id_dok") & "' and sifra='" & RS.Fields("sifra") & "'"), 4), ",", ".") & " where tip_dok='" & RS.Fields("tip_dok") & "' and id_dok='" & RS.Fields("id_dok") & "' and sifra='" & RS.Fields("sifra") & "'")

rs.MoveNext

Loop
myConection.Execute ("update nabasif set ZNES=(CENA*kol) where FAKTOR<0")
'MsgBox "Konèano"
osve = frmControlMain.MSHFlexGrid1.Row
Me.ProgressBar.Visible = False
Close #1

Dim rest As New ADODB.Recordset
Dim rstc As New ADODB.Recordset
If katera = "" Then
myConection.Execute ("update mada set madazalo=0")
If rs.State = 1 Then rs.Close
rs.Open "select sifra,sum(prosta) as zalo from zaloga group by sifra order by sifra", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If rstc.State = 1 Then rstc.Close
rstc.Open "select madazalo from mada where madasifr='" & rs.Fields("sifra") & "'", myConection, adOpenDynamic, adLockOptimistic
rstc.Fields("madazalo") = rs.Fields("zalo")
rstc.Update

rs.MoveNext
Loop
End If
Else

If rest.State = 1 Then rest.Close

rest.Open "select sifra from nabasif where tip_dok='" & Left(katera, 2) & "' and id_dok='" & Mid(katera, 3) & "' group by sifra", myConection, adOpenDynamic, adLockOptimistic
If Not rest.EOF Then
rest.MoveFirst
Do While Not rest.EOF
If rs.State = 1 Then rs.Close

rs.Open "select sifra,sum(prosta) as zalo from zaloga where sifra='" & rest.Fields("sifra") & "' group by sifra order by sifra", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
'RS.MoveFirst
If rstc.State = 1 Then rstc.Close
rstc.Open "select madazalo from mada where madasifr='" & rs.Fields("sifra") & "'", myConection, adOpenDynamic, adLockOptimistic
rstc.Fields("madazalo") = rs.Fields("zalo")
rstc.Update
'myConection.Execute ("update mada set madazalo=" & RS.Fields("zalo") & " where madasifr='" & RS.Fields("sifra") & "'")
Else
myConection.Execute ("update mada set madazalo=0 where madasifr='" & rs.Fields("sifra") & "'")

End If
rest.MoveNext
Loop
End If
End If
End Sub


