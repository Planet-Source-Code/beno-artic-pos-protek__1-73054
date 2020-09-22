VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
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
   Begin VB.CommandButton Command2 
      Caption         =   "frm"
      Height          =   375
      Left            =   7680
      TabIndex        =   27
      Top             =   600
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   8640
      Top             =   1680
   End
   Begin LVbuttons.LaVolpeButton tvor 
      Height          =   375
      Left            =   10080
      TabIndex        =   21
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
      MICON           =   "frmControlMain.frx":0000
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
      ItemData        =   "frmControlMain.frx":001C
      Left            =   9480
      List            =   "frmControlMain.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton XX 
      Caption         =   "Command2"
      Height          =   255
      Left            =   9600
      TabIndex        =   18
      Top             =   480
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMain.frx":0020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMain.frx":5F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlMain.frx":6BCE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton izpi 
      Height          =   735
      Left            =   2220
      TabIndex        =   14
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      MICON           =   "frmControlMain.frx":6CD8
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
      Picture         =   "frmControlMain.frx":6CF4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
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
      MICON           =   "frmControlMain.frx":6DEE
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
      MICON           =   "frmControlMain.frx":6E0A
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
      Format          =   70975489
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
      Format          =   70975489
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
      MICON           =   "frmControlMain.frx":6E26
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
      Height          =   735
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      MICON           =   "frmControlMain.frx":6E42
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
      Height          =   735
      Left            =   1125
      TabIndex        =   16
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      MICON           =   "frmControlMain.frx":6E5E
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
      Left            =   3360
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
      MICON           =   "frmControlMain.frx":6E7A
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
      DragIcon        =   "frmControlMain.frx":6E96
      Height          =   3840
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   6773
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   3360
      TabIndex        =   22
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
      MICON           =   "frmControlMain.frx":71A0
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
      TabIndex        =   23
      Top             =   960
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
      MICON           =   "frmControlMain.frx":71BC
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
      TabIndex        =   24
      Top             =   960
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
      MICON           =   "frmControlMain.frx":71D8
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
      TabIndex        =   25
      Top             =   960
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
      MICON           =   "frmControlMain.frx":71F4
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
      TabIndex        =   26
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
      MICON           =   "frmControlMain.frx":7210
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
      Height          =   1200
      Left            =   0
      Picture         =   "frmControlMain.frx":722C
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
' variables for data binding
Private datPrimaryRS As ADODB.Recordset

' variables for enabling column sort
Private m_iSortCol As Integer
Private m_iSortType As Integer

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub Command1_Click()
SaveFlexGridColumnWidths MSHFlexGrid1, CatalogueName

End Sub

Private Sub Command2_Click()
frmblag.Show
End Sub

Private Sub lansiraj_Click()
Dim norma As String
Dim koli As Long
Dim kol, XX, yy As Long
myConection.Execute "delete from normati"
norma = Trim(Getnazi("select sifra from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'"))
koli = Val(Getnazi("select kol from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'"))
XX = Val(Getnazi("select x from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'"))
yy = Val(Getnazi("select y from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'"))
Dim rst As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
rst.Open "select * from nabasif where tip_dok='NT' and id_dok='" & norma & "'", myConection, adOpenDynamic, adLockOptimistic
rst.MoveFirst
Dim sii, nazii As String

Dim fixx As Integer
rsta.Open "select * from normati", myConection, adOpenDynamic, adLockOptimistic
Do While Not rst.EOF
sii = rst.Fields("sifra")
nazii = rst.Fields("naziv")
kol = rst.Fields("kol")

fixx = rst.Fields("pop")
rsta.AddNew
rsta.Fields("sifr") = sii
rsta.Fields("naz") = nazii
If fixx = 0 Then
If Getnazi("select madaenme from mada where madasifr='" & sii & "'") = "KOM" Then
rsta.Fields("kol") = Round(kol * koli * (XX / 100) * (yy / 100), 0)
Else
rsta.Fields("kol") = Round(kol * koli * (XX / 100) * (yy / 100), 2)
End If
Else
rsta.Fields("kol") = kol * koli

End If
rsta.Update
rst.MoveNext
Loop

kosovni = 1
tip_dok = "IZ"
NOVA_Click
End Sub

Private Sub LaVolpeButton6_Click()
If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
myConection.Execute ("update nabasif set poknj='K' where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'")
End If
End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
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

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        ydn = Y
        m_iDragCol = -1     ' clear drag flag
        m_bDragOK = True
       
    End If
    

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    ' test to see if we should start drag
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' wrong button
    If m_iDragCol <> -1 Then Exit Sub                   ' already dragging
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' didn't move enough yet
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' must drag header

    ' if got to here then start the drag
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

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
        .CellBackColor = &HFFFFFF
        ' grey every other row
        Dim iLoop As Integer
       If CatalogueName = "Category" Then
       Else
        For iLoop = .FixedRows To .Rows - 1
        Dim asx As String
        asx = MSHFlexGrid1.TextMatrix(iLoop, 1)
         .Row = iLoop
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
           ' MsgBox asx
            'MsgBox (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'"))
        If (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'")) = "K" Then
       
           
            .CellBackColor = &HC0C0FF
            Else
            .CellBackColor = &HC0FFC0
            
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




Private Sub datod_Change()
dod = Me.datod.Value
Call fref(cst)

End Sub
Private Sub Form_Unload(cancel As Integer)
'UnHook
End Sub
 Public Sub MouseWheel(ByVal fwKeys As Long, ByVal zDelta As Long, ByVal Xpos As Long, _
    ByVal Ypos As Long)

   'put a label on your for to check changing values
   ' Label1.Caption = "Keys=" & fwKeys & " Delta=" & zDelta & " xPos=" & Xpos & " yPos=" & Ypos
   If zDelta > 0 Then
   SendKeys "{up}"
   Else
   SendKeys "{down}"
   End If
'then you can change toprow of flex grid accordingly
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

 Hook Me.hWnd
 Call CMB1("users", "username1", Text1)
FillC_ Combo1, "select tvorba from dokumenti where tip_dok='NA'"
Dim SqLargs As String

        SqLargs = "SELECT madasifr,madanazi,madazalo,madampcd From mada WHERE ((madazalo)<=0) Order by madazalo DESC"
    Call CreateStartPage(SqLargs)
 
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    End
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Image1.Width = Me.ScaleWidth
    If MSHFlexGrid1.Visible = True Then
        MSHFlexGrid1.Move 0, Image1.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (Image1.Height + 100)
    End If
        WBrow.Move 0, Image1.Height, Me.ScaleWidth, Me.ScaleHeight - (Image1.Height + 150) '- 100
End Sub

    
  

Sub CreateStartPage(strSqry As String)
On Error GoTo adder:
Dim Rs1 As New ADODB.Recordset
    Rs1.CursorLocation = adUseClient
    GetNewConnection2
    Call Rs1.Open(strSqry, DCON, adOpenForwardOnly, adLockReadOnly)
Dim i As Integer
Dim data2 As Variant
WBrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        With WBrow.Document
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
       WBrow.Document.Script.Document.clear
        WBrow.Document.Script.Document.Close
End With
adder:
Exit Sub
End Sub

Private Sub izpi_Click()
Dim doku As String
doku = ""
If tip_dok = "DO" Then
doku = "DOBAVNICA ST.: " & tip_dok & MSHFlexGrid1.text
End If
If tip_dok = "FA" Then
doku = "FAKTURA ST.: " & tip_dok & MSHFlexGrid1.text
End If
If tip_dok = "PR" Then
doku = "PREDRACUN ST.: " & tip_dok & MSHFlexGrid1.text
End If

'doku = doku & Trim(DataGrid1.Columns(0)).Text
'& DataGrid1.Columns(1).Text
 Const C0 = 1
  Const C1 = 1.1
  Const C2 = 2.2
  Const C3 = 7
  Dim memCurrentYt As Single, memCurrentYb As Single
Dim rst As New ADODB.Recordset
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    picPrinting.Visible = True
    DoEvents
    
    cPrint.pStartDoc
    If ConnectRS(myConection, rst, "select * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & MSHFlexGrid1.text & "'") = True Then
    End If
    '/* Print Cover page ********************************************************************
    cPrint.pPrintPicture LoadPicture(App.path & "\gaber.jpg"), 0.2, 0.2, 5, 1
    cPrint.pFontName
    cPrint.FontSize = 13
     cPrint.CurrentY = 1.5
    cPrint.pPrint "Kupec: ", 1, False
    cPrint.pPrint rst.Fields("sifrapart"), 1, True
    cPrint.pPrint Getnazi("select naziv from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint "", 1, False
    cPrint.pPrint Getnazi("select ulica from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint "", 1, False
    cPrint.pPrint Getnazi("select posta from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint Getnazi("select mesto from partner where sifra=" & rst.Fields("sifrapart")), 2, True
    
    cPrint.pPrintedDate False, 6
   ' cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, 3
    cPrint.FontSize = 24
    cPrint.FontBold = True
    cPrint.CurrentY = 2.3
    'cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0E0FF, , vbFSSolid
    cPrint.FontTransparent = True
    cPrint.pPrint doku, 1
       cPrint.FontSize = 12
         
    cPrint.CurrentY = 3
     cPrint.pPrint "Ident ", 0.5, True
     cPrint.pPrint "Naziv ", 1.1, True
     cPrint.pPrint "Kolicina ", 4.5, True
     cPrint.pPrint "Cena ", 5.4, True
     cPrint.pPrint "Rabat ", 6.3, True
     cPrint.pPrint "Vrednost ", 7.2, True
      cPrint.CurrentY = 3.2
     cPrint.pLine 0.4, 8
    cPrint.FontSize = 10
     cPrint.CurrentY = 3.4
    'baza
    cPrint.FontBold = False
Dim znes, ddva, ddvb As Long
znes = 0
ddva = 0
ddvb = 0
 If rst.EOF Then
    rst.MoveFirst
    End If
    Dim vRS As Integer
    vRS = 1
    Dim xopi
    xopi = Getnazi("select tekst from dokm where iddo='" & Trim(rst.Fields("tip_dok")) & rst.Fields("id_dok") & "' and atribut='opis'")
    Do While Not rst.EOF
    If rst.EOF Then
    Exit Do
    End If
    znes = znes + Round(rst.Fields("cena") * rst.Fields("kol"), 2)
    ddva = ddva + Round((rst.Fields("cena") * rst.Fields("kol") / 1.2), 2)
    cPrint.pPrint rst.Fields("sifra"), 0.5, True
     cPrint.pPrint Getnazi("select madanazi from mada where madasifr='" & rst.Fields("sifra") & "'"), 1.1, True
     cPrint.pPrint rst.Fields("kol"), 4.5, True
     cPrint.pRightJust rst.Fields("cena"), 5.8, True
     cPrint.pRightJust rst.Fields("pop"), 6.7, True
     cPrint.pRightJust Round(rst.Fields("cena") * rst.Fields("kol"), 2), 7.6, True
     cPrint.pPrint "", 1, False
     'MsgBox Getnazi("select tekst from dokm where iddo='" & Trim(rst.Fields("tip_dok")) & rst.Fields("id_dok") & "' and atribut='" & Trim(Str(vrs)) & "'")
     If Getnazi("select tekst from dokm where iddo='" & Trim(rst.Fields("tip_dok")) & rst.Fields("id_dok") & "' and atribut='" & Trim(Str(vRS)) & "'") <> "" Then
      cPrint.pPrint Getnazi("select tekst from dokm where iddo='" & Trim(rst.Fields("tip_dok")) & rst.Fields("id_dok") & "' and atribut='" & Trim(Str(vRS)) & "'"), 1, False
      
     'cPrint.pPrint "", 1, False
     End If
     vRS = vRS + 1
    rst.MoveNext
   ' MsgBox (cPrint.CurrentY)
    If cPrint.CurrentY > 8.5 Then
             cPrint.pPrint "", 1, False
            cPrint.pLine 0, 9
            cPrint.pPrint "Naslednja stran ====>>>", 3, False
           ' cPrint.pFooter
            cPrint.pNewPage
            
    
    '/* Print Cover page ********************************************************************
    cPrint.pPrintPicture LoadPicture(App.path & "\gaber.jpg"), 0.2, 0.2, 5, 1
    cPrint.pFontName
    cPrint.FontName = "arial"
    cPrint.FontSize = 12
     cPrint.CurrentY = 1.5
    cPrint.pPrint "Kupec: ", 1, False
    cPrint.pPrint rst.Fields("sifrapart"), 1, True
    cPrint.pPrint Getnazi("select naziv from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint "", 1, False
    cPrint.pPrint Getnazi("select ulica from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint "", 1, False
    cPrint.pPrint Getnazi("select posta from partner where sifra=" & rst.Fields("sifrapart")), 1.5, True
    cPrint.pPrint Getnazi("select mesto from partner where sifra=" & rst.Fields("sifrapart")), 2, True
    
    cPrint.pPrintedDate False, 6
   ' cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, 3
    cPrint.FontSize = 24
    cPrint.FontBold = True
    cPrint.CurrentY = 2.3
    'cPrint.pBox 1, , cPrint.GetPaperWidth - 2, , &HC0E0FF, , vbFSSolid
    cPrint.FontTransparent = True
    cPrint.pPrint doku, 1
       cPrint.FontSize = 12
         
    cPrint.CurrentY = 3
     cPrint.pPrint "Ident ", 0.5, True
     cPrint.pPrint "Naziv ", 1.1, True
     cPrint.pPrint "Kolicina ", 4.5, True
     cPrint.pPrint "Cena ", 5.4, True
     cPrint.pPrint "Rabat ", 6.3, True
     cPrint.pPrint "Vrednost ", 7.2, True
      cPrint.CurrentY = 3.2
     cPrint.pLine 0.4, 8
     cPrint.FontSize = 12
     cPrint.CurrentY = 3.4
           ' GoSub PrintHeader
        End If
    Loop
     'cPrint.CurrentY = 3.5
    cPrint.pPrint "", 0, False
    cPrint.pLine 0.4, 8
    cPrint.pPrint "", 1, False
    cPrint.pPrint "Za placilo: ", 6, True
       cPrint.pPrint znes, 7.2, True
      cPrint.pPrint "", 0, False
      cPrint.pLine 1, 4.3
     cPrint.pPrint "", 0, False
     cPrint.pPrint "Osnova      DDV(%)        Znes.DDV     Znesek", 1, True
     cPrint.pPrint "", 0, False
     cPrint.pPrint ddva, 1.1, True
     cPrint.pPrint "20 %", 2, True
     cPrint.pPrint znes - ddva, 2.9, True
     cPrint.pPrint znes, 3.7, True
      cPrint.pPrint "", 0, False
      cPrint.pLine 1, 4.3
    cPrint.FontBold = False
    cPrint.pPrint "", 1, False
    cPrint.pPrint "", 1, False
    cPrint.pPrint "", 1, False
    If xopi <> "" Then
    cPrint.pPrint xopi, 1, False
      End If
    cPrint.FontSize = 12

    '/* Two different ways to center a long text string
   
    cPrint.pPrint
    'cPrint.ForeColor = vbRed
    'cPrint.pCenter "Please look at Readme.htm for additional information."
    'cPrint.ForeColor = vbBlack

    cPrint.CurrentY = cPrint.GetPaperHeight - 1.5
    'cPrint.pPrintedDate True
    'cPrint.pNewPage
     picPrinting.Visible = False
    Screen.MousePointer = vbDefault
    
    cPrint.ReportTitle = Command1.Caption
    
    cPrint.pFooter
    cPrint.pEndDoc
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing
PrintHeader:
'     cPrint.ReportTitle = Command1.Caption

End Sub

Private Sub knj_Click()
'knjiz = frmControlMain.DataGrid1.Columns("tip_dok").text & frmControlMain.DataGrid1.Columns("sifrapart").text
'knji.Show
End Sub

Private Sub LaVolpeButton1_Click()
Dim sq As String
If Me.Text1.text = "" Then
sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(RACUSIF.KOL) AS [KOL], Sum(RACUSIF.ZNESEK) AS [znesek]" _
& " FROM mada RIGHT JOIN RACUSIF ON mada.MADASIFR = RACUSIF.SIFRA" _
& " Where RACUSIF.DATUM>=#" & Me.datod.Value & "# And RACUSIF.DATUM<=#" & Me.datdo.Value & "#" _
& " GROUP BY  RACUSIF.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
& " order by mada.madagrup"
CreateH_Page sq, "X Pregled prodaje po grupah"
Else
sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(RACUSIF.KOL) AS [KOL], Sum(RACUSIF.ZNESEK) AS [znesek]" _
& " FROM mada RIGHT JOIN RACUSIF ON mada.MADASIFR = RACUSIF.SIFRA" _
& " Where racusif.oseba='" & Me.Text1.text & "' and RACUSIF.DATUM>=#" & Me.datod.Value & "# And RACUSIF.DATUM<=#" & Me.datdo.Value & "#" _
& " GROUP BY  RACUSIF.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
& " order by mada.madagrup"
CreateH_Page sq, "X Pregled prodaje po zaposlemen: " & Me.Text1.text
End If

 

End Sub

Private Sub LaVolpeButton2_Click()
Dim sq As String
sq = "SELECT DISTINCTROW mada.MADAGRUP, mada.MADANAZI,Sum(RACUSIF.KOL) AS [KOL], Sum(RACUSIF.ZNESEK) AS [znesek]" _
& " FROM mada RIGHT JOIN RACUSIF ON mada.MADASIFR = RACUSIF.SIFRA" _
& " Where RACUSIF.DATUM>=#" & Me.datod.Value & "# And RACUSIF.DATUM<=#" & Me.datdo.Value & "#" _
& " GROUP BY  RACUSIF.SIFRA, mada.MADANAZI, mada.MADAGRUP" _
& " order by mada.madagrup"

 CreateH_Page sq, "X Pregled prodaje po grupah"
End Sub

Private Sub NOVA_Click()

If CatalogueName = "potni" Then
potni.Show
End If
If CatalogueName = "Customer" Then
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
frmblag.Show
End If
End Sub

Public Sub osv_Click()
 LoadFlexGridColumnWidths MSHFlexGrid1, CatalogueName
 DoColumnSort
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

If CatalogueName = "Sales Return" Then
Me.LaVolpeButton1.Visible = True

Me.Text1.Visible = True

Else
Me.LaVolpeButton1.Visible = False

Me.Text1.Visible = False
End If
End Sub

Private Sub Timer2_Timer()
If tip_dok = "DN" Then
Me.lansiraj.Visible = True
Else
Me.lansiraj.Visible = False
End If
If osve = 1 Then
osv_Click
osve = 0
End If
End Sub

Private Sub tvor_Click()
If Combo1.text <> "" Then
ma_ured = 2
dtip_dok = Trim(Combo1.text)
blag.Show
End If

End Sub

Private Sub UREDI_Click()
If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "id_dok" Then
ma_ured = "1"
frmblag.Show
End If

If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "sifra" Or frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "madasifr" Then
If CatalogueName = "Customer" Then
Load frmProdEntry
frmProdEntry.ShowEdit frmControlMain.MSHFlexGrid1.text
End If
If CatalogueName = "zaposleni" Then
zaposle = frmControlMain.MSHFlexGrid1.text
zaposleni.Show
End If
If CatalogueName = "avtom" Then
avtomob = frmControlMain.MSHFlexGrid1.text
avtomobil.Show
End If
If CatalogueName = "relacija" Then
relacij = frmControlMain.MSHFlexGrid1.text
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
End Sub

Private Sub wbrow_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, cancel As Boolean)
On Error GoTo adder:
    Dim pos As Integer
    Dim newString As String
    pos = InStr(URL, "?")
    
     If pos > 0 Then
        cancel = True
        newString = (URL)
        
        newString = Replace(newString, "%20", " ", 1, Len(URL), vbTextCompare)
            SQL = "SELECT  madazacs,madazalo,madanazi from mada where madasifr=" & Val(newString)
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

Sub CreateSubPage(strSqry As String, title As String)
On Error Resume Next
Dim tempRs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim i As Integer
Dim data2 As Variant

Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute(strSqry)
        
        WBrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
     With WBrow.Document
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write (title)
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

        WBrow.Document.Script.Document.clear
        WBrow.Document.Script.Document.Close
End With
'adder:
End Sub
'

Private Sub wbrow_NavigateError(ByVal pDisp As Object, URL As Variant, frame As Variant, StatusCode As Variant, cancel As Boolean)
    MsgBox URL
End Sub
Sub CreateDataPage(strSqry As String, titiles As String)
On Error Resume Next
Dim fld As ADODB.Field
Dim i As Integer
Dim J As Integer
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    
    WBrow.Navigate2 "about:blank"
    'Wbrow.Navigate2 "about:blank"
        Do While WBrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
    With WBrow.Document
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

        WBrow.Document.Script.Document.clear
        WBrow.Document.Script.Document.Close
End With
'adder:

Set DCON = Nothing

End Sub
'

Private Sub ref()
frmControlMain.WBrow.Visible = False
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
If RS.State = 1 Then RS.Close
  RS.Open "SELECT *,space(249) as opis from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & frmControlMain.MSHFlexGrid1.text & "'", myConection, adOpenDynamic, adLockOptimistic
  RS.MoveFirst
  Do While Not RS.EOF
 RS.Fields("opis") = Left(Getnazi("select tekst from dokm where iddo='" & Trim(RS.Fields("tip_dok")) & Trim(RS.Fields("id_dok")) & "'"), 245)
  RS.Update
  RS.MoveNext
 Loop
  RS.MoveFirst
  
   While RS.EOF = False
Set izpis.DataSource = RS
izpis.Sections("glava").Controls.Item("label3").Caption = Date
izpis.Show
RS.MoveNext
Wend


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
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Rekapitulacija za dan:", 0.1, True
    cPrint.pPrint Format(Me.datdo.Value, "dd/mm/yyyy"), 2.5, True
    cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
    
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If RS.State = 1 Then RS.Close
   
 
RS.Open "select znesek,sifra,vst from racusif where datum=#" & Me.datdo.Value & "#", myConection, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
RS.MoveFirst
End If

Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim orr As Double
Dim hrana As Double
Dim pijaca As Double
Dim cig As Double
Dim vsto As Double
zne = 0
ddva = 0
ddvb = 0
hrana = 0
pijaca = 0
cig = 0
storitve = 0
vsto = 0
Dim davek As Double
Dim vrsta As Integer
Do While Not RS.EOF
If RS.Fields("vst") <> 0 Then
orr = orr + RS.Fields(0)
End If
If IsNull(RS.Fields("sifra")) Then
RS.Fields("sifra") = 0
RS.Update
End If
vrsta = Getnazi("select madagrup from mada where madasifr=" & Round(RS.Fields("sifra"), 0))
If Getnazi("select vr from grupa where sifra=" & vrsta) = 0 Then
pijaca = pijaca + RS.Fields(0)
End If
If Getnazi("select vr from grupa where sifra=" & vrsta) = 1 Then
hrana = hrana + RS.Fields(0)
End If
If Getnazi("select vr from grupa where sifra=" & vrsta) = 2 Then
cig = cig + RS.Fields(0)
End If
If Getnazi("select vr from grupa where sifra=" & vrsta) = 3 Then
storitve = storitve + RS.Fields(0)
End If
If Getnazi("select vr from grupa where sifra=" & vrsta) = 4 Then
vsto = vsto + RS.Fields(0)
End If



zne = zne + RS.Fields(0)
If Getnazi("select madapd from mada where madasifr=" & Round(RS.Fields("sifra"), 0)) = "20" Then
ddva = ddva + RS.Fields(0)
End If
If Getnazi("select madapd from mada where madasifr=" & Round(RS.Fields("sifra"), 0)) = "8.5" Then
ddvb = ddvb + RS.Fields(0)
End If

RS.MoveNext
Loop
Dim aa As Double
Dim bb As Double
'aa = Format(Me.Text1.text, "0.00")
'bb = Format(Me.Text2.text, "0.00")

If RS.State = 1 Then RS.Close
   
RS.Open "select min(st) as minst, max(st) as maxst from racusif where datum=#" & Me.datdo.Value & "#", myConection, adOpenStatic, adLockOptimistic
cPrint.pPrint "", 0.1, False
Dim ee As Integer
Dim ff As Integer
ee = 0
ff = 0

If Not RS.EOF Then

If IsNull(RS.Fields(0)) Then
ee = 0
Else
ee = RS.Fields(0)
End If
If IsNull(RS.Fields(1)) Then

ff = 0
Else
ff = RS.Fields(1)
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
   If pijaca <> 0 Then
   cPrint.pPrint "Skupaj znesek pijace : " & pijaca, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   If hrana <> 0 Then
   cPrint.pPrint "Skupaj znesek hrane : " & hrana, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   If cig <> 0 Then
   cPrint.pPrint "Skupaj znesek cigaretov : " & cig, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   If storitve <> 0 Then
   cPrint.pPrint "Skupaj znesek storitev : " & storitve, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   
   If vsto <> 0 Then
   cPrint.pPrint "Skupaj znesek vstopnic : " & vsto, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   
   
   If orr <> 0 Then
   cPrint.pPrint "Skupaj znesek Orginalov : " & orr, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   
  cPrint.pPrint
  
  
  If RS.State = 1 Then RS.Close
   
RS.Open "select DISTINCTROW racusif.sifra,sum(racusif.kol) as [koli],mada.madagrup,mada.madanazi" _
& " FROM RACUSIF LEFT JOIN mada ON RACUSIF.SIFRA = mada.MADASIFR" _
& " where mada.madagrup=10" _
& " group by mada.madagrup,racusif.sifra,mada.madanazi", myConection, adOpenStatic, adLockOptimistic
'& " and racusif.org=0 and racusif.oseba='" & Me.Label3.Caption & "'"
If Not RS.EOF Then
RS.MoveFirst
End If
Do While Not RS.EOF
'cPrint.pPrint RS.Fields("madanazi"), 0.1, True
'cPrint.pRightJust RS.Fields("koli"), 3, True
  RS.MoveNext
  cPrint.pPrint
Loop
  
      If ddva <> 0 Or ddvb <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddva <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddva / 1.2, "standard"), 1.2, True
    cPrint.pRightJust " 20 %", 1.9, True
    cPrint.pRightJust Format(ddva - (ddva / 1.2), "standard"), 3, True
    cPrint.pRightJust Format(ddva, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
     If ddvb <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddvb / 1.085, "standard"), 1.2, True
    cPrint.pRightJust "8.5 %", 1.9, True
    cPrint.pRightJust Format(ddvb - (ddvb / 1.085), "standard"), 3, True
    cPrint.pRightJust Format(ddvb, "standard"), 4, True
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
    
    
 odrez
    
    cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
End Sub

Private Sub zalog_Click()
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
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Pregled zalog za dan:", 0.1, True
    cPrint.pPrint Format(Date, "dd/mm/yyyy"), 2.5, True
    cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
    
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If RS.State = 1 Then RS.Close
   
 
RS.Open "select madasifr,madanazi,madazalo,madagrup from mada order by madagrup,madasifr", myConection, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
RS.MoveFirst
End If

Dim zalo As Double
    cPrint.pPrint "=======================================", 0.1, False
  Dim grpa As Integer
  grpa = 0
  Do While Not RS.EOF
  If Round(RS.Fields("madagrup"), 0) <> grpa Then
      cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Grupa : " & RS.Fields("madagrup"), 0.1, True
    cPrint.pPrint Getnazi("select grupa from grupa where sifra=" & Round(RS.Fields("madagrup"), 0)), 1, True
        cPrint.pPrint "", 0.1, False
    End If
    cPrint.pPrint Round(RS.Fields("madasifr"), 0), 0.1, True
    cPrint.pPrint Left(RS.Fields("madanazi"), 20), 1, True
    cPrint.pRightJust Format(Round(RS.Fields("madazalo"), 2), "standard"), 3.5, True
    cPrint.pPrint "", 0.1, False
 grpa = Round(RS.Fields("madagrup"), 0)
  RS.MoveNext
  Loop
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
    
    
 odrez
    
    cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing

End Sub

Private Sub odrez()
Open "be1.txt" For Output As #1
Print #1, Chr(27) & Chr(105)
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
Close #1
Call Shell("print /d:LPT1 be1.txt", 6)
   
End Sub

