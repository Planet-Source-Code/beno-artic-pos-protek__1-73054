VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmblag 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DOKUMENT"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   15465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10080
      TabIndex        =   48
      Text            =   "Text4"
      Top             =   120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IZVOZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   46
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12120
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   255
      Left            =   12840
      TabIndex        =   42
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   2
      TX              =   "OSV.CENE"
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
      MICON           =   "frmblag.frx":0000
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
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10800
      TabIndex        =   40
      Text            =   "0,00"
      Top             =   3840
      Width           =   975
   End
   Begin LVbuttons.LaVolpeButton cmdadd 
      Height          =   735
      Left            =   360
      TabIndex        =   28
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Dodaj"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":001C
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
   Begin ProsVent.UserControl1 sklad 
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "skladisce"
      ssql            =   "select * from skla"
      TextLocked      =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71106561
      CurrentDate     =   39507
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9000
      Top             =   600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   3
      Top             =   3840
      Width           =   7335
   End
   Begin VB.TextBox txtNewData 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1060
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13560
      Top             =   2880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TrialGrd"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TrialGrd"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "GrdData"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   4
      Left            =   8640
      TabIndex        =   6
      Top             =   1920
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   5
      Left            =   8640
      TabIndex        =   7
      Top             =   2400
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   6
      Left            =   8640
      TabIndex        =   8
      Top             =   2880
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   7
      Left            =   8640
      TabIndex        =   9
      Top             =   3360
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0
      locked          =   0
   End
   Begin LVbuttons.LaVolpeButton cmddel 
      Height          =   735
      Left            =   2160
      TabIndex        =   29
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Briši"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":0038
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
   Begin LVbuttons.LaVolpeButton Isc 
      Height          =   735
      Left            =   5880
      TabIndex        =   30
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Iskanje"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":0054
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
   Begin LVbuttons.LaVolpeButton Uvoz 
      Height          =   735
      Left            =   6960
      TabIndex        =   31
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Uvoz"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":0070
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
   Begin LVbuttons.LaVolpeButton Zapis 
      Height          =   735
      Left            =   11880
      TabIndex        =   32
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Zapiši"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":008C
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
   Begin LVbuttons.LaVolpeButton prekin 
      Height          =   735
      Left            =   13560
      TabIndex        =   33
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Prekini"
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
      COLTYPE         =   3
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":00A8
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgtrial 
      DragIcon        =   "frmblag.frx":00C4
      Height          =   5040
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   8890
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
   Begin LVbuttons.LaVolpeButton opiss 
      Height          =   495
      Left            =   600
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "OPIS"
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
      COLTYPE         =   2
      BCOL            =   14737632
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmblag.frx":03CE
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
   Begin ProsVent.UserControl1 vskla 
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   1490
      Visible         =   0   'False
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      ssql            =   "select * from skla"
      polje           =   "skladisce"
      textlocked      =   0
      locked          =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   10380
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   11819
            MinWidth        =   11819
            Picture         =   "frmblag.frx":03EA
            Text            =   "Artikel"
            TextSave        =   "Artikel"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Zaloga"
            TextSave        =   "Zaloga"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "0,00"
            TextSave        =   "0,00"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Prosta zaloga"
            TextSave        =   "Prosta zaloga"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "0,00"
            TextSave        =   "0,00"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Picture         =   "frmblag.frx":123E
            Text            =   "POTREBE"
            TextSave        =   "POTREBE"
            Object.ToolTipText     =   "Izracun potreb"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   10440
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71172097
      CurrentDate     =   39507
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   12840
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   71172097
      CurrentDate     =   39507
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nova vrsta stolpec"
      Height          =   255
      Left            =   8520
      TabIndex        =   47
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label do_daxx 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
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
      Index           =   3
      Left            =   12120
      TabIndex        =   45
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label do_daxx 
      BackStyle       =   0  'Transparent
      Caption         =   "VAL"
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
      Index           =   2
      Left            =   12960
      TabIndex        =   44
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label do_daxx 
      BackStyle       =   0  'Transparent
      Caption         =   "DUR"
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
      Index           =   0
      Left            =   10560
      TabIndex        =   43
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rabat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   41
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label vskl 
      BackStyle       =   0  'Transparent
      Caption         =   "V Skladisce"
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
      Left            =   3720
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label skup 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "0,00"
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
      Left            =   11760
      TabIndex        =   36
      Top             =   9840
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SKUPAJ"
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
      Left            =   10080
      TabIndex        =   35
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Label do_daxx 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
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
      Index           =   1
      Left            =   7440
      TabIndex        =   27
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label dok 
      BackStyle       =   0  'Transparent
      Caption         =   "DO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label gskl 
      BackStyle       =   0  'Transparent
      Caption         =   "Skladisce"
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
      Left            =   3960
      TabIndex        =   25
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "la"
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
      Index           =   0
      Left            =   600
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "la"
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
      Index           =   1
      Left            =   600
      TabIndex        =   22
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   600
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   600
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   4
      Left            =   7440
      TabIndex        =   19
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   5
      Left            =   7440
      TabIndex        =   18
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   6
      Left            =   7440
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   7
      Left            =   7440
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   9720
      Width           =   15255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   3495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   15135
   End
End
Attribute VB_Name = "frmblag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

Dim ftchFlag As Boolean     ' 0 - default 1 - add,  2 - modi, +3 - del
Dim adRwFlag As Boolean
Dim edRwFalg As Boolean
Dim svRwFlag As Boolean

Dim edCol As Integer
Dim curCol As Integer
Dim curRow As Integer
Dim msgFlag As Boolean

Dim clk As Boolean

Dim st As String

Private Sub cmdAdd_Click()
    Dim imepol As String
imepol = ""
    Dim lstRow As Integer
  ' myConection.Execute ("delete  from trenutna where ltrim(pozicija)='" & LTrim(Str(fgtrial.Row)) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

   Dim xro As Integer
   If fgtrial.Rows = 2 Then
  ' myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   End If
   If RS.State = 1 Then RS.Close
   RS.Open "select * from trenutna where  pozicija='" & levi_pres(LTrim(Str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   xro = 1
   If Not RS.EOF Then
'

      For i = fgtrial.FixedCols To fgtrial.Cols - 1
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "POZICIJA" Then
       RS.Fields(Trim(fgtrial.TextMatrix(0, i))) = levi_pres(LTrim(Str(Val(fgtrial.TextMatrix(fgtrial.Row, i)))), 4)
       Else
       imepol = Trim(fgtrial.TextMatrix(0, i))
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "VISINA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "ZAPIRNIK" Then
       imepol = "placilo"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "SIRINA" Then
       imepol = "y"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "STEKLO" Then
       imepol = "stdok"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "LES" Then
       imepol = "kopija"
       End If
If Not imepol = "" Then
     If Not UCase(imepol) = "EM" Then
        RS.Fields(imepol) = fgtrial.TextMatrix(fgtrial.Row, i)
    End If
        End If
       End If
      Next i
         RS.Fields("datum") = Me.DTPicker1.Value
         RS.Fields("skl") = Me.sklad.BoundDatax
          RS.Fields("znes") = (RS.Fields("kol") * RS.Fields("cena")) * (1 - (RS.Fields("pop") / 100))
          RS.Fields("doza") = 1
          
    RS.Update
    End If
    
  '*pogleda zadnjo
   If RS.State = 1 Then RS.Close
   RS.Open "select * from trenutna where  pozicija='" & levi_pres(LTrim(Str(fgtrial.Rows - 1)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   xro = 1
   If Not RS.EOF Then
'
      For i = fgtrial.FixedCols To fgtrial.Cols - 1
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "POZICIJA" Then
       RS.Fields(Trim(fgtrial.TextMatrix(0, i))) = levi_pres(LTrim(Str(Val(fgtrial.TextMatrix(fgtrial.Rows - 1, i)))), 4)
       Else
       imepol = Trim(fgtrial.TextMatrix(0, i))
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "VISINA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "ZAPIRNIK" Then
       imepol = "placilo"
       End If
       
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "SIRINA" Then
       imepol = "y"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "STEKLO" Then
       imepol = "stdok"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "LES" Then
       imepol = "kopija"
       End If
      If Not imepol = "" Then
      If Not UCase(imepol) = "EM" Then
        RS.Fields(imepol) = fgtrial.TextMatrix(fgtrial.Rows - 1, i)
      End If
        End If
       End If
      Next i
         RS.Fields("datum") = Me.DTPicker1.Value
         RS.Fields("skl") = Me.sklad.BoundDatax
          RS.Fields("znes") = (RS.Fields("kol") * RS.Fields("cena")) * (1 - (RS.Fields("pop") / 100))
          RS.Fields("doza") = 1
          
    RS.Update
    End If
  
    
    txtNewData.Visible = False
        txtNewData.text = ""
   '     If fgtrial.TextMatrix(fgtrial.Rows - 1, 1) = "" Then
   '     Else
   '     fgtrial.Rows = fgtrial.Rows + 1
   '     End If
Dim fakkx As Long
fakkx = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")

 If Trim(fgtrial.TextMatrix(fgtrial.Rows - 1, 1)) <> "" Then
' MsgBox Trim(fgtrial.TextMatrix(fgtrial.Rows - 1, 1))
   If RS.State = 1 Then RS.Close
   RS.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   RS.AddNew
   RS.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")
   RS.Fields("tip_dok") = Left(Me.dok.Caption, 2)
   RS.Fields("id_dok") = Mid(Me.dok.Caption, 3)
   RS.Fields("datum") = Me.DTPicker1.Value
   RS.Fields("skl") = Me.sklad.BoundDatax
   RS.Fields("faktor") = fakkx
   RS.Fields("pozicija") = levi_pres(LTrim(Str(Val(Getnazi("select count(pozicija) as cc from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")) + 1)), 4)
   RS.Update
  End If
DoColumnSort
refre
dell = 1
   ' fgtrial.DataSource = RS
'    MsgBox ""
    lstRow = fgtrial.Rows - 1
    fgtrial.Row = lstRow
    fgtrial.Col = coollsi
     fgtrial.TextMatrix(lstRow, 0) = lstRow
     fgtrial.TextMatrix(lstRow, 1) = ""
     fgtrial.SetFocus
     
'Me.Text1.SetFocus
'     fgtrial.SetFocus
'DoColumnSort
'refre
End Sub


Private Sub cmdSave_Click()
    If svRwFlag = True Then
        Dim id As String
        Dim fld As String
        Dim dt As String
        Dim ftFg As Integer
                
        'Open "E:\VishwaPrg\Rohini\all_vb_prog\TrialGrid\trialSqul.sql" For Output As FreeFile
        For i = 1 To fgtrial.Rows - 1
            fgtrial.Row = i
            Dim rw As Integer
            Cols = fgtrial.Cols - 1
            
            fgtrial.Col = 0
            id = fgtrial.text
            dt = ""
            
            For k = 1 To Cols - 1
                fgtrial.Col = k
                dt = dt & "," & fgtrial.text
            Next k
            
            fgtrial.Col = k
            ftFg = Val(fgtrial.text)
        
            Select Case ftFg
                Case 0
                    MsgBox "Fetched" & " - " & Mid(dt, 2, Len(dt))
                Case 1
                    MsgBox "Added" & " - " & Mid(dt, 2, Len(dt))
                Case 2
                    MsgBox "Modified" & " - " & Mid(dt, 2, Len(dt))
                Case Else
                    MsgBox "Dele" & " - " & dt
            End Select
        Next i
        'Close
        svRwFlag = False
    Else
        MsgBox "Data Allready saved"
    End If
End Sub

Private Sub cmdUnDel_Click()
    For i = 1 To fgtrial.Cols - 1
        fgtrial.Col = i
        fgtrial.CellForeColor = vbBlack
    Next i
    
    If Val(fgtrial.text) > 2 Then
        fgtrial.text = Val(fgtrial.text) - 3
    Else
        MsgBox "The Record Is Not Deleted"
    End If
    
    fgtrial.Col = curCol
End Sub

Private Sub cmdDel_Click()
Dim rro As Integer
rro = fgtrial.Row
myConection.Execute ("delete  from trenutna where (pozicija)='" & levi_pres(LTrim(Str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
fgtrial.clear
 If Getnazi("select tekst from dokm where atribut='" & levi_pres(LTrim(Str(fgtrial.Row)), 4) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
 myConection.Execute ("delete  from dokm where atribut='" & levi_pres(LTrim(Str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 End If
    'Dim s As Integer
    'For s = 1 To fgtrial.Rows - 1
    'fgtrial.TextMatrix(s, 0) = LTrim(Str(s))
    'Next
    If RS.State = 1 Then RS.Close
    RS.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
    RS.MoveFirst
    
    Dim s As Integer
    s = 1
    Do While Not RS.EOF
    RS.Fields("pozicija") = levi_pres(LTrim(Str(s)), 4)
    s = s + 1
    RS.MoveNext
    Loop
   
  '  myConection.Execute ("update  trenutna set pozicija='" & LTrim(Str(recno())) & "' where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
    refre
    DoColumnSort
    End If
    If Getnazi("select tip_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") = "" Then
    Dim fakk As Long
    fakk = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
    myConection.Execute ("insert into trenutna (tip_dok,id_dok,pozicija,faktor,uporabnik) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','   1'," & fakk & ",'" & Getnazi("select up from users where username1='" & UPORABNIK & "'") & "')")
    refre
    DoColumnSort
     End If
  '  myConection.Execute ("update  trenutna set pozicija='" & LTrim(Str(recno())) & "' where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   If rro > fgtrial.Rows - 1 Then
   
   fgtrial.Row = fgtrial.Rows - 1
   Else
   fgtrial.Row = rro
   End If
    End Sub

Private Sub cmdFetch_Click()
    If RS.State = 0 And ftchFlag = False Then
        ftchFlag = True
      '  rs.Source = "Select * from utna"
        RS.Open "select pozicija,sifra,naziv,cena,kol,znes,x,y,chk_fix  from trenutna", myConection, adOpenDynamic, adLockOptimistic
        'Set fgTrial.DataSource = rs.Source
        
        If RS.EOF = False Then
            RS.MoveFirst
            For i = 1 To RS.RecordCount - 1
                If i <> fgtrial.Rows - 1 Then
                    fgtrial.Rows = fgtrial.Rows + 1
                End If
                fgtrial.Row = i
                For J = 0 To fgtrial.Cols - 2
                    fgtrial.Col = J
                    fgtrial.text = RS.Fields(J)
                Next J
                ' to set last col as fetch flag - 0
                fgtrial.Col = J
                fgtrial.text = 0
                RS.MoveNext
            Next i
        End If
    End If
End Sub

Private Sub DTPicker2_Change()
Me.DTPicker3.Value = Me.DTPicker2.Value + Me.Text5.text
End Sub

Private Sub DTPicker3_Change()
Me.Text5.text = Me.DTPicker3.Value - Me.DTPicker2.Value
End Sub

Private Sub fgTrial_Click()
 
    clk = False
    curRow = fgtrial.Row
    curCol = fgtrial.Col
    msgFlag = False
  If fgtrial.Col = 0 Then
    xopis = levi_pres(LTrim(Str(fgtrial.Row)), 4)
    xid_dok = Trim(dok.Caption)
    Dialog.Show vbModal
    fgtrial.Col = 1
  End If
  If fgtrial.Col = collchk Then
  If Trim(fgtrial.TextMatrix(curRow, curCol)) = "c" Then
  fgtrial.text = "b"
  Else
  fgtrial.text = "c"
  End If
  cmdAdd_Click
  End If
End Sub

Private Sub fgTrial_DblClick()
    clk = True
    
    
    If ftchFlag = True And adRwFlag = False Then
        edRwFalg = True
    End If
    
    curCol = 1
    curRow = fgtrial.Rows - 1
    fgTrial_KeyPress (0)
End Sub

Private Sub fgTrial_KeyPress(KeyAscii As Integer)
    Dim tmpCol As Integer
    'clk = True
If KeyAscii = 13 And fgtrial.Col <> coollsi Then
    If fgtrial.Col < fgtrial.Cols - 1 Then
      fgtrial.Col = fgtrial.Col + 1
      Else
       txtNewData.Visible = False
       
       
  '    cmdAdd_Click
     ' MsgBox fgtrial.Col
      'fgtrial.Col = coollsi
      End If
      
Else
    tmpCol = fgtrial.Col
    'fgtrial.Col = fgtrial.Cols - 1
        curRow = fgtrial.Row
       ' If adRwFlag = True Then
       '     curCol = 1
       ' Else
            curCol = fgtrial.Col
        'End If
       ' MsgBox curRow
        'MsgBox Chr(KeyAscii')
        
     '   fgtrial.Col = curCol
        
        txtNewData.text = Chr(KeyAscii)
        txtNewData.Move fgtrial.CellLeft + fgtrial.Left - 10, fgtrial.CellTop + _
                        fgtrial.Top - 10, fgtrial.CellWidth + 10, fgtrial.CellHeight - 40
        
        'to set col no to previous value
        txtNewData.Visible = True
'        Me.Text1.SetFocus
        txtNewData.SetFocus
      '  End If
    'Else
    '    MsgBox "Double Click To Make Edit Mode Active"
    'End If
    End If
End Sub
Private Sub fgTrial_GotFocus()

If fgtrial.Rows = 2 And fgtrial.TextMatrix(1, 0) = "" Then

fgtrial.Row = 1
fgtrial.Col = 0

fgtrial.text = 1
fgtrial.Col = 1

End If


End Sub
Private Sub fgTrial_LeaveCell()
 If fgtrial.Col = coollem Then
    If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='AVTOOP'") = "D" Then
     xopis = levi_pres(LTrim(Str(fgtrial.Row)), 4)
    xid_dok = Trim(dok.Caption)
    Dialog.Show vbModal
    fgtrial.SetFocus
    fgtrial.Col = fgtrial.Col + 1
    End If
    End If

     
    If txtNewData.Visible = False Then
        Exit Sub
    Else
        'If IsNumeric(txtNewData.text) Then
        '    fgtrial.text = Str(txtNewData.text)
        'Else
            fgtrial.text = txtNewData.text
        'End If
        txtNewData.Visible = False
        txtNewData.text = ""
    End If
    If fgtrial.Col = coollsi Or fgtrial.Col = coollstek Or fgtrial.Col = coollles Then
    Else
    If fgtrial.Col < fgtrial.Cols - 1 Then
      
      fgtrial.Col = fgtrial.Col + 1
      Else
     
      cmdAdd_Click
      fgtrial.Col = coollsi
      End If
     End If
    

End Sub



Private Sub fgtrial_RowColChange()
Dim axsi As String
  Dim das, dodx
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
  
axsi = fgtrial.TextMatrix(fgtrial.Row, 1)
Me.StatusBar1.Panels(1).text = Getnazi("select madanazi from mada where madasifr='" & axsi & "'") & " " & Getnazi("select madaenme from mada where madasifr='" & axsi & "'")
     Me.StatusBar1.Panels(3).text = Getnazi("select sum(kol) as ss from zaloga where sifra='" & axsi & "'")
     Me.StatusBar1.Panels(5).text = Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & axsi & "' and datum<=#" & dodx & "#")
End Sub

Private Sub Form_Load()
collchk = 0
coollsi = 0
coollna = 0
coollce = 0
coollko = 0
coollles = 0
coollstek = 0
Me.DTPicker1 = Date
If normati = "" Then
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from glavna where tip_dok='" & Trim(tip_dok) & "'")) + 1, 6)
Else
Me.dok.Caption = normati
normati = ""
End If
Me.sklad.BoundDatax = Getnazi("select skl from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Me.Text4.text = Getnazi("select dol_ce from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
If ma_ured <> 0 Then
Me.dok.Caption = Trim(tip_dok) & Trim(frmControlMain.MSHFlexGrid1.text)
'napolni
Else

End If

Dim Y As Integer
Y = 0

For Y = 0 To 7
Me.do_da(Y).Caption = Getnazi("select dod" & Trim(Str(Y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
If Me.do_da(Y).Caption = "" Then
Me.UserControl11(Y).Visible = False
Else
Me.UserControl11(Y).sSQL = Getnazi("select sqdo" & Trim(Str(Y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Me.UserControl11(Y).polje = Getnazi("select dpo" & Trim(Str(Y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
End If
Next
'napolni
Dim upor As String
upor = Getnazi("select up from users where username1='" & UPORABNIK & "'")

If Getnazi("select tip_dok from trenutna where tip_dok='" & tip_dok & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and uporabnik='" & upor & "'") <> "" Then
 boolConfirm = MsgBox("Ta datoteka že obstaja prepišem? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
       ' myConection.Execute ("insert into trenutna select * from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
       myConection.Execute ("delete from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
       Else
       myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
        End If
Else
myConection.Execute ("insert into trenutna select * from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija")
End If
napolni

Call GetNewConnection2
If Getnazi("select tip_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") = "" Then
Dim fakk As Long
fakk = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
'Set Rs1 = New Recordset
myConection.Execute ("insert into trenutna (tip_dok,id_dok,pozicija,faktor,uporabnik) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','   1'," & fakk & ",'" & Getnazi("select up from users where username1='" & UPORABNIK & "'") & "')")
End If
If kosovni = 1 Then
napolni
End If

refre
DoColumnSort

ReSizeForm Me
izja = 1
Set Rs1 = Nothing
Set DCON = Nothing
 Call WheelHook(Me.hWnd)
 If Left(Me.dok.Caption, 3) = "NTX" Or Left(Me.dok.Caption, 3) = "NTY" Then
 Me.Text3.Visible = True
 'Me.Text4.Visible = True
Me.Text3.text = Getnazi("select stdok from nabasif where tip_dok='NT' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'Me.Text4.text = Getnazi("select y from nabasif where tip_dok='NT' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 End If

'datumi
If Getnazi("select tekst from dokm where atribut='DUR' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.DTPicker2.Value = Getnazi("select tekst from dokm where atribut='DUR' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Else
Me.DTPicker2.Value = Date
End If
If Getnazi("select tekst from dokm where atribut='DNI' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.Text5.text = Getnazi("select tekst from dokm where atribut='DNI' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Else
Me.Text5.text = 0
End If
If Getnazi("select tekst from dokm where atribut='VAL' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.DTPicker3.Value = Getnazi("select tekst from dokm where atribut='VAL' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Else
Me.DTPicker3.Value = Date
End If
If imedn <> "" Then

Me.Text1.text = Getnazi("select opis from glavna where tip_dok='DN' and id_dok='" & Trim(imedn) & "'")
imedn = ""
End If
If Getnazi("select placilo from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
If Getnazi("select placilo from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") = 1 Then
Me.Check1.Value = 1
End If
End If
If Left(Me.dok.Caption, 2) = "DO" Or Left(Me.dok.Caption, 2) = "FA" Then
Me.Check1.Visible = True
If Left(Me.dok.Caption, 2) = "FA" Then
Me.Check1.Enabled = False
Else
Me.Check1.Enabled = True
End If
Else
Me.Check1.Visible = False
End If
If Left(Me.dok.Caption, 2) = "IZ" Then
Me.DTPicker2.Visible = False
Me.DTPicker3.Visible = False
Me.Text5.Visible = False
Me.do_daxx(0).Visible = False
Me.do_daxx(2).Visible = False
Me.do_daxx(3).Visible = False
End If

End Sub

Private Sub Form_Unload(cancel As Integer)
myConection.Execute ("delete  from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
    
    If RS.State = 1 Then
        RS.Close
    End If
'    cn.Close
If Left(Me.dok.Caption, 2) <> "NT" Then
osve = 1
End If
    Call WheelUnHook(Me.hWnd)
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
    If TypeOf ctl Is MSHFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then HorFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
     If TypeOf ctl Is DataGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then DataGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub
Sub refre()
Dim cooo As Integer
cooo = fgtrial.Col
SQL = "select " & Getnazi("select polja from dokumenti where tip_dok='" & tip_dok & "'") & " from trenutna  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija"
If Rs1.State = 1 Then Rs1.Close
Rs1.Open SQL, myConection, adOpenDynamic, adLockOptimistic
Set fgtrial.DataSource = Rs1
        
'Set Rs1 = DCON.Execute(SQL)
'ssqq = SQL
coollsi = 0
coollce = 0
coollchk = 0
coollko = 0
coollzn = 0
coollpop = 0
coollles = 0
coollstek = 0
coollzn = ""
fgtrial.Redraw = False
   ' For I = fgtrial.FixedCols To fgtrial.Cols - 1
    For i = fgtrial.Col To fgtrial.Cols - 1
        Dim asx As String
        
        asx = LCase(Trim(fgtrial.TextMatrix(0, i)))
        'MsgBox asx
        If asx = "sifra" Then
        coollsi = i
        'MsgBox coollsi
        'Exit For
        End If
        If UCase(asx) = "EM" Then
        coollem = i
        'MsgBox coollsi
        'Exit For
        End If
         If asx = "pop" Then
        coollpop = i
        'MsgBox coollsi
        'Exit For
        End If
         If UCase(asx) = "X" Then
        coollx = i
        End If
         If UCase(asx) = "Y" Then
        coolly = i
        End If
        If UCase(asx) = "STEKLO" Then
        coollstek = i
        End If
        If UCase(asx) = "LES" Then
        coollles = i
        End If
        
        If Left(asx, 3) = "chk" Then
        collchk = i
        'fgtrial.Row = fgtrial.Rows - 1
        'fgtrial.Col = i
        'fgtrial.ColWidth(i) = Check1(0).Width
        'Check1(0).Move fgtrial.CellLeft - fgtrial.Left, fgtrial.CellTop + fgtrial.Top _
        '                , fgtrial.CellWidth, fgtrial.CellHeight
                    
       ' Check1(0).Move
        'Exit For
        
        End If
        
        If asx = "naziv" Then
        coollna = i
        
        'Exit For
        End If
         If asx = "znes" Then
        coollzn = i
        
        'Exit For
        End If
        If asx = "cena" Then
        coollce = i
        'Exit For
        End If
        If asx = "kol" Then
        coollko = i
        'Exit For
        End If
         
      
      '   fgtrial.Col = iLoop
       Next i
       
  
     Dim lngX As Long
        fgtrial.Col = coollce
       
       ' cene
       
        'lngX = 1
        'While lngX + 1 < fgtrial.Rows
            
            ' fgtrial.TextMatrix(lngX, coollce) = FormatNumber(fgtrial.TextMatrix(lngX, coollce), 4)
             
         '   lngX = lngX + 1
        'Wend
  Dim cenn As Double
        With fgtrial
       ' MsgBox fgtrial.TextMatrix(lCount, coollce)
        .Redraw = False ' makes it about 10x faster !
        For lCount = .FixedRows To .Rows - 1
           'cena
            fgtrial.TextMatrix(lCount, coollem) = Getnazi("select madaenme from mada where madasifr='" & fgtrial.TextMatrix(lCount, coollsi) & "'")
           If coollce <> 0 Then
           cenn = Replace(fgtrial.TextMatrix(lCount, coollce), ".", ",")
             fgtrial.TextMatrix(lCount, coollce) = FormatNumber(cenn, 4)
             .ColAlignment(coollce) = flexAlignRightCenter
             End If
            'kol
             If coollko <> 0 Then
             cenn = Replace(Replace(fgtrial.TextMatrix(lCount, coollko), ",", ""), ".", ",")
             fgtrial.TextMatrix(lCount, coollko) = FormatNumber(cenn, 3)
             .ColAlignment(coollko) = flexAlignRightCenter
             End If
            'znes
             If coollzn = "" Then
             coollzn = 0
             End If
             If coollzn <> 0 Then
            cenn = Replace(fgtrial.TextMatrix(lCount, coollzn), ".", ",")
             fgtrial.TextMatrix(lCount, coollzn) = FormatNumber(cenn, 4)
             .ColAlignment(coollzn) = flexAlignRightCenter
             End If
              If coollpop <> 0 Then
              cenn = Replace(fgtrial.TextMatrix(lCount, coollpop), ".", ",")
             fgtrial.TextMatrix(lCount, coollpop) = FormatNumber(cenn, 2)
             .ColAlignment(coollpop) = flexAlignRightCenter
             End If
        Next
      
        
        
        Dim xro As Integer

'barvam pozicije
        If Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") < 0 Then
         If Getnazi("select pozicija from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
         For xro = fgtrial.FixedRows To fgtrial.Rows - 1
        fgtrial.Row = xro
        If Trim(fgtrial.TextMatrix(xro, coollsi)) <> "" Then
        If Getnazi("select madanazi from mada where madasifr='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
      If Getnazi("select sum(kol*faktor) as ss from nabasif where sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
        Dim das, dodx
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
  If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") <> "" Then
        If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") < 0 + fgtrial.TextMatrix(fgtrial.Row, coollko) Then
            
            fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            Me.fgtrial.Refresh
            Else
            fgtrial.CellBackColor = &HFFFFFF
            
       End If
       Else
       
            fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            Me.fgtrial.Refresh
        End If
       End If
       End If
       End If
       If Trim(fgtrial.TextMatrix(xro, coollko)) = 0 Then
       fgtrial.CellBackColor = 255
       Else
       fgtrial.CellBackColor = &HFFFFFF
         
       End If
       Next xro
       End If
       End If
        
        
         
        .Redraw = True ' dont forget to do this !
        End With

   If collchk <> 0 Then
   With fgtrial
   .Redraw = False
   If .Row > 0 Then
       .FillStyle = flexFillRepeat
       .Col = collchk
        .Row = 1
        .RowSel = .Rows - 1
        .CellFontName = "Marlett"
        .CellAlignment = 2
     End If
     End With
     End If
       ' For iLoop = fgtrial.FixedRows To fgtrial.Rows - 1
       '  fgtrial.Row = iLoop
       '  fgtrial.Col = collchk
       '  fgtrial.ColAlignment(coollchk) = flexAlignCenterCenter
        
       ' fgtrial.CellFontName = "Marlett"
        
       ' Next iLoop
        fgtrial.Redraw = True
          cenn = Replace(Getnazi("select sum(znes) as znes from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"), ".", ",")
             Me.skup.Caption = FormatNumber(cenn, 2)
'fgtrial.Col = cooo
End Sub
Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------

    With fgtrial
    
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
        .Redraw = True
        
    End With
   '  Dim iLoop As Integer
     fgtrial.Redraw = False
        For i = fgtrial.FixedCols To fgtrial.Cols - 1
        Dim asx As String
        
        asx = Trim(fgtrial.TextMatrix(0, i))
        
        If asx = "sifra" Then
        coollsi = i
        'Exit For
        End If
        
        If Left(asx, 3) = "chk" Then
        collchk = i
        'fgtrial.Row = fgtrial.Rows - 1
        'fgtrial.Col = i
        'fgtrial.ColWidth(i) = Check1(0).Width
        'Check1(0).Move fgtrial.CellLeft - fgtrial.Left, fgtrial.CellTop + fgtrial.Top _
        '                , fgtrial.CellWidth, fgtrial.CellHeight
                    
       ' Check1(0).Move
        'Exit For
        
        End If
        
        If asx = "naziv" Then
        coollna = i
        
        'Exit For
        End If
        If asx = "cena" Then
        coollce = i
        'Exit For
        End If
        If asx = "kol" Then
        coollko = i
        'Exit For
        End If
        If UCase(asx) = "EM" Then
        coollem = i
        'MsgBox coollsi
        'Exit For
        End If
        
        If UCase(asx) = "STEKLO" Then
        coollstek = i
        End If
        If UCase(asx) = "LES" Then
        coollles = i
        End If
      
      '   fgtrial.Col = iLoop
       Next i
       
       
fgtrial.ColWidth(coollna) = 6000
fgtrial.ColWidth(collchk) = 1000
fgtrial.ColAlignment(coollchk) = flexAlignCenterCenter
Dim xro As Integer

'barvam pozicije
        If Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") < 0 Then
         If Getnazi("select pozicija from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
         For xro = fgtrial.FixedRows To fgtrial.Rows - 1
        fgtrial.Row = xro
        If Trim(fgtrial.TextMatrix(xro, coollsi)) <> "" Then
        If Getnazi("select madanazi from mada where madasifr='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
      If Getnazi("select sum(kol*faktor) as ss from nabasif where sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
        Dim das, dodx
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
  If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") <> "" Then
        If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") < 0 + fgtrial.TextMatrix(fgtrial.Row, coollko) Then
            
            fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            Me.fgtrial.Refresh
            Else
            fgtrial.CellBackColor = &HFFFFFF
            
       End If
       Else
       
            fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            Me.fgtrial.Refresh
        End If
       End If
       End If
       End If
       Next xro
       End If
       End If
      fgtrial.Redraw = True

End Sub

Private Sub Isc_Click()
Dim iskan As String

iskan = UCase(InputBox("Vnesi isklani niz", "Vnesi iskalni niz"))
iskan = "*" & iskan & "*"
'fgtrial.Col = 1
'fgtrial.Row = 1
fgtrial.Redraw = False

For x = fgtrial.Row To fgtrial.Rows - 1
For i = fgtrial.Col To fgtrial.Cols - 1
fgtrial.Col = i
fgtrial.Row = x
If UCase(fgtrial.TextMatrix(x, i)) Like iskan Then
fgtrial.Redraw = True
Exit Sub
End If
Next i
fgtrial.Col = 0
Next x
fgtrial.Redraw = True

End Sub

Private Sub Label2_Click()
Dim sss As String
sss = "update trenutna set znes=kol*(cena*(1-(pop/100))) where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        
       
        myConection.Execute (sss)
        DoColumnSort
        refre
End Sub

Private Sub LaVolpeButton1_click()
Dim sss As String
If RS.State = 1 Then RS.Close
RS.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
If Not RS.EOF Then
RS.MoveFirst
End If
Do While Not RS.EOF
MsgBox Getnazi("select madampcd from mada where madasifr='" & RS.Fields("sifra") & "'")
If Getnazi("select madampcd from mada where madasifr='" & RS.Fields("sifra") & "'") <> "" Then
RS.Fields("cena") = FormatNumber(Getnazi("select madampcd from mada where madasifr='" & RS.Fields("sifra") & "'"), 4)
RS.Fields("znes") = FormatNumber(RS.Fields("cena") * RS.Fields("kol"), 4)
End If
RS.Update
RS.MoveNext
Loop
       
       ' myConection.Execute (sss)
        DoColumnSort
        refre
End Sub

Private Sub prekin_Click()
Unload Me
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Left(Me.dok.Caption, 2) = "DN" Then
Dim norma, stek, lesi As String
Dim koli As Long
Dim xrsn As New ADODB.Recordset
Dim xox, yoy, zapp, dkr As Integer
Dim kol, XX, yy As Long
dkr = 1

imedn = frmControlMain.MSHFlexGrid1.text
If xrsn.State = 1 Then xrsn.Close
xrsn.Open "select * from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic
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
'MsgBox Getnazi("select madaemba from mada where madasifr='" & norma & "'")
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
yy = xrsn.Fields("y")
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
rsta.Fields("poz") = levi_pres(LTrim(Str(ss)), 4)
If fixx = 0 Then
If Getnazi("select madaenme from mada where madasifr='" & sii & "'") = "KOM" Then
rsta.Fields("kol") = FormatNumber(kol * koli * (((XX / 100) * (yy / 100)) / xfxt), 0)
Else
rsta.Fields("kol") = FormatNumber(kol * koli * (((XX / 100) * (yy / 100)) / xfxt), 4)
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
rsta.Fields("poz") = levi_pres(LTrim(Str(ss)), 4)
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
rsta.Fields("poz") = levi_pres(LTrim(Str(ss)), 4)
rsta.Fields("kol") = kol * koli
rsta.Fields("zap") = zapp
rsta.Update
rst.MoveNext
Loop
'steklo
rsta.AddNew
rsta.Fields("sifr") = stek
rsta.Fields("naz") = Getnazi("select madanazi from mada where madasifr='" & stek & "'")
rsta.Fields("poz") = levi_pres(LTrim(Str(ss)), 4)
rsta.Fields("kol") = XX * yy * koli / 10000 * 0.76
rsta.Update
'les
If RS.State = 1 Then RS.Close
If lesi <> "" Then
RS.Open "select * from sestavi where sifra=" & lesi, myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
RS.MoveFirst
End If
Dim faktx As Double
faktx = Getnazi("select madaemba from mada where madasifr='" & lesi & "'")
Do While Not RS.EOF

rsta.AddNew
rsta.Fields("sifr") = RS.Fields("sifras")
rsta.Fields("naz") = Getnazi("select madanazi from mada where madasifr='" & RS.Fields("sifras") & "'")
rsta.Fields("poz") = levi_pres(LTrim(Str(ss)), 4)
rsta.Fields("kol") = FormatNumber(((XX * yy * koli / 10000) / faktx) * RS.Fields("kol"), 3)
rsta.Update
RS.MoveNext
Loop
End If
xrsn.MoveNext
Loop
If RS.State = 1 Then RS.Close
RS.Open "select sifr,min(naz) as naz,sum(kol) as kol,sum(zap) as zap from xnorm group by sifr", myConection, adOpenDynamic, adLockOptimistic
Dim Rsa As New ADODB.Recordset
If Rsa.State = 1 Then Rsa.Close
Rsa.Open "select * from normati", myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
RS.MoveFirst
End If

Do While Not RS.EOF
Rsa.AddNew
Rsa.Fields("sifr") = RS.Fields("sifr")
Rsa.Fields("naz") = RS.Fields("naz")
Rsa.Fields("kol") = RS.Fields("kol")
Rsa.Fields("zap") = RS.Fields("zap")
Rsa.Update
RS.MoveNext
Loop
If Rsa.State = 1 Then Rsa.Close
Rsa.Open "select * from normati where sifr='10217'", myConection, adOpenDynamic, adLockOptimistic
If Not Rsa.EOF Then
Rsa.Fields("kol") = Getnazi("select sum(zap) as x from normati")
Rsa.Update

End If
preg.Show
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdAdd_Click
End If
End Sub

Private Sub Text1_LostFocus()

'cmdAdd_Click
End Sub

Private Sub Text2_LostFocus()
 boolConfirm = MsgBox("Dam rabat na vse pozicije?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
        Dim sss As String
        sss = "update trenutna set pop=" & Replace(Me.Text2.text, ",", ".") & " where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        'MsgBox sss
        myConection.Execute (sss)
       sss = "update trenutna set znes=kol*(cena*(1-(pop/100))) where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        
       
        myConection.Execute (sss)
        DoColumnSort
        refre
        End If
End Sub

Private Sub Text3_LostFocus()
If Me.Text3.text = "" Then
MsgBox "Obvezen vnos!!"
Me.Text3.SetFocus

End If
End Sub

Private Sub Text5_Change()
If Me.Text5.text = "" Then
Me.Text5.text = "0"
End If
Me.DTPicker3.Value = Me.DTPicker2.Value + Me.Text5.text
End Sub

Private Sub UserControl11_LostFocus(Index As Integer)
If Me.UserControl11(0).BoundDatax <> "" Then
If Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'") <> "" Then
If Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'") = 0 Then
Me.Text5.text = Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PLDNI'")
Else
Me.Text5.text = Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'")
End If
Text5_Change
End If
End If
End Sub

Private Sub zapis_Click()
If Me.sklad.BoundDatax = "" Then
MsgBox "Vnos skladisca je Obvezen!!"
Exit Sub
End If
If Me.Check1.Value Then
myConection.Execute ("update trenutna set placilo=1 where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
Else
myConection.Execute ("update trenutna set placilo=0 where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
End If
myConection.Execute ("update trenutna set skl='" & Me.sklad.BoundDatax & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("update trenutna set faktor='" & Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("update trenutna set datum='" & Me.DTPicker1.Value & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
If Left(Me.dok.Caption, 3) = "NTX" Or Left(Me.dok.Caption, 3) = "NTY" Then
myConection.Execute ("update trenutna set stdok='" & Me.Text3.text & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'myConection.Execute ("update trenutna set y='" & Me.Text4.text & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
End If
myConection.Execute ("delete  from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete  from trenutna where (sifra)='' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete  from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
If RS.State = 1 Then RS.Close
RS.Open "select * from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija", myConection, adOpenDynamic, adLockOptimistic
Dim aa As Integer
aa = 1
If Not RS.EOF Then
RS.MoveFirst
Do While Not RS.EOF
RS.Fields("pozicija") = levi_pres(aa, 4)
aa = aa + 1
RS.MoveNext
Loop
End If
'myConection.Execute ("delete  from trenutna where ltrim(sifra)='' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
Dim upor As String
upor = Getnazi("select up from users where username1='" & UPORABNIK & "'")
myConection.Execute ("insert into nabasif select * from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 SQL = "insert into glavna (tip_dok,id_dok,opis,dod0,dod1,dod2,dod3,dod4,dod5,dod6,dod7,skl) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','" & Me.Text1.text & "','" & Me.UserControl11(0).BoundDatax & "','" & Me.UserControl11(1).BoundDatax & "','" & Me.UserControl11(2).BoundDatax & "','" & Me.UserControl11(3).BoundDatax & "','" & Me.UserControl11(4).BoundDatax & "','" & Me.UserControl11(5).BoundDatax & "','" & Me.UserControl11(6).BoundDatax & "','" & Me.UserControl11(7).BoundDatax & "','" & Me.sklad.BoundDatax & "')"
 ' MsgBox SQL
    myConection.Execute SQL
'datumi
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='DUR'")
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='DNI'")
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='VAL'")
If RS.State = 1 Then RS.Close
RS.Open "Select * from dokm where atribut='DUR'", myConection, adOpenDynamic, adLockOptimistic
RS.AddNew
RS.Fields("tip_dok") = Left(Me.dok.Caption, 2)
RS.Fields("id_dok") = Mid(Me.dok.Caption, 3)
RS.Fields("atribut") = "DUR"
RS.Fields("tekst") = Me.DTPicker2.Value
RS.Update
RS.AddNew
RS.Fields("tip_dok") = Left(Me.dok.Caption, 2)
RS.Fields("id_dok") = Mid(Me.dok.Caption, 3)
RS.Fields("atribut") = "DNI"
RS.Fields("tekst") = Me.Text5.text
RS.Update
RS.AddNew
RS.Fields("tip_dok") = Left(Me.dok.Caption, 2)
RS.Fields("id_dok") = Mid(Me.dok.Caption, 3)
RS.Fields("atribut") = "VAL"
RS.Fields("tekst") = Me.DTPicker3.Value
RS.Update
If Left(Getnazi("select tvorba from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'"), 1) = "=" Then
myConection.Execute ("delete from nabasif where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete from glavna where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete from dokm where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into nabasif select 'FA' as tip_dok,DATUM, STDOK, SIFRAPART, SIFRA, EMBALAZA, KOL, CENA, ZNES, pop,  id_dok, poknj, faktor, naziv, SIFRAPLAC, mpc, x, y, uporabnik, pozicija, chk_fix, skl, kopija, dat_k, PLACILO, DOZA, org from nabasif where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into glavna select 'FA' as tip_dok,id_dok, faktor, dod0, dod1, dod2, dod3, dod4, dod5, dod6, dod7, opis, skl from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into dokm select 'FA' as tip_dok,atribut, id_dok, tekst, poz from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

End If
myConection.Execute ("update dokumenti set dol_ce=" & Val(Me.Text4.text) & " where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
    Unload Me
End Sub

Private Sub opiss_Click()
xopis = "opis"
  xid_dok = Trim(dok.Caption)
  Dialog.Show

End Sub

Private Sub Timer1_Timer()
'Exit Sub
If fgtrial.Col = Val(Me.Text4.text) Then
       cmdAdd_Click
      fgtrial.Col = coollsi
      Else
      'fgtrial.Col = fgtrial.Col + 1
      End If
If fgtrial.Col = coollem Then
fgTrial_LeaveCell
 fgtrial.Col = fgtrial.Col + 1
 Exit Sub
 End If
If dell = 1 Then
dell = 0
If fgtrial.Col = 0 Then
refre
fgtrial.Col = 1
End If
If izja <> 0 Then
If Getnazi("select tekst from dokm where atribut='opis' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
opiss.BackColor = 255
Else
opiss.BackColor = &HE0E0E0

End If
fgtrial.Redraw = False
For i = 1 To fgtrial.Rows - 1
fgtrial.Col = 0
fgtrial.Row = i
If Getnazi("select tekst from dokm where atribut='" & levi_pres(LTrim(Str(i)), 4) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then

fgtrial.CellBackColor = 255
Else
fgtrial.CellBackColor = &H80000005
End If
Next
fgtrial.Redraw = True
fgtrial.Row = izja
izja = 0
fgtrial.Col = 1
End If
lstRow = fgtrial.Rows - 1
    fgtrial.Row = lstRow
    fgtrial.Col = coollsi
     fgtrial.TextMatrix(lstRow, 0) = lstRow
     fgtrial.TextMatrix(lstRow, 1) = ""
     fgtrial.SetFocus
End If

End Sub

Private Sub txtNewData_GotFocus()
tresi = fgtrial.text
If clk = True Then
   txtNewData.text = fgtrial.text
clk = False
End If
If fgtrial.Col = 0 Then
fgTrial_LeaveCell
Else
      fgtrial.Row = curRow
    fgtrial.Col = curCol
   adRwFlag = False
   txtNewData.SelStart = Len(txtNewData)
   txtNewData.SelLength = Len(txtNewData) + 1



End If
 
End Sub

Private Sub txtNewData_LostFocus()
'If Getnazi("select postava from mada where madasifr='" & Trim(txtNewData.text) & "'") <> "" Then
'       zai = fgtrial.Row
'       Postava.Show
'End If
 If curCol < edCol Or curCol = Val(Me.Text4.text) Then
        fgtrial.Row = curRow
        fgtrial.Col = curCol + 1
        fgTrial_KeyPress (0)
    Else
    
      fgTrial_LeaveCell
    End If
       End Sub
Private Sub txtNewData_KeyPress(KeyAscii As Integer)
'MsgBox ("3")
If KeyAscii = 27 Then
fgtrial.text = tresi
txtNewData.Visible = False
End If
If KeyAscii = 13 Then
 fgtrial.text = txtNewData.text
    If fgtrial.Col = coollsi Then
       If RS.State = 1 Then RS.Close
       Dim ax As String
       ax = ""
       ax = (Getnazi("select madasifr from mada where madasifr='" & txtNewData.text & "'"))
       If ax = "" Then
       Dim novas, vi, dol As String
       vi = ""
       dol = ""
       novas = "/" & Trim(txtNewData.text) & "/"
       ax = (Getnazi("select madasifr from mada where dobavit_id like '%" & novas & "%'"))
       End If
      
       If ax = "" Then
       idar = ""
       iskalni = fgtrial.text
       pritisk = txtNewData.text
      ' DoSQL = ""
      If tip_dok = "DN" Then
       ax = DoSQL("mada where tip_art='IZD'", "madasifr", "madanazi", "madaenme")
       Else
       ax = DoSQL("mada where tip_art<>'IZD'", "madasifr", "madanazi", "madaenme")
    End If
       'MsgBox ax
       End If
       txtNewData.text = Trim((ax))
       fgtrial.text = Trim(ax)
       sifrt = (ax)
    If sifrt = "" Then
    Else
       'StatusBar1.Panels.Remove 1
        
        'StatusBar1.Panels.Add 1, , "Artikel " & ax
        RS.Open "select MADANAZI,MADAnabc,madampcd,madapd,postava,madaenme from MADA where MADASIFR='" & ax & "'", myConection, adOpenStatic, adLockOptimistic
          If Not RS.EOF Then
              fgtrial.TextMatrix(fgtrial.Row, coollna) = Trim(RS!MADANAZI) & " "
              fgtrial.TextMatrix(fgtrial.Row, coollem) = Trim(RS!MADAenme) & " "
              If Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") <= 0 Then
              fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(RS!MADAMPCD, 4)
              Else
               fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(RS!madanabc, 4)
               End If
          End If
          Call txtNewData_LostFocus
          
          fgtrial.Col = 3
    Dim zalog, prost As Double
    Dim das, dodx
    Dim nazb As String
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
     nazb = Trim(RS.Fields("madanazi"))
     Me.StatusBar1.Panels(1).text = Trim(RS!MADANAZI) & " " & RS!MADAenme
     
     Me.StatusBar1.Panels(3).text = Getnazi("select sum(kol) as ss from zaloga where sifra='" & ax & "'")
     Me.StatusBar1.Panels(5).text = Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and  sifra='" & ax & "'  and datum<=#" & dodx & "#")
        'End If
    End If
    
      ElseIf fgtrial.Col = coollko Then
       If IsNumber(txtNewData.text) Then
       Else
       txtNewData.text = 0
       End If
      'fgtrial.Rows = fgtrial.Rows + 1
      'If tip_dok <> "NA" Then
      'cmdAdd_Click
      'fgtrial.Col = coollsi
      'End If
      
      Else
     
     If fgtrial.Col = coollles Then
     
idar = ""
       iskalni = fgtrial.text
       pritisk = txtNewData.text
     If Getnazi("select madasifr from mada where madasifr='" & pritisk & "'") <> "" Then
     ax = pritisk
     'fgtrial.Col = fgtrial.Col + 1
     Else
      ' DoSQL = ""
     ' MsgBox "mada where madagrup=14 and madanazi like '%" & fgtrial.TextMatrix(fgtrial.Row, coollna) & "'"
       ax = DoSQL("mada where madagrup=14 and madanazi like '%" & Trim(fgtrial.TextMatrix(fgtrial.Row, coollna)) & "'", "madasifr", "madanazi", "madanaz1")
      
       fgtrial.TextMatrix(fgtrial.Row, coollles) = Trim((ax))
      'MsgBox Me.txtNewData.text
     
     'fgtrial.text = Trim(ax)
     fgtrial.Col = fgtrial.Col + 1
     End If
       
       
End If

 If fgtrial.Col = coollstek Then
idar = ""
       iskalni = fgtrial.text
       pritisk = txtNewData.text
      ' DoSQL = ""
       If Getnazi("select madasifr from mada where madasifr='" & pritisk & "'") <> "" Then
     ax = pritisk
     'fgtrial.Col = fgtrial.Col + 1
     Else
       ax = DoSQL("mada where madagrup=2", "madasifr", "madanazi", "madanaz1")
       'MsgBox ax
      
       fgtrial.TextMatrix(fgtrial.Row, coollstek) = Trim((ax))
      ' txtNewData.text = Trim((ax))
      ' fgtrial.text = Trim(ax)
      fgtrial.Col = fgtrial.Col + 1
      End If
       
       
End If
      
      
      If fgtrial.Col = coollzn Then
      fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(FormatNumber(txtNewData.text, 4) / FormatNumber(fgtrial.TextMatrix(fgtrial.Row, coollko), 4), 7)
      
      If tip_dok = "NA" Then
  '    cmdAdd_Click
  '    fgtrial.Col = coollsi
      End If
      Else
      If fgtrial.Col = coollpop Then
       If IsNumber(txtNewData.text) Then
      txtNewData.text = FormatNumber(txtNewData.text, 2)
      Else
      txtNewData.text = 0
      End If
      If tip_dok = "NA" Then
 '     cmdAdd_Click
 '     fgtrial.Col = coollsi
      End If
      End If
      End If
      If fgtrial.Col = coollce Then
      Dim ben As Double
    '  ben = txtNewData.text
         If txtNewData.text = "" Then
         txtNewData.text = 0
         End If
         If IsNumber(txtNewData.text) Then
          txtNewData.text = FormatNumber(txtNewData.text, 4)
      If tip_dok <> "NA" Then
'           cmdAdd_Click
'          fgtrial.Col = coollsi
    End If
          Else
          txtNewData.text = 0
         End If
 '
        
      End If
      End If

If fgtrial.Col = Val(Me.Text4.text) Then
 cmdAdd_Click
 fgtrial.Col = coollsi
 End If
fgtrial.SetFocus
End If
End Sub

Private Sub napolni()
Dim i, stot, fa
 If RS.State = 1 Then RS.Close
 If kosovni = 1 Then
 Else
  

 RS.Open "select * from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
Dim C As Integer
If Not RS.EOF Then
If Not RS.Fields("skl") = "" Then
Me.sklad.BoundDatax = RS.Fields("skl")
End If
If Not RS.Fields("opis") = "" Then
Me.Text1.text = RS.Fields("opis")
End If

End If
If Getnazi("select datum from nabasif  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
Me.DTPicker1.Value = Getnazi("select datum from nabasif  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
Else
Me.DTPicker1.Value = Date
End If
For C = 0 To 7
If Not RS.EOF Then
If Not RS.Fields(C + 3) = "" Then
Me.UserControl11(C).BoundDatax = RS.Fields(C + 3)
End If
End If
Next
End If
'MsgBox (aaa)
   If RS.State = 1 Then RS.Close
   If kosovni = 1 Then
 Dim Rsa As New ADODB.Recordset
  
  Rsa.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
   RS.Open "select * from normati ", myConection, adOpenStatic, adLockOptimistic
   If Not Rsa.EOF Then
   myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   End If
   If Not RS.EOF Then
   RS.MoveFirst
   Dim aa As Integer
   aa = 1
   Do While Not RS.EOF
    Rsa.AddNew
    Rsa.Fields("tip_dok") = Left(Me.dok.Caption, 2)
    Rsa.Fields("id_dok") = Mid(Me.dok.Caption, 3)
    Rsa.Fields("pozicija") = levi_pres(LTrim(Str(aa)), 4)
    Rsa.Fields("sifra") = RS.Fields("sifr")
    Rsa.Fields("kol") = RS.Fields("kol")
    Rsa.Fields("naziv") = RS.Fields("naz")
    Rsa.Update
    
    aa = aa + 1
    
    RS.MoveNext
    
   Loop
'   Call cmdAdd_Click
kosovni = 0
   refre
   
   End If
   Else
 RS.Open "select * from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
 End If
  If ma_ko = 1 Then
 'MsgBox tip_dok
 myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   
 'myConection.Execute ("insert into trenutna select * from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija")
myConection.Execute ("update trenutna set id_dok='" & novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='" & Trim(dtip_dok) & "'")) + 1, 6) & "'  where tip_dok='XS' and id_dok='AAA'")
myConection.Execute ("update trenutna set tip_dok='" & Trim(dtip_dok) & "'  where tip_dok='XS'")
Me.dok.Caption = Trim(dtip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='" & Trim(dtip_dok) & "'")) + 1, 6)
myConection.Execute ("update dokm set atribut='opis'  where tip_dok='XS' and id_dok='AAA'")
myConection.Execute ("update dokm set id_dok='" & novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='" & Trim(dtip_dok) & "'")) + 1, 6) & "'  where tip_dok='XS' and id_dok='AAA'")
myConection.Execute ("update DOKM set tip_dok='" & Trim(dtip_dok) & "'  where tip_dok='XS'")
refre

ma_ko = 0
End If

Dim po As Integer
Dim kol As Integer
Dim znes As Double
po = 1
If kosovni = 1 Then
'Me.DTPicker1.Value = Date
Else
If Left(Me.dok.Caption, 2) = "NT" Then
Else
'Me.DTPicker1.Value = RS.Fields("datum")
End If
End If

kosovni = 0
End Sub

