VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmblag 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DOKUMENT"
   ClientHeight    =   10680
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton shran 
      Height          =   270
      Left            =   0
      MaskColor       =   &H8000000F&
      Picture         =   "frmblag.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   9600
      Width           =   270
   End
   Begin LVbuttons.LaVolpeButton Breme 
      Height          =   495
      Left            =   3960
      TabIndex        =   58
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "BREMEPIS"
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
      MICON           =   "frmblag.frx":00FA
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
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   14520
      TabIndex        =   57
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin LVbuttons.LaVolpeButton sturdod 
      Height          =   495
      Left            =   13320
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "DODAJ"
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
      MICON           =   "frmblag.frx":0116
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
   Begin VB.TextBox stur 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1060
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   55
      Text            =   "10"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4560
      TabIndex        =   51
      Top             =   480
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10080
      TabIndex        =   47
      Text            =   "Text4"
      Top             =   240
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IZVOZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   46
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12120
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   255
      Left            =   11400
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
      MICON           =   "frmblag.frx":0132
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10320
      TabIndex        =   40
      Text            =   "0,00"
      Top             =   3840
      Width           =   975
   End
   Begin LVbuttons.LaVolpeButton cmdadd 
      Height          =   735
      Left            =   360
      TabIndex        =   27
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
      MICON           =   "frmblag.frx":014E
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
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      ssql            =   "select * from skla"
      polje           =   "skladisce"
      textlocked      =   0
      locked          =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51707905
      CurrentDate     =   39507
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9000
      Top             =   600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   14
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
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   13560
      Top             =   1080
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
         Name            =   "Arial"
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
      TabIndex        =   19
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
      TabIndex        =   20
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
      TabIndex        =   21
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
      Left            =   7920
      TabIndex        =   22
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
      Left            =   7920
      TabIndex        =   23
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
      Left            =   7920
      TabIndex        =   24
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
      Left            =   7920
      TabIndex        =   25
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
      TabIndex        =   28
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
      MICON           =   "frmblag.frx":016A
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
      TabIndex        =   29
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
      MICON           =   "frmblag.frx":0186
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
      TabIndex        =   30
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
      MICON           =   "frmblag.frx":01A2
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
      TabIndex        =   31
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
      MICON           =   "frmblag.frx":01BE
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
      TabIndex        =   32
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
      MICON           =   "frmblag.frx":01DA
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
      DragIcon        =   "frmblag.frx":01F6
      Height          =   5040
      Left            =   120
      TabIndex        =   33
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
         Name            =   "Arial"
         Size            =   9.75
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
      MICON           =   "frmblag.frx":0500
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
      Top             =   10305
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   10584
            MinWidth        =   10584
            Picture         =   "frmblag.frx":051C
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
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51707905
      CurrentDate     =   39507
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   12840
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51707905
      CurrentDate     =   39507
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1680
      TabIndex        =   54
      Top             =   1920
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.Label av_zne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13680
      TabIndex        =   53
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Avans:"
      Height          =   255
      Left            =   12960
      TabIndex        =   52
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13680
      TabIndex        =   50
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Znesek:"
      Height          =   255
      Left            =   12960
      TabIndex        =   49
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Skok v nov stolpec:"
      Height          =   255
      Left            =   8520
      TabIndex        =   48
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label do_daxx 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      BeginProperty Font 
         Name            =   "Arial"
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
         Name            =   "Arial"
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
         Name            =   "Arial"
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   41
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label vskl 
      BackStyle       =   0  'Transparent
      Caption         =   "V Skladisce"
      BeginProperty Font 
         Name            =   "Arial"
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
         Name            =   "Arial"
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
         Name            =   "Arial"
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
         Name            =   "Arial"
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label gskl 
      BackStyle       =   0  'Transparent
      Caption         =   "Skladisce"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "la"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "la"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label do_da 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   2
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
   Begin VB.Menu mnu_uvoz 
      Caption         =   "mnu_uvoz"
      Visible         =   0   'False
      Begin VB.Menu mnuvozotv 
         Caption         =   "Uvoz Otvoritev(FIFO)"
      End
      Begin VB.Menu mnuvozgost 
         Caption         =   "Uvoz zalog(GOSTINSTVO)"
      End
      Begin VB.Menu mnuuvozin 
         Caption         =   "Uvoz Inventure"
      End
      Begin VB.Menu mnuuvoxls 
         Caption         =   "Uvoz XLS"
      End
   End
End
Attribute VB_Name = "frmblag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim st As String
Private dokl As String

Private Sub av_zne_Change()
Me.Label5.Caption = Me.skup.Caption - Me.av_zne.Caption
End Sub





Private Sub Breme_Click()
If Breme.BackColor = &HFF& Then
Breme.BackColor = &HFFFFFF
bremepis = 0
Else
Breme.BackColor = &HFF&
bremepis = 1
End If
End Sub

Public Function doddaa()
cmdDel_Click
cmdAdd_Click
End Function
Public Sub beref()
Zapis_Click
End Sub
Private Sub cmdAdd_Click()
'Me.WindowState = 1
'frmMAIN.WindowState = 1
On Error GoTo bbr:
fgtrial.Redraw = False
    Dim imepol As String
imepol = ""
    Dim lstRow As Integer
  ' myConection.Execute ("delete  from trenutna where ltrim(pozicija)='" & LTrim(Str(fgtrial.Row)) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

   Dim xro As Integer
   If fgtrial.Rows = 2 Then
  ' myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   End If
   If rs.State = 1 Then rs.Close
   rs.Open "select * from trenutna where  pozicija='" & levi_pres(LTrim(str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   xro = 1
   If Not rs.EOF Then
'

      For i = fgtrial.FixedCols To fgtrial.Cols - 1
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "POZICIJA" Then
       xpozi = levi_pres(LTrim(str(Val(fgtrial.TextMatrix(fgtrial.Row, i)))), 4)
       rs.Fields(Trim(fgtrial.TextMatrix(0, i))) = levi_pres(LTrim(str(Val(fgtrial.TextMatrix(fgtrial.Row, i)))), 4)
       Else
       imepol = Trim(fgtrial.TextMatrix(0, i))
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "VISINA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "DAT_PRE" Then
       imepol = "dat_k"
       End If
       
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "UR" Then
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
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "PROSTA" Then
       imepol = "sifraplac"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "MARZA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "MPC" Then
       imepol = "y"
       End If
       
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "LES" Then
       imepol = "kopija"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "ZALOGA" Then
       imepol = "x"
       End If
If Not imepol = "" Then
     If Not UCase(imepol) = "EM" Then
        If Not UCase(imepol) = "DAT_K" Then
      
        rs.Fields(imepol) = fgtrial.TextMatrix(fgtrial.Row, i)
        Else
        rs.Fields(imepol) = ctod(fgtrial.TextMatrix(fgtrial.Row, i))
        
    End If
    End If
        End If
       End If
      Next i
         rs.Fields("datum") = Me.DTPicker1.Value
         rs.Fields("skl") = Me.sklad.BoundDatax
          rs.Fields("znes") = (rs.Fields("kol") * rs.Fields("cena")) * (1 - (rs.Fields("pop") / 100))
          rs.Fields("y") = rs.Fields("cena") * (1 + (rs.Fields("x") / 100))
          rs.Fields("doza") = 1
'If Left(Me.dok.Caption, 2) = "IZ" Then
'   RS.Fields("sifraplac") = Me.StatusBar1.Panels(5).text
'   End If
          
    rs.Update
    End If
    
  '*pogleda zadnjo
   If rs.State = 1 Then rs.Close
   rs.Open "select * from trenutna where  pozicija='" & levi_pres(LTrim(str(fgtrial.Rows - 1)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   xro = 1
   If Not rs.EOF Then
'
      For i = fgtrial.FixedCols To fgtrial.Cols - 1
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "POZICIJA" Then
       rs.Fields(Trim(fgtrial.TextMatrix(0, i))) = levi_pres(LTrim(str(Val(fgtrial.TextMatrix(fgtrial.Rows - 1, i)))), 4)
       Else
       imepol = Trim(fgtrial.TextMatrix(0, i))
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "VISINA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "UR" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "DAT_PRE" Then
       imepol = "dat_k"
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
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "PROSTA" Then
       imepol = "sifraplac"
       End If
         If UCase(Trim(fgtrial.TextMatrix(0, i))) = "MARZA" Then
       imepol = "x"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "MPC" Then
       imepol = "y"
       End If
       
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "LES" Then
       imepol = "kopija"
       End If
       If UCase(Trim(fgtrial.TextMatrix(0, i))) = "ZALOGA" Then
       imepol = "x"
       End If
      If Not imepol = "" Then
      If Not UCase(imepol) = "EM" Then
      If Not UCase(imepol) = "DAT_K" Then
      
        rs.Fields(imepol) = fgtrial.TextMatrix(fgtrial.Rows - 1, i)
          Else
        rs.Fields(imepol) = ctod(fgtrial.TextMatrix(fgtrial.Row, i))
        
      End If
      End If
        End If
       End If
      Next i
         rs.Fields("datum") = Me.DTPicker1.Value
         rs.Fields("skl") = Me.sklad.BoundDatax
          rs.Fields("znes") = (rs.Fields("kol") * rs.Fields("cena")) * (1 - (rs.Fields("pop") / 100))
          rs.Fields("doza") = 1
  ' If Left(Me.dok.Caption, 2) = "IZ" Then
  ' RS.Fields("sifraplac") = Me.StatusBar1.Panels(5).text
  ' End If
       
    rs.Update
    End If
  
    
    txtNewData.Visible = False
        txtNewData.Text = ""
   '     If fgtrial.TextMatrix(fgtrial.Rows - 1, 1) = "" Then
   '     Else
   '     fgtrial.Rows = fgtrial.Rows + 1
   '     End If
Dim fakkx As Long
fakkx = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")

 If Trim(fgtrial.TextMatrix(fgtrial.Rows - 1, 1)) <> "" Then
' MsgBox Trim(fgtrial.TextMatrix(fgtrial.Rows - 1, 1))
   If rs.State = 1 Then rs.Close
   rs.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic

   rs.AddNew
   rs.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")
   rs.Fields("tip_dok") = Left(Me.dok.Caption, 2)
   rs.Fields("id_dok") = Mid(Me.dok.Caption, 3)
   rs.Fields("datum") = Me.DTPicker1.Value
   rs.Fields("skl") = Me.sklad.BoundDatax
   rs.Fields("faktor") = fakkx
   rs.Fields("pozicija") = levi_pres(LTrim(str(Val(Getnazi("select count(pozicija) as cc from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")) + 1)), 4)
      rs.Update
  End If
'DoColumnSort
fgtrial.Redraw = True

refre
dell = 1
   ' fgtrial.DataSource = RS
'    MsgBox ""
    lstRow = fgtrial.Rows - 1
    fgtrial.Row = lstRow
    fgtrial.Col = coollsi
     fgtrial.TextMatrix(lstRow, 0) = lstRow
     fgtrial.TextMatrix(lstRow, 1) = ""
    ' opiss.SetFocus
     fgtrial.SetFocus
     
'Me.Text1.SetFocus
'     fgtrial.SetFocus
'DoColumnSort
'refre
bbr:
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
            id = fgtrial.Text
            dt = ""
            
            For k = 1 To Cols - 1
                fgtrial.Col = k
                dt = dt & "," & fgtrial.Text
            Next k
            
            fgtrial.Col = k
            ftFg = Val(fgtrial.Text)
        
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
    
    If Val(fgtrial.Text) > 2 Then
        fgtrial.Text = Val(fgtrial.Text) - 3
    Else
        MsgBox "The Record Is Not Deleted"
    End If
    
    fgtrial.Col = curCol
End Sub

Private Sub cmdDel_Click()
Dim rro As Integer
rro = fgtrial.Row
myConection.Execute ("delete  from trenutna where (pozicija)='" & levi_pres(LTrim(str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
fgtrial.clear
 If Getnazi("select tekst from dokm where atribut='" & levi_pres(LTrim(str(fgtrial.Row)), 4) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
 myConection.Execute ("delete  from dokm where atribut='" & levi_pres(LTrim(str(fgtrial.Row)), 4) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and poz=0")
 End If
    'Dim s As Integer
    'For s = 1 To fgtrial.Rows - 1
    'fgtrial.TextMatrix(s, 0) = LTrim(Str(s))
    'Next
    If rs.State = 1 Then rs.Close
    rs.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic
   If Not rs.EOF Then
    rs.MoveFirst
    
    Dim s As Integer
    s = 1
    Do While Not rs.EOF
    If Getnazi("select tekst from dokm where atribut='" & rs.Fields("pozicija") & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and poz=0") <> "" Then
    myConection.Execute ("update dokm set atribut='" & levi_pres(LTrim(str(s)), 4) & "' where atribut='" & rs.Fields("pozicija") & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and poz=0")
    End If
    rs.Fields("pozicija") = levi_pres(LTrim(str(s)), 4)
    s = s + 1
    rs.MoveNext
    Loop
   
  '  myConection.Execute ("update  trenutna set pozicija='" & LTrim(Str(recno())) & "' where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
    refre
    'DoColumnSort
    End If
    If Getnazi("select tip_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") = "" Then
    Dim fakk As Long
    fakk = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
    myConection.Execute ("insert into trenutna (tip_dok,id_dok,pozicija,faktor,uporabnik) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','   1'," & fakk & ",'" & Getnazi("select up from users where username1='" & UPORABNIK & "'") & "')")
    refre
    'DoColumnSort
     End If
  '  myConection.Execute ("update  trenutna set pozicija='" & LTrim(Str(recno())) & "' where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   If rro > fgtrial.Rows - 1 Then
   
   fgtrial.Row = fgtrial.Rows - 1
   Else
   fgtrial.Row = rro
   End If
    End Sub

Private Sub cmdFetch_Click()
    If rs.State = 0 And ftchFlag = False Then
        ftchFlag = True
      '  rs.Source = "Select * from utna"
        rs.Open "select pozicija,sifra,naziv,cena,kol,znes,x,y,chk_fix  from trenutna", myConection, adOpenDynamic, adLockOptimistic
        'Set fgTrial.DataSource = rs.Source
        
        If rs.EOF = False Then
            rs.MoveFirst
            For i = 1 To rs.RecordCount - 1
                If i <> fgtrial.Rows - 1 Then
                    fgtrial.Rows = fgtrial.Rows + 1
                End If
                fgtrial.Row = i
                For J = 0 To fgtrial.Cols - 2
                    fgtrial.Col = J
                    fgtrial.Text = rs.Fields(J)
                Next J
                ' to set last col as fetch flag - 0
                fgtrial.Col = J
                fgtrial.Text = 0
                rs.MoveNext
            Next i
        End If
    End If
End Sub

Private Sub Command1_Click()
MsgBox coollem
If Me.Timer1.Enabled = True Then
MsgBox "da"
End If

End Sub

Private Sub DTPicker2_Change()
Me.DTPicker3.Value = Me.DTPicker2.Value + Val(Me.Text5.Text)
End Sub

Private Sub DTPicker3_Change()
Me.Text5.Text = Me.DTPicker3.Value - Me.DTPicker2.Value
End Sub

Private Sub fgTrial_Click()
 
    clk = False
    curRow = fgtrial.Row
    curCol = fgtrial.Col
    msgFlag = False
  If fgtrial.Col = 0 Then
    xopis = levi_pres(LTrim(str(fgtrial.Row)), 4)
    xid_dok = Trim(dok.Caption)
    Dialog.Show vbModal
    fgtrial.Col = 1
  End If
  If fgtrial.Col = collchk Then
  If Trim(fgtrial.TextMatrix(curRow, curCol)) = "c" Then
  fgtrial.Text = "b"
  Else
  fgtrial.Text = "c"
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
    If fgtrial.Col = coollce Then
        If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='VNMA'") = "D" Then
         cmdAdd_Click
       '
        vnoscen.cene Me.dok.Caption, xpozi
          cmdAdd_Click
         fgtrial.Col = coollsi
         'refre
        End If
    Else
        If fgtrial.Col < fgtrial.Cols - 1 Then
          fgtrial.Col = fgtrial.Col + 1
   

      Else
       txtNewData.Visible = False
       
       
    ' cmdAdd_Click
     ' MsgBox fgtrial.Col
      'fgtrial.Col = coollsi
      End If
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
      If fgtrial.Col = coollce Then
        If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='VNMA'") = "D" Then
         cmdAdd_Click
       '
        vnoscen.cene Me.dok.Caption, xpozi
          cmdAdd_Click
         fgtrial.Col = coollsi
         'refre
        End If
Else
        
        txtNewData.Text = Chr(KeyAscii)
        txtNewData.Move fgtrial.CellLeft + fgtrial.Left - 10, fgtrial.CellTop + _
                        fgtrial.Top - 10, fgtrial.CellWidth + 10, fgtrial.CellHeight - 40
        
        'to set col no to previous value
        txtNewData.Visible = True
'        Me.Text1.SetFocus
        txtNewData.SetFocus
        End If
    'Else
    '    MsgBox "Double Click To Make Edit Mode Active"
    'End If
    End If
End Sub
Private Sub fgTrial_GotFocus()

If fgtrial.Rows = 2 And fgtrial.TextMatrix(1, 0) = "" Then

fgtrial.Row = 1
fgtrial.Col = 0

fgtrial.Text = 1
fgtrial.Col = 1

End If


End Sub
Private Sub fgTrial_LeaveCell()

 If fgtrial.Col = coollem Then
    If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='AVTOOP'") = "D" Then
     xopis = levi_pres(LTrim(str(fgtrial.Row)), 4)
    xid_dok = Trim(dok.Caption)
   If IsLoaded(Dialog) Then
   Else
    Dialog.Show vbModal
    fgtrial.SetFocus
    End If
    
    
    End If
    fgtrial.Col = fgtrial.Col + 1
    End If

     
    If txtNewData.Visible = False Then
        Exit Sub
    Else
        'If IsNumeric(txtNewData.text) Then
        '    fgtrial.text = Str(txtNewData.text)
        'Else
            fgtrial.Text = txtNewData.Text
        'End If
        txtNewData.Visible = False
        txtNewData.Text = ""
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
Dim ssks, sskp As Double
Dim axsi As String
axsi = ""
  Dim das, dodx
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
  
axsi = fgtrial.TextMatrix(fgtrial.Row, 1)
If axsi <> "" Then
Me.StatusBar1.Panels(1).Text = Getnazi("select madanazi from mada where madasifr='" & axsi & "'") & " " & Getnazi("select madaenme from mada where madasifr='" & axsi & "'")
     If Getnazi("select sum(kol) as ss from zaloga where sifra='" & axsi & "'") <> "" Then
     ssks = Getnazi("select sum(kol) as ss from zaloga where sifra='" & axsi & "'")
     End If
     If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & axsi & "' and datum<=#" & dodx & "#") <> "" Then
     sskp = Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & axsi & "' and datum<=#" & dodx & "#")
     End If
     Me.StatusBar1.Panels(3).Text = FormatNumber(ssks, 3)
     Me.StatusBar1.Panels(5).Text = FormatNumber(sskp, 3)
     If Left(Me.dok.Caption, 2) = "IZ" Then
     fgtrial.TextMatrix(fgtrial.Row, coollpro) = Me.StatusBar1.Panels(5).Text
     End If
End If
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
collchk = 0
coollsi = 0
coollna = 0
coollce = 0
coollko = 0
coollles = 0
coollmarz = 0
coollmpc = 0
coollstek = 0
Me.DTPicker1 = Date
If normati = "" Then
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from glavna where tip_dok='" & Trim(tip_dok) & "'")) + 1, 6)
Else
Me.dok.Caption = normati
normati = ""
End If
Dim upor As String
upor = Getnazi("select up from users where username1='" & UPORABNIK & "'")

If Getnazi("select id_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and uporabnik<>'" & upor & "'") <> "" Then
If Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "' and uporabnik='" & upor & "'") <> "" Then
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "' and uporabnik='" & upor & "'")), 6)
Else
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "'")) + 1, 6)
End If
End If

Me.sklad.BoundDatax = Getnazi("select skl from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Me.Text4.Text = Getnazi("select dol_ce from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
If ma_ured <> 0 Then
Me.dok.Caption = Trim(tip_dok) & Trim(frmControlMain.MSHFlexGrid1.Text)
'napolni
Else

End If

If Left(Me.dok.Caption, 2) = "IZ" Then
dokl = "select " & Replace(Getnazi("select polja from dokumenti where tip_dok='" & tip_dok & "'"), "kol", "kol,format(sifraplac,'0.000') as prosta") & " from trenutna  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija"
Else
dokl = "select " & Getnazi("select polja from dokumenti where tip_dok='" & tip_dok & "'") & " from trenutna  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija"
End If


Dim y As Integer
y = 0

For y = 0 To 7
Me.do_da(y).Caption = Getnazi("select dod" & Trim(str(y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
If Me.do_da(y).Caption = "" Then
Me.UserControl11(y).Visible = False
Else
Me.UserControl11(y).sSQL = Getnazi("select sqdo" & Trim(str(y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Me.UserControl11(y).polje = Getnazi("select dpo" & Trim(str(y)) & " from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
End If
Next
'napolni

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
myConection.Execute ("insert into dokm select atribut,tip_dok,id_dok,tekst,999 as poz from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by poz")
End If
napolni

Call GetNewConnection2
If Getnazi("select tip_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") = "" Then
Dim fakk As Long
fakk = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
'Set Rs1 = New Recordset
myConection.Execute ("insert into trenutna (tip_dok,id_dok,pozicija,faktor,uporabnik,stdok) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','   1'," & fakk & ",'" & Getnazi("select up from users where username1='" & UPORABNIK & "'") & "','" & Pblagajna & "')")
End If
If kosovni = 1 Then
napolni
End If

refre
'DoColumnSort

ReSizeForm Me
izja = 1
Set Rs1 = Nothing
Set DCON = Nothing
 Call WheelHook(Me.hwnd)
 If Left(Me.dok.Caption, 3) = "NTX" Or Left(Me.dok.Caption, 3) = "NTY" Then
 Me.Text3.Visible = True
 'Me.Text4.Visible = True
Me.Text3.Text = Getnazi("select stdok from nabasif where tip_dok='NT' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'Me.Text4.text = Getnazi("select y from nabasif where tip_dok='NT' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 End If

'datumi
If Getnazi("select tekst from dokm where atribut='DUR' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.DTPicker2.Value = Getnazi("select tekst from dokm where atribut='DUR' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Else
Me.DTPicker2.Value = Date
End If
If Getnazi("select tekst from dokm where atribut='DNI' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.Text5.Text = Val(Getnazi("select tekst from dokm where atribut='DNI' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'"))
Else
Me.Text5.Text = 0
End If
If Getnazi("select tekst from dokm where atribut='VAL' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then
Me.DTPicker3.Value = Getnazi("select tekst from dokm where atribut='VAL' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'")
Else
Me.DTPicker3.Value = Date
End If
If imedn <> "" Then

Me.Text1.Text = Getnazi("select opis from glavna where tip_dok='DN' and id_dok='" & Trim(imedn) & "'")
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
If Left(Getnazi("select tvorba from dokumenti where tip_dok='DO'"), 1) = "=" Then

Me.Check1.Enabled = False
Me.Zapis.Enabled = False
Else
Me.Check1.Enabled = True
Me.Zapis.Enabled = True

End If
Else
Me.Check1.Enabled = True
Me.Zapis.Enabled = True

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

' opisi pozicij


coollsi = 0
coollce = 0
coollchk = 0
coollko = 0
coollzn = 0
coollpop = 0
coollles = 0
coollstek = 0
coollzn = 0
coollpro = 0
coollmarz = 0
coollmpc = 0

    For i = fgtrial.Col To fgtrial.Cols - 1
        Dim asx As String
        
        asx = LCase(Trim(fgtrial.TextMatrix(0, i)))
        If asx = "sifra" Then
        coollsi = i
        End If
        If UCase(asx) = "EM" Then
        coollem = i
        End If
        If UCase(asx) = "DAT_PRE" Then
        coolldat_k = i
        End If
        If UCase(asx) = "UR" Then
        coollur = i
        End If
         If asx = "pop" Then
        coollpop = i
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
        If UCase(asx) = "PROSTA" Then
        coollpro = i
        End If
         If UCase(asx) = "MARZA" Then
        coollmarz = i
        End If
         If UCase(asx) = "MPC" Then
        coollmpc = i
        End If
        If UCase(asx) = "LES" Then
        coollles = i
        End If
        If UCase(asx) = "ZALOGA" Then
        coollzal = i
        End If
        If Left(asx, 3) = "chk" Then
        collchk = i
        
        End If
        
        If asx = "naziv" Then
        coollna = i
      
        End If
         If asx = "znes" Then
        coollzn = i
        
        End If
        If asx = "cena" Then
        coollce = i
       
        End If
        If asx = "kol" Then
        coollko = i
       
        End If
         
   
       Next i
If Left(Me.dok.Caption, 2) = "IZ" Then
Dim das, dodx
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
Dim rrrt As New ADODB.Recordset
If rrrt.State = 1 Then rrrt.Close
rrrt.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenDynamic, adLockOptimistic
If Not rrrt.EOF Then
Do While Not rrrt.EOF
If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & rrrt.Fields("sifra") & "' and datum<=#" & dodx & "#") <> "" Then
rrrt.Fields("sifraplac") = FormatNumber(Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & rrrt.Fields("sifra") & "' and datum<=#" & dodx & "#"), 3)
rrrt.Update
End If
rrrt.MoveNext
Loop
End If
End If
refre
fgtrial.ColWidth(coollna) = 6000
fgtrial.ColWidth(collchk) = 1000
fgtrial.ColAlignment(coollchk) = flexAlignCenterCenter


If Not IsNumber(Mid(Me.dok.Caption, 4, 1)) Then
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from glavna where tip_dok='" & Trim(tip_dok) & "'")) + 1, 6)
If Getnazi("select id_dok from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and uporabnik<>'" & upor & "'") <> "" Then
If Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "' and uporabnik='" & upor & "'") <> "" Then
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "' and uporabnik='" & upor & "'")), 6)
Else
Me.dok.Caption = Trim(tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from trenutna where tip_dok='" & Trim(tip_dok) & "'")) + 1, 6)
End If
End If

End If
If tip_dok = "PT" Then
Me.List1.Visible = True
Me.sturdod.Visible = True
Me.UpDown1.Visible = True
Me.stur.Visible = True
Filipotne List1, "select madasifr,madanazi,madazacs from mada where tip_art='POT'"
End If
If tip_dok = "FA" Then
Me.av_zne.Visible = True
Me.Label6.Visible = True
If Me.UserControl11(7).BoundDatax <> "" Then
Dim vezz As String
vezz = "'" & Replace(Me.UserControl11(7).BoundDatax, ",", "','") & "'"
Me.av_zne.Caption = Getnumb("select sum(znes) as zne from nabasif where tip_dok='AR' and id_dok in (" & vezz & ")")

End If
'Me.av_la.Visible = True
End If
Me.fgtrial.Col = 0


  cmdAdd_Click
      fgtrial.Col = coollsi
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
myConection.Execute ("delete  from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
    
    If rs.State = 1 Then
        rs.Close
    End If
'    cn.Close
If Left(Me.dok.Caption, 2) <> "NT" Then
osve = 1
End If
    Call WheelUnHook(Me.hwnd)
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
Sub refre()
'On Error GoTo bbb:
'Me.WindowState = 1
Dim cooo As Integer
Dim znesav As Double
znesav = Me.av_zne.Caption
cooo = fgtrial.Col

If Rs1.State = 1 Then Rs1.Close
Rs1.Open dokl, myConection, adOpenDynamic, adLockOptimistic
fgtrial.Redraw = False
'If Getnazi("select sum(znesek) as zne from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
Dim zxn As Double
zxn = 0
zxn = Getnazi("select sum(znes) as zne from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'MsgBox zxn
Me.Label5.Caption = FormatNumber(zxn - znesav, 2)
'End If
Set fgtrial.DataSource = Rs1
        
'Set Rs1 = DCON.Execute(SQL)
'ssqq = SQL
       
  
     Dim lngX As Long
        fgtrial.Col = coollce
       
       ' cene
       
        'lngX = 1
        'While lngX + 1 < fgtrial.Rows
            
            ' fgtrial.TextMatrix(lngX, coollce) = FormatNumber(fgtrial.TextMatrix(lngX, coollce), 4)
             
         '   lngX = lngX + 1
        'Wend
  Dim cenn As Double
  Dim dass  As Date
  
        With fgtrial
       ' MsgBox fgtrial.TextMatrix(lCount, coollce)
        '.Redraw = False ' makes it about 10x faster !
        For lcount = .FixedRows To .Rows - 1
           'cena
            fgtrial.TextMatrix(lcount, coollem) = Getnazi("select madaenme from mada where madasifr='" & fgtrial.TextMatrix(lcount, coollsi) & "'")
           If coollce <> 0 Then
           If fgtrial.TextMatrix(lcount, coollce) <> "" Then
           cenn = Replace(fgtrial.TextMatrix(lcount, coollce), ".", ",")
             fgtrial.TextMatrix(lcount, coollce) = FormatNumber(cenn, 4)
             .ColAlignment(coollce) = flexAlignRightCenter
             End If
             End If
            'kol
             If coollko <> 0 Then
             If fgtrial.TextMatrix(lcount, coollko) <> "" Then
            cenn = Replace(Replace(fgtrial.TextMatrix(lcount, coollko), ",", ""), ".", ",")
             fgtrial.TextMatrix(lcount, coollko) = FormatNumber(cenn, 3)
             .ColAlignment(coollko) = flexAlignRightCenter
             End If
             End If
             If coolldat_k <> 0 Then
             If fgtrial.TextMatrix(lcount, coolldat_k) <> "" Then
            dass = fgtrial.TextMatrix(lcount, coolldat_k)
             fgtrial.TextMatrix(lcount, coolldat_k) = dtoc(dass)
             .ColAlignment(coolldat_k) = flexAlignRightCenter
             End If
             End If
            'znes
             If coollzn = "" Then
             coollzn = 0
             End If
             If coollzn <> 0 Then
             If fgtrial.TextMatrix(lcount, coollzn) <> "" Then
            cenn = Replace(fgtrial.TextMatrix(lcount, coollzn), ".", ",")
             fgtrial.TextMatrix(lcount, coollzn) = FormatNumber(cenn, 4)
             .ColAlignment(coollzn) = flexAlignRightCenter
             End If
             End If
             If coollmarz = "" Then
             coollmarz = 0
             End If
             If coollmarz <> 0 Then
             If fgtrial.TextMatrix(lcount, coollmarz) <> "" Then
            cenn = Replace(fgtrial.TextMatrix(lcount, coollmarz), ".", ",")
             fgtrial.TextMatrix(lcount, coollmarz) = FormatNumber(cenn, 2)
             .ColAlignment(coollmarz) = flexAlignRightCenter
             End If
             End If
               If coollmpc = "" Then
             coollmpc = 0
             End If
             If coollmpc <> 0 Then
             If fgtrial.TextMatrix(lcount, coollmpc) <> "" Then
            cenn = Replace(fgtrial.TextMatrix(lcount, coollmpc), ".", ",")
             fgtrial.TextMatrix(lcount, coollmpc) = FormatNumber(cenn, 2)
             .ColAlignment(coollmpc) = flexAlignRightCenter
             End If
             End If
             
             
             If coollzal <> 0 Then
             If fgtrial.TextMatrix(lcount, coollzal) <> "" Then
            cenn = Replace(fgtrial.TextMatrix(lcount, coollzal), ".", ",")
             fgtrial.TextMatrix(lcount, coollzal) = FormatNumber(cenn, 4)
             .ColAlignment(coollzal) = flexAlignRightCenter
             End If
             End If
              If coollpop <> 0 Then
              If fgtrial.TextMatrix(lcount, coollpop) <> "" Then
              cenn = Replace(fgtrial.TextMatrix(lcount, coollpop), ".", ",")
             fgtrial.TextMatrix(lcount, coollpop) = FormatNumber(cenn, 2)
             .ColAlignment(coollpop) = flexAlignRightCenter
             End If
             End If
             If Left(Me.dok.Caption, 2) = "DN" Then
               ' If fgtrial.Col = coollko Then
               ' MsgBox ""
                   If Trim(fgtrial.TextMatrix(lcount, coollko)) = 0 Then
                   If Trim(fgtrial.TextMatrix(lcount, coollsi)) <> "" Then
                   fgtrial.Row = lcount
                   fgtrial.Col = coollko
                   fgtrial.CellBackColor = 255
                   End If
                '   End If
                   End If
                End If
        Next
      
        
        
        Dim xro As Integer
If Left(Me.dok.Caption, 2) = "IZ" Then
'barvam pozicije
'        If Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") < 0 Then
'         If Getnazi("select pozicija from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
         fgtrial.Col = coollsi
         Dim kolic, proo As Double
         kolic = 0
         proo = 0
         For xro = fgtrial.FixedRows To fgtrial.Rows - 1
'fgtrial_RowColChange
fgtrial.Row = xro
'        If Trim(fgtrial.TextMatrix(xro, coollsi)) <> "" Then
'        If Getnazi("select madanazi from mada where madasifr='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
'       If Getnazi("select sum(kol*faktor) as ss from nabasif where sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "'") <> "" Then
'        Dim das, dodx
'        das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
'        dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
'        If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") <> "" Then
'        If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and sifra='" & Trim(fgtrial.TextMatrix(xro, coollsi)) & "' and datum<=#" & dodx & "#") < 0 + fgtrial.TextMatrix(fgtrial.Row, coollko) Then
       kolic = fgtrial.TextMatrix(xro, coollko)
      'MsgBox coollpro
       If (fgtrial.TextMatrix(xro, coollpro)) = "" Then
       proo = 0
       Else
       proo = fgtrial.TextMatrix(xro, coollpro)
       End If
      
        If kolic > proo Then
       
           ' fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            'Me.fgtrial.Refresh
       'MsgBox fgtrial.Row
            Else
            fgtrial.CellBackColor = &HFFFFFF
            
       End If
'       Else
       
'            fgtrial.Col = fgtrial.FixedCols
'            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
'            fgtrial.CellBackColor = &HC0C0FF
            'Me.fgtrial.Refresh
'        End If
'       End If
'       End If
'       End If
       
       'fgtrial.CellBackColor = &HFFFFFF
    fgtrial.Col = coollsi
    
       'End If
       Next xro
       End If
       'End If
       'End If
        
        
         
        '.Redraw = True ' dont forget to do this !
        End With

   If collchk <> 0 Then
   With fgtrial
   '.Redraw = False
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
          cenn = Replace(Getnazi("select sum(znes) as znes from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"), ".", ",")
             Me.skup.Caption = FormatNumber(cenn, 2)
             LoadFlexGridColumnWidths fgtrial, "nabava"
'fgtrial.Col = cooo
fgtrial.Redraw = True
bbb:
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
        
        If UCase(asx) = "MARZA" Then
        coollmarz = i
        'Exit For
        End If
          If UCase(asx) = "MPC" Then
        coollmpc = i
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
        
        If UCase(asx) = "ZALOGA" Then
        coollzal = i
        End If
        If UCase(asx) = "STEKLO" Then
        coollstek = i
        End If
        If UCase(asx) = "LES" Then
        coollles = i
        End If
      If UCase(asx) = "DAT_PRE" Then
        coolldat_k = i
        End If
        If UCase(asx) = "UR" Then
        coollur = i
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
            'Me.fgtrial.Refresh
            Else
            fgtrial.CellBackColor = &HFFFFFF
            
       End If
       Else
       
            fgtrial.Col = fgtrial.FixedCols
            fgtrial.ColSel = fgtrial.Cols() - fgtrial.FixedCols - 1
           
            fgtrial.CellBackColor = &HC0C0FF
            'Me.fgtrial.Refresh
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

For X = fgtrial.Row To fgtrial.Rows - 1
For i = fgtrial.Col To fgtrial.Cols - 1
fgtrial.Col = i
fgtrial.Row = X
If UCase(fgtrial.TextMatrix(X, i)) Like iskan Then
fgtrial.Redraw = True
Exit Sub
End If
Next i
fgtrial.Col = 0
Next X
fgtrial.Redraw = True

End Sub

Private Sub Label2_Click()
Dim sss As String
sss = "update trenutna set znes=kol*(cena*(1-(pop/100))) where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        
       
        myConection.Execute (sss)
        'DoColumnSort
        refre
End Sub

Private Sub LaVolpeButton1_Click()
Dim sss As String
If rs.State = 1 Then rs.Close
rs.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
If Not rs.EOF Then
rs.MoveFirst
End If
Do While Not rs.EOF
MsgBox Getnazi("select madampcd from mada where madasifr='" & rs.Fields("sifra") & "'")
If Getnazi("select madampcd from mada where madasifr='" & rs.Fields("sifra") & "'") <> "" Then
rs.Fields("cena") = FormatNumber(Getnazi("select madampcd from mada where madasifr='" & rs.Fields("sifra") & "'"), 4)
rs.Fields("znes") = FormatNumber(rs.Fields("cena") * rs.Fields("kol"), 4)
End If
rs.Update
rs.MoveNext
Loop
       
       ' myConection.Execute (sss)
        'DoColumnSort
        refre
End Sub

Private Sub List1_DblClick()
Call sturdod_Click
End Sub

Private Sub prekin_Click()
'If Left(Me.dok.Caption, 2) <> "FA" Then
myConection.Execute ("delete  from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and poz=0")
myConection.Execute ("update dokm  set poz=0 where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'End If
Unload Me
End Sub

Private Sub shran_Click()
SaveFlexGridColumnWidths fgtrial, "nabava"
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Left(Me.dok.Caption, 2) = "DN" Then
Dim norma, stek, lesi As String
Dim koli As Long
Dim xrsn As New ADODB.Recordset
Dim xox, yoy, zapp, dkr As Integer
Dim kol, XX, YY As Long
dkr = 1

imedn = frmControlMain.MSHFlexGrid1.Text
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
Rsa.Fields("kol") = Getnazi("select sum(zap) as x from normati")
Rsa.Update

End If
preg.Show
End If
End Sub

Private Sub sturdod_Click()
If Me.txtNewData.Visible = True Then
Sendkeys "{ENTER}"
End If
'Call cmdAdd_Click
'cmdAdd_Click
Dim xdatt, xpozzx As String
xdatt = ""
xpozzx = "   1"
If Getnazi("select max(dat_k) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
xdatt = dtoc(getdate("select max(dat_k) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'") + 1)
'MsgBox (getdate("select max(dat_k) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'"))
Else
xdatt = dtoc(Me.DTPicker1.Value)
End If
If Getnazi("select max(pozicija) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
xpozzx = levi_pres(LTrim(str(Val(Getnazi("select max(pozicija) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'")) + 1)), 4)
End If
If xdatt = "31.12.1899" Then
xdatt = dtoc(Me.DTPicker1.Value)
End If
'MsgBox (UCase(Format(ctod(xdatt), "DDDD")))
If UCase(Format(ctod(xdatt), "DDDD")) = "NEDELJA" Then
xdatt = dtoc(ctod(xdatt) + 1)
End If
'fgtrial.TextMatrix(Me.fgtrial.Row, coolldat_k) = xdatt
'fgtrial.TextMatrix(Me.fgtrial.Row, coollur) = Me.stur.text
'fgtrial.Col = coollsi
''SendKeys Trim(Left(Me.List1.text, 10))
''Me.txtNewData.Visible = True
''txtNewData_LostFocus
''Me.txtNewData.text = Trim(Left(Me.List1.text, 10))
''SendKeys "{ENTER}"
''fgtrial.TextMatrix(fgtrial.Row, coollsi) = Trim(Left(Me.List1.text, 10)) & " "
'fgtrial.TextMatrix(fgtrial.Row, coollna) = Trim(Getnazi("select madanazi from mada where madasifr='" & Trim(Left(Me.List1.text, 10)) & "'")) & " "
'fgtrial.TextMatrix(fgtrial.Row, coollem) = Trim(Getnazi("select madaenme from mada where madasifr='" & Trim(Left(Me.List1.text, 10)) & "'")) & " "
'fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(Getnumb("select madampcd from mada where madasifr='" & Trim(Left(Me.List1.text, 10)) & "'"), 4)
'fgtrial.TextMatrix(fgtrial.Row, coollko) = FormatNumber(Getnumb("select madazacs from mada where madasifr='" & Trim(Left(Me.List1.text, 10)) & "'"), 2)
''fgtrial.Col = 5
''txtNewData_LostFocus
'Sleep 500
'SendKeys "{ENTER}"
'fgtrial.TextMatrix(fgtrial.Row, coollsi) = Trim(Left(Me.List1.text, 10)) & " "

Dim rsta As New ADODB.Recordset
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
rsta.Open "select * from trenutnA WHERE TIP_DOK='" & tii & "' AND ID_DOK='" & idd & "'", myConection, adOpenDynamic, adLockOptimistic

'Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
If Val(xpozzx) = 1 Then
Else
rsta.AddNew
End If
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = Trim(Left(Me.List1.Text, 10))
rsta.Fields("naziv") = Trim(Getnazi("select madanazi from mada where madasifr='" & Trim(Left(Me.List1.Text, 10)) & "'"))
rsta.Fields("cena") = FormatNumber(Getnumb("select madampcd from mada where madasifr='" & Trim(Left(Me.List1.Text, 10)) & "'"), 3)
rsta.Fields("kol") = FormatNumber(Getnumb("select madazacs from mada where madasifr='" & Trim(Left(Me.List1.Text, 10)) & "'"), 2)
rsta.Fields("znes") = FormatNumber(Getnumb("select madampcd from mada where madasifr='" & Trim(Left(Me.List1.Text, 10)) & "'"), 3) * FormatNumber(Getnumb("select madazacs from mada where madasifr='" & Trim(Left(Me.List1.Text, 10)) & "'"), 2)
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("x") = Me.stur.Text
rsta.Fields("dat_k") = ctod(xdatt)
rsta.Fields("faktor") = 0
rsta.Fields("pozicija") = xpozzx
rsta.Fields("skl") = Me.sklad.BoundDatax
rsta.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")

rsta.Update
If Val(Me.stur.Text) > 9 Then
xpozzx = levi_pres(LTrim(str(Val(Getnazi("select max(pozicija) as xx from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok ='" & Mid(Me.dok.Caption, 3) & "'")) + 1)), 4)
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = "300"
rsta.Fields("naziv") = Trim(Getnazi("select madanazi from mada where madasifr='300'"))
rsta.Fields("cena") = FormatNumber(Getnumb("select madampcd from mada where madasifr='300'"), 3)
rsta.Fields("kol") = FormatNumber(1, 2)
rsta.Fields("znes") = FormatNumber(Getnumb("select madampcd from mada where madasifr='300'"), 3)
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("x") = Me.stur.Text
rsta.Fields("dat_k") = ctod(xdatt)
rsta.Fields("faktor") = 0
rsta.Fields("pozicija") = xpozzx
rsta.Fields("skl") = Me.sklad.BoundDatax
rsta.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")

rsta.Update
End If
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
'RSS.MoveNext
'Loop

refre
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
        sss = "update trenutna set pop=" & Replace(Me.Text2.Text, ",", ".") & " where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        'MsgBox sss
        myConection.Execute (sss)
       sss = "update trenutna set znes=kol*(cena*(1-(pop/100))) where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
        
       
        myConection.Execute (sss)
        'DoColumnSort
        refre
        End If
End Sub

Private Sub Text3_LostFocus()
If Me.Text3.Text = "" Then
MsgBox "Obvezen vnos!!"
Me.Text3.SetFocus

End If
End Sub

Private Sub Text5_Change()
If Me.Text5.Text = "" Then
Me.Text5.Text = "0"
End If
Me.DTPicker3.Value = Me.DTPicker2.Value + Val(Me.Text5.Text)
End Sub

Private Sub UpDown1_DownClick()
Me.stur.Text = Me.stur.Text - 1
End Sub

Private Sub UpDown1_UpClick()
Me.stur.Text = Me.stur.Text + 1
End Sub

Private Sub UserControl11_LostFocus(Index As Integer)
On Error GoTo bbnn
If Me.UserControl11(0).BoundDatax <> "" Then


If Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'") <> "" Then
If Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'") = 0 Then
Me.Text5.Text = Val(Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='PLDNI'"))
Else
Me.Text5.Text = Val(Getnazi("select maxlimit from partner where naziv='" & Me.UserControl11(0).BoundDatax & "'"))
End If
Text5_Change
End If
End If
'If Me.UserControl11(7).BoundDatax <> "" Then

If Left(Me.do_da(7).Caption, 4) = "Veza" Then
Dim vezz As String
vezz = "'" & Replace(Me.UserControl11(7).BoundDatax, ",", "','") & "'"
Me.av_zne.Caption = Getnumb("select sum(znes) as zne from nabasif where tip_dok='AR' and id_dok in (" & vezz & ")")
'MsgBox (Getnazi("select sum(znes) as zne from nabasif where tip_dok='AR' and id_dok in (" & vezz & ")"))
myConection.Execute ("update glavna set dod7='" & Me.UserControl11(7).BoundDatax & "' where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'End If
End If
bbnn:
End Sub

Private Sub Uvoz_Click()
 Me.mnu_uvoz.Enabled = True
      PopupMenu Me.mnu_uvoz
End Sub
Private Sub mnuvozotv_Click()

Dim RSS As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
RSS.Open "select sifra,min(naziv) as naziv,cena,sum(prosta) as prosta from zaloga where prosta>0 group by sifra,cena order by sifra", myConection, adOpenDynamic, adLockOptimistic
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
Dim xpo As Integer
xpo = 1
rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = RSS.Fields("sifra")
rsta.Fields("naziv") = RSS.Fields("naziv")
rsta.Fields("cena") = RSS.Fields("cena")
rsta.Fields("kol") = RSS.Fields("prosta")
rsta.Fields("znes") = RSS.Fields("cena") * RSS.Fields("prosta")
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("pozicija") = levi_pres(LTrim(str(xpo)), 4)
'rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
xpo = xpo + 1
Loop
Call cmdAdd_Click
MsgBox "Konèano"
End Sub
Private Sub mnuvozgost_Click()
Dim RSS As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
RSS.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)

rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = RSS.Fields("madasifr")
rsta.Fields("naziv") = RSS.Fields("madanazi")
rsta.Fields("cena") = RSS.Fields("madanabc")
rsta.Fields("kol") = RSS.Fields("madazalo")
rsta.Fields("znes") = RSS.Fields("madazalo") * RSS.Fields("madanabc")
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
Loop
End Sub
Private Sub mnuuvozin_Click()
Dim RSS As New ADODB.Recordset

Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
Dim ain As String
ain = InputBox("Vnesi številko inventure (000001)!")
If Getnazi("select id_dok from nabasif where stdok='IN" & ain & "'") <> "" Then
MsgBox ("Ta inventura je že bila uvožena v NA" & Getnazi("select id_dok from nabasif where stdok='IN" & ain & "'") & "!")
Exit Sub
End If
RSS.Open "select * from nabasif where tip_dok='IN' and id_dok='" & ain & "'", myConection, adOpenDynamic, adLockOptimistic
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
Dim sifrar As String
rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
 'sifrar = Getnazi("SELECT MADASIFR FROM MADA WHERE DOBAVIT_ID like '%" & RSS.Fields("sifr") & "%'")
 
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = RSS.Fields("sifra")
rsta.Fields("naziv") = RSS.Fields("naziv")
rsta.Fields("cena") = RSS.Fields("cena")
rsta.Fields("kol") = RSS.Fields("kol") - RSS.Fields("x")
rsta.Fields("znes") = (RSS.Fields("kol") - RSS.Fields("x")) * RSS.Fields("cena")
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("doza") = 1
rsta.Fields("stdok") = "IN" & RSS.Fields("id_dok")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
Loop
cmdAdd_Click
End Sub
Private Sub mnuuvoxls_Click()
Excelimp.odpri Me.dok.Caption
End Sub
Private Sub mnuuvoxls1_Click()
Dim RSS As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
RSS.Open "select * from xlse", myConection, adOpenDynamic, adLockOptimistic
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
Dim sifrar As String
rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
 sifrar = Getnazi("SELECT MADASIFR FROM MADA WHERE DOBAVIT_ID like '%" & RSS.Fields("sifr") & "%'")
 
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = sifrar
rsta.Fields("naziv") = RSS.Fields("nazi")
rsta.Fields("cena") = RSS.Fields("cen")
rsta.Fields("kol") = RSS.Fields("zalo")
rsta.Fields("znes") = RSS.Fields("zalo") * RSS.Fields("cen")
rsta.Fields("datum") = Me.DTPicker1.Value
rsta.Fields("doza") = 1
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
Loop
End Sub
Public Sub mviski()
Dim RSS As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
RSS.Open "select * from nabasif where tip_dok='IN' and id_dok='" & id_inv & "' and kol-x>0", myConection, adOpenDynamic, adLockOptimistic
If Not RSS.EOF Then
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
Me.UserControl11(2).BoundDatax = "INVENTURA " & id_inv
Me.DTPicker1.Value = RSS.Fields("datum")
rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = RSS.Fields("sifra")
rsta.Fields("naziv") = RSS.Fields("naziv")
rsta.Fields("cena") = RSS.Fields("cena")
rsta.Fields("kol") = RSS.Fields("kol") - RSS.Fields("x")
rsta.Fields("znes") = (RSS.Fields("kol") - RSS.Fields("x")) * RSS.Fields("cena")
rsta.Fields("datum") = Me.DTPicker1.Value
'rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
Loop
Me.Text1.Text = "VIŠKI"
Zapis_Click
End If
End Sub
Public Sub mmanjki()
Dim RSS As New ADODB.Recordset
Dim rsta As New ADODB.Recordset
If RSS.State = 1 Then RSS.Close
RSS.Open "select * from nabasif where tip_dok='IN' and id_dok='" & id_inv & "' and kol-x<0", myConection, adOpenDynamic, adLockOptimistic
If Not RSS.EOF Then
RSS.MoveFirst
Dim tii, idd As String
tii = Left(Me.dok.Caption, 2)
idd = Mid(Me.dok.Caption, 3)
Me.UserControl11(2).BoundDatax = "INVENTURA " & id_inv
Me.DTPicker1.Value = RSS.Fields("datum")
rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
Do While Not RSS.EOF
'If rsta.State = 1 Then rsta.Close
rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = RSS.Fields("sifra")
rsta.Fields("naziv") = RSS.Fields("naziv")
rsta.Fields("cena") = RSS.Fields("cena")
rsta.Fields("kol") = RSS.Fields("kol") - RSS.Fields("x")
rsta.Fields("znes") = (RSS.Fields("kol") - RSS.Fields("x")) * RSS.Fields("cena")
rsta.Fields("datum") = Me.DTPicker1.Value
'rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))
RSS.MoveNext
Loop
Me.Text1.Text = "MANJKI"
Zapis_Click
End If
End Sub
Private Sub Zapis_Click()
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
myConection.Execute ("update trenutna set stdok='" & Me.Text3.Text & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
'myConection.Execute ("update trenutna set y='" & Me.Text4.text & "' where  tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
End If
myConection.Execute ("delete  from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete  from trenutna where (sifra)='' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete  from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
If rs.State = 1 Then rs.Close
rs.Open "select * from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' order by pozicija", myConection, adOpenDynamic, adLockOptimistic
Dim aa As Integer
aa = 1
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
rs.Fields("pozicija") = levi_pres(aa, 4)
aa = aa + 1
rs.MoveNext
Loop
End If
'myConection.Execute ("delete  from trenutna where ltrim(sifra)='' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

'CENIK
If Left(Me.dok.Caption, 2) = "NA" Then
If obstaja("cenik") Then
myConection.Execute ("delete from cenik where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 myConection.Execute ("insert into cenik select datum,sifra,cena,tip_dok,id_dok from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
Else

 myConection.Execute ("select datum,sifra,cena,tip_dok,id_dok into cenik from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

 End If
 End If
Dim upor As String

upor = Getnazi("select up from users where username1='" & UPORABNIK & "'")
myConection.Execute ("insert into nabasif select * from trenutna where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
 SQL = "insert into glavna (tip_dok,id_dok,opis,dod0,dod1,dod2,dod3,dod4,dod5,dod6,dod7,skl) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','" & Me.Text1.Text & "','" & Me.UserControl11(0).BoundDatax & "','" & Me.UserControl11(1).BoundDatax & "','" & Me.UserControl11(2).BoundDatax & "','" & Me.UserControl11(3).BoundDatax & "','" & Me.UserControl11(4).BoundDatax & "','" & Me.UserControl11(5).BoundDatax & "','" & Me.UserControl11(6).BoundDatax & "','" & Me.UserControl11(7).BoundDatax & "','" & Me.sklad.BoundDatax & "')"
 ' MsgBox SQL
    myConection.Execute SQL
'datumi
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='DUR'")
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='DNI'")
myConection.Execute ("delete from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and atribut='VAL'")
If rs.State = 1 Then rs.Close
rs.Open "Select * from dokm where atribut='DUR'", myConection, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("tip_dok") = Left(Me.dok.Caption, 2)
rs.Fields("id_dok") = Mid(Me.dok.Caption, 3)
rs.Fields("atribut") = "DUR"
rs.Fields("tekst") = Me.DTPicker2.Value
rs.Update
rs.AddNew
rs.Fields("tip_dok") = Left(Me.dok.Caption, 2)
rs.Fields("id_dok") = Mid(Me.dok.Caption, 3)
rs.Fields("atribut") = "DNI"
rs.Fields("tekst") = Me.Text5.Text
rs.Update
rs.AddNew
rs.Fields("tip_dok") = Left(Me.dok.Caption, 2)
rs.Fields("id_dok") = Mid(Me.dok.Caption, 3)
rs.Fields("atribut") = "VAL"
rs.Fields("tekst") = Me.DTPicker3.Value
rs.Update
If Left(Getnazi("select tvorba from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'"), 1) = "=" Then
myConection.Execute ("delete from nabasif where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete from glavna where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("delete from dokm where tip_dok='FA' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into nabasif select 'FA' as tip_dok,DATUM, STDOK, SIFRAPART, SIFRA, EMBALAZA, KOL, CENA, ZNES, pop,  id_dok, poknj, faktor, naziv, SIFRAPLAC, mpc, x, y, uporabnik, pozicija, chk_fix, skl, kopija, dat_k, PLACILO, DOZA, org from nabasif where sifra<>'' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into glavna select 'FA' as tip_dok,id_dok, faktor, dod0, dod1, dod2, dod3, dod4, dod5, dod6, dod7, opis, skl from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
myConection.Execute ("insert into dokm select 'FA' as tip_dok,atribut, id_dok, tekst, poz from dokm where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")

End If
myConection.Execute ("update dokumenti set dol_ce=" & Val(Me.Text4.Text) & " where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
myConection.Execute ("delete  from dokm where id_dok='" & Mid(Me.dok.Caption, 3) & "' and poz=999")
 Dim xpozic As String
 xpozic = ""
 xpozic = Getnazi("SELECT dokm.atribut  From dokm GROUP BY dokm.atribut, dokm.id_dok, dokm.poz, dokm.tip_dok Having ((dokm.id_dok) = '" & Mid(Me.dok.Caption, 3) & "') And ((dokm.tip_dok) = '" & Left(Me.dok.Caption, 2) & "') And ((Count(dokm.atribut)) > 1) ORDER BY dokm.atribut")
If xpozic <> "" Then
myConection.Execute ("delete from dokm where atribut='" & xpozic & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
End If
    
    Unload Me
End Sub

Private Sub opiss_Click()
xopis = "opis"
  xid_dok = Trim(dok.Caption)
  Dialog.Show

End Sub

Private Sub Timer1_Timer()
'Exit Sub

If fgtrial.Col = Val(Me.Text4.Text) Then
       cmdAdd_Click
      fgtrial.Col = coollsi
      
      End If
If fgtrial.Col = coollem Then
fgTrial_LeaveCell
 'fgtrial.Col = fgtrial.Col + 1
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
If Getnazi("select tekst from dokm where atribut='" & levi_pres(LTrim(str(i)), 4) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "' and tip_dok='" & Left(Me.dok.Caption, 2) & "'") <> "" Then

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
'     fgtrial.SetFocus
End If
Timer1.Enabled = False
End Sub

Private Sub txtNewData_gotfocus()
tresi = fgtrial.Text
If clk = True Then
   txtNewData.Text = fgtrial.Text
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
 Timer1.Enabled = True
End Sub

Private Sub txtNewData_LostFocus()
'If Getnazi("select postava from mada where madasifr='" & Trim(txtNewData.text) & "'") <> "" Then
'       zai = fgtrial.Row
'       Postava.Show
'End If
 If curCol < edCol Or curCol = Val(Me.Text4.Text) Then
        fgtrial.Row = curRow
        fgtrial.Col = curCol + 1
        fgTrial_KeyPress (0)
    Else
    
      fgTrial_LeaveCell
    End If
    Timer1.Enabled = True
       End Sub
Private Sub txtNewData_KeyPress(KeyAscii As Integer)
'MsgBox ("3")
If KeyAscii = 27 Then
fgtrial.Text = tresi
txtNewData.Visible = False
End If
If KeyAscii = 13 Then
 fgtrial.Text = txtNewData.Text
    If fgtrial.Col = coollsi Then
       If rs.State = 1 Then rs.Close
       Dim ax As String
       ax = ""
       ax = (Getnazi("select madasifr from mada where madasifr='" & txtNewData.Text & "'"))
       If ax = "" Then
       Dim novas, vi, dol As String
       vi = ""
       dol = ""
       novas = "/" & Trim(txtNewData.Text) & "/"
       ax = (Getnazi("select madasifr from mada where dobavit_id like '%" & novas & "%'"))
       End If
      If ax = "" Then
       Dim novasx, vix, dolx As String
       vix = ""
       dolx = ""
       novasx = LTrim(RTrim(txtNewData.Text))
       If Right(novasx, 1) = "J" Then
       novasx = Left(novasx, Len(novasx) - 1)
       End If
      
       ax = (Getnazi("select madasifr from mada where madaean ='" & novasx & "'"))
       
       End If
      
       If ax = "" Then
       idar = ""
       iskalni = fgtrial.Text
       pritisk = txtNewData.Text
      ' DoSQL = ""
      If tip_dok = "DN" Then
       ax = DoSQL("mada where tip_art='IZD'", "madasifr", "madanazi", "madaenme")
       Else
       ax = DoSQL("mada where tip_art<>'IZD'", "madasifr", "madanazi", "madaenme")
    End If
       'MsgBox ax
       End If
       txtNewData.Text = Trim((ax))
       fgtrial.Text = Trim(ax)
       sifrt = (ax)
    If sifrt = "" Then
    Else
       'StatusBar1.Panels.Remove 1
        
        'StatusBar1.Panels.Add 1, , "Artikel " & ax
        rs.Open "select MADANAZI,MADAnabc,madampcd,madapd,postava,madaenme,madazalo from MADA where MADASIFR='" & ax & "'", myConection, adOpenStatic, adLockOptimistic
          If Not rs.EOF Then
              fgtrial.TextMatrix(fgtrial.Row, coollna) = Trim(rs!MADANAZI) & " "
              fgtrial.TextMatrix(fgtrial.Row, coollem) = Trim(rs!MADAenme) & " "
              If Left(Me.dok.Caption, 2) = "IN" Then
              fgtrial.TextMatrix(fgtrial.Row, coollzal) = rs!madazalo
              End If
              If Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'") <= 0 Then
              fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(rs!MADAMPCD, 4)
              Else
               fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(rs!madanabc, 4)
               End If
          End If
          Call txtNewData_LostFocus
       fgtrial.Col = 3
       If bremepis = 1 Then
       fgtrial.Col = 1
       prosti.Show vbModal
       End If
          
          
    Dim zalog, prost, sks, skp As Double
    Dim das, dodx
    Dim nazb As String
das = Format(Me.DTPicker1.Value, "dd.mm.yyyy")
dodx = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
     nazb = Trim(rs.Fields("madanazi"))
    If ax <> "" Then
     Me.StatusBar1.Panels(1).Text = Trim(rs!MADANAZI) & " " & rs!MADAenme
     If Getnazi("select sum(kol) as ss from zaloga where sifra='" & ax & "'") <> "" Then
     sks = Getnazi("select sum(kol) as ss from zaloga where sifra='" & ax & "'")
     End If
     If Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and  sifra='" & ax & "'  and datum<=#" & dodx & "#") <> "" Then
     skp = Getnazi("select sum(kol*faktor) as ss from nabasif where faktor<>0 and poknj='K' and  sifra='" & ax & "'  and datum<=#" & dodx & "#")
     End If
     Me.StatusBar1.Panels(3).Text = FormatNumber(sks, 3)
     Me.StatusBar1.Panels(5).Text = FormatNumber(skp, 3)
        'End If
      If Left(Me.dok.Caption, 2) = "IZ" Then
      fgtrial.TextMatrix(fgtrial.Row, coollpro) = Me.StatusBar1.Panels(5).Text
      End If
   End If
    End If
    
      ElseIf fgtrial.Col = coollko Then
       If IsNumber(txtNewData.Text) Then
       Else
       txtNewData.Text = 0
       End If
      
       
      'fgtrial.Rows = fgtrial.Rows + 1
      'If tip_dok <> "NA" Then
      'cmdAdd_Click
      'fgtrial.Col = coollsi
      'End If
      
      Else
     
     If fgtrial.Col = coollles Then
     
idar = ""
       iskalni = fgtrial.Text
       pritisk = txtNewData.Text
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
 If fgtrial.Col = coollmarz Then
 fgtrial.TextMatrix(fgtrial.Row, coollmpc) = (1 + (FormatNumber(txtNewData.Text, 2) / 100)) * fgtrial.TextMatrix(fgtrial.Row, coollce)
 fgtrial.TextMatrix(fgtrial.Row, coollmarz) = FormatNumber(txtNewData.Text, 2)
 End If
 If fgtrial.Col = coollstek Then
idar = ""
       iskalni = fgtrial.Text
       pritisk = txtNewData.Text
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
      fgtrial.TextMatrix(fgtrial.Row, coollce) = FormatNumber(FormatNumber(txtNewData.Text, 4) / FormatNumber(fgtrial.TextMatrix(fgtrial.Row, coollko), 4), 7)
      
      If tip_dok = "NA" Then
  '    cmdAdd_Click
  '    fgtrial.Col = coollsi
      End If
      Else
      If fgtrial.Col = coollpop Then
       If IsNumber(txtNewData.Text) Then
      txtNewData.Text = FormatNumber(txtNewData.Text, 2)
      Else
      txtNewData.Text = 0
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
         If txtNewData.Text = "" Then
         txtNewData.Text = 0
         End If
         If IsNumber(txtNewData.Text) Then
          txtNewData.Text = FormatNumber(txtNewData.Text, 4)
      If tip_dok <> "NA" Then
'           cmdAdd_Click
'          fgtrial.Col = coollsi
    End If
          Else
          txtNewData.Text = 0
         End If
 '
        
      End If
      End If

If fgtrial.Col = Val(Me.Text4.Text) Then
 cmdAdd_Click
 fgtrial.Col = coollsi
 End If
fgtrial.SetFocus
End If
End Sub

Private Sub napolni()
Dim i, stot, fa
 If rs.State = 1 Then rs.Close
 If kosovni = 1 Then
 Else
  

 rs.Open "select * from glavna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
Dim c As Integer
If Not rs.EOF Then
If Not rs.Fields("skl") = "" Then
Me.sklad.BoundDatax = rs.Fields("skl")
End If
If Not rs.Fields("opis") = "" Then
Me.Text1.Text = rs.Fields("opis")
End If

End If
If Getnazi("select datum from nabasif  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'") <> "" Then
Me.DTPicker1.Value = Getnazi("select datum from nabasif  where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
Else
Me.DTPicker1.Value = Date
End If
For c = 0 To 7
If Not rs.EOF Then
If Not rs.Fields(c + 3) = "" Then
Me.UserControl11(c).BoundDatax = rs.Fields(c + 3)
End If
End If
Next
End If
'MsgBox (aaa)
   If rs.State = 1 Then rs.Close
   If kosovni = 1 Then
 Dim Rsa As New ADODB.Recordset
  
  Rsa.Open "select * from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
   rs.Open "select * from normati ", myConection, adOpenStatic, adLockOptimistic
   If Not Rsa.EOF Then
   myConection.Execute ("delete from trenutna where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'")
   End If
   If Not rs.EOF Then
   rs.MoveFirst
   Dim aa As Integer
   aa = 1
   Do While Not rs.EOF
    Rsa.AddNew
    Rsa.Fields("tip_dok") = Left(Me.dok.Caption, 2)
    Rsa.Fields("id_dok") = Mid(Me.dok.Caption, 3)
    Rsa.Fields("pozicija") = levi_pres(LTrim(str(aa)), 4)
    Rsa.Fields("sifra") = rs.Fields("sifr")
    Rsa.Fields("kol") = rs.Fields("kol")
    Rsa.Fields("naziv") = rs.Fields("naz")
    Rsa.Update
    
    aa = aa + 1
    
    rs.MoveNext
    
   Loop
'   Call cmdAdd_Click
kosovni = 0
   refre
   
   End If
   Else
 rs.Open "select * from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'", myConection, adOpenStatic, adLockOptimistic
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



