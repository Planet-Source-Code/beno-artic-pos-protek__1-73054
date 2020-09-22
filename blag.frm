VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form blag 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BLAGAJNA  "
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   10365
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   4575
      Left            =   1440
      TabIndex        =   42
      Top             =   4080
      Visible         =   0   'False
      Width           =   9735
      Begin LVbuttons.LaVolpeButton LaVolpeButton2532 
         Height          =   495
         Left            =   7800
         TabIndex        =   43
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "POTRDI"
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
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "blag.frx":0000
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
      Begin MSForms.ListBox ListBox1 
         Height          =   3615
         Left            =   0
         TabIndex        =   44
         Top             =   240
         Width           =   9375
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "16536;6376"
         MatchEntry      =   0
         MultiSelect     =   1
         FontName        =   "Courier New"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   238
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin LVbuttons.LaVolpeButton opiss 
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   2
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
      MICON           =   "blag.frx":001C
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   25
      Top             =   3360
      Width           =   7335
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      Top             =   960
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   7440
      TabIndex        =   22
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      Format          =   67239937
      CurrentDate     =   39472
   End
   Begin VB.TextBox cmbItmcode 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   3
      Left            =   3600
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox txtEnter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   3
      Index           =   5
      Left            =   10200
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox txtEnter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   3
      Index           =   4
      Left            =   9480
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox txtEnter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   3
      Index           =   3
      Left            =   6720
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox nazivv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   4320
      TabIndex        =   17
      Top             =   0
      Width           =   5955
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   840
      Top             =   9240
   End
   Begin VB.TextBox pop 
      Height          =   465
      Left            =   10440
      TabIndex        =   13
      Text            =   "0"
      Top             =   9120
      Width           =   855
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2522 
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   9720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Delni - Briši - F7"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "blag.frx":0038
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   2
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   5520
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2460
      Top             =   5760
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1200
      Left            =   6480
      TabIndex        =   11
      Top             =   7920
      Width           =   7935
   End
   Begin VB.ComboBox tip_c 
      BackColor       =   &H00D1DECD&
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
      Left            =   2880
      TabIndex        =   2
      Top             =   4725
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtEnter 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   2
      Index           =   2
      Left            =   1140
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1665
   End
   Begin MSFlexGridLib.MSFlexGrid MsfBill 
      Height          =   3840
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   6773
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13754061
      BackColorBkg    =   14807790
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton46 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   9720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "ZAPIŠI-F4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "blag.frx":0054
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   1
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton45 
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   9720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "STORNO - F3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "blag.frx":0070
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   2
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton44 
      Height          =   495
      Left            =   11400
      TabIndex        =   6
      Top             =   9720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "IZHOD - F10"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "blag.frx":008C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   3
   End
   Begin LVbuttons.LaVolpeButton pred 
      Height          =   495
      Left            =   6840
      TabIndex        =   15
      Top             =   9120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "PREDAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "blag.frx":00A8
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
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   27
      Top             =   1440
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   28
      Top             =   1920
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   29
      Top             =   2400
      Width           =   4815
      _extentx        =   6800
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   4
      Left            =   8640
      TabIndex        =   30
      Top             =   960
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   5
      Left            =   8640
      TabIndex        =   31
      Top             =   1440
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   6
      Left            =   8640
      TabIndex        =   32
      Top             =   1920
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   7
      Left            =   8640
      TabIndex        =   33
      Top             =   2400
      Width           =   4815
      _extentx        =   8493
      _extenty        =   661
      ssql            =   "select * from partner"
      polje           =   "naziv"
      textlocked      =   0   'False
      locked          =   0   'False
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
      Left            =   6960
      TabIndex        =   41
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
      Left            =   6960
      TabIndex        =   40
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
      Left            =   6960
      TabIndex        =   39
      Top             =   1560
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
      Left            =   6960
      TabIndex        =   38
      Top             =   1080
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
      Left            =   120
      TabIndex        =   37
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
      Left            =   120
      TabIndex        =   36
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
      Left            =   120
      TabIndex        =   35
      Top             =   1560
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
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   1080
      Width           =   1455
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
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   2055
   End
   Begin MSForms.CheckBox inter 
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   9240
      Visible         =   0   'False
      Width           =   2535
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "4471;661"
      Value           =   "0"
      Caption         =   "INTERNA"
      FontEffects     =   1073741825
      FontHeight      =   270
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "POPUST"
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
      Left            =   9120
      TabIndex        =   14
      Top             =   9240
      Width           =   1095
   End
   Begin MSForms.CheckBox kart 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   9240
      Width           =   2535
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "4471;661"
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lbst 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label stranka 
      BackColor       =   &H00FFC0C0&
      Caption         =   " "
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   11280
      TabIndex        =   7
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label LblDateTime 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "blag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gSlno, gItemCode, gItemname, gQty, gRate, gTotal, Inti, miz, i, gpop, gce
Dim Indx
Public ahha As Long

Private Sub cena_Change()
MsfBill.text = cena.text
End Sub
Private Sub cena_gotfocus()
cena.text = MsfBill.text
End Sub

Private Sub cmbItmcode_LostFocus()

'MsgBox ("0")
If fora = 0 Then
If deln = 1 Then
Else
SendKeys 1
kolik = 1
End If
'MsfBill.TextMatrix(Indx, 4) = 1
Else
fora = 0
End If
End Sub
Private Sub cmbItmcode_GotFocus()
'MsgBox ("1")
cmbItmcode.text = MsfBill.text
'SendKeys "{BS 6}"
If ahha <> 0 Then
SendKeys (ahha)
SendKeys "{enter}"
ahha = 0
End If
If fora = 9 Then
fora = 0
SendKeys "{BS}", 1

End If
End Sub
Private Sub cmbItmcode_Change()
'MsgBox ("2")
MsfBill.text = cmbItmcode.text

End Sub

Private Sub cmbItmcode_KeyPress(KeyAscii As Integer)
'MsgBox ("3")
If KeyAscii = 13 And cmbItmcode.text <> "" Then

   cmbItmcode.Visible = False
     MsfBill.TextMatrix(Indx, 1) = cmbItmcode.text
     
   If RS.State = 1 Then RS.Close
   If Len(cmbItmcode.text) > 12 Then
   RS.Open "select MADANAZI,MADAMPCD from MADA where MADAean='" & cmbItmcode.text & "'", myConection, adOpenStatic, adLockOptimistic
   Else
   Dim ax As String
 '  ax = Getnazi("select madanazi from mada where madasifr=" & cmbItmcode.text)
    RS.Open "select MADANAZI,MADAMPCD,madapd from MADA where MADASIFR=" & cmbItmcode.text, myConection, adOpenStatic, adLockOptimistic
   End If
      If Not RS.EOF Then
         MsfBill.TextMatrix(Indx, 2) = RS!MADANAZI & ""
          MsfBill.TextMatrix(Indx, 5) = Me.pop.text
         
         'MsfBill.TextMatrix(Indx, 3) = RS!MADAMPCD / (1 + (RS!madapd / 100))
         'MsfBill.TextMatrix(Indx, 6) = RS!MADAMPCD
         'MsfBill.TextMatrix(Indx, 7) = RS!MADAMPCD
         'Else
         MsfBill.TextMatrix(Indx, 3) = Round(RS!MADAMPCD / (1 + (RS!madapd / 100)), 2)
         MsfBill.TextMatrix(Indx, 6) = Round((RS!MADAMPCD) / (1 + (MsfBill.TextMatrix(Indx, 5) / 100)), 2)
         MsfBill.TextMatrix(Indx, 7) = Round((RS!MADAMPCD) / (1 + (MsfBill.TextMatrix(Indx, 5) / 100)), 2)
         'End If
         MsfBill.Col = 4
         ArrangeTextbox txtEnter(MsfBill.Col)
      Else
         MsgBox "Ta šifra ne obstaja preveri prosim! ", vbCritical
         MsfBill.Col = 1
         ArrangeTextbox cmbItmcode
      End If
End If

End Sub
Private Sub cmbItmcode_KeyUp(KeyCode As Integer, Shift As Integer)
If zap <> 0 Then

MsfBill.Row = zap
'MsfBill.TextMatrix(ZAP, 0) = ZAP

Indx = zap
zap = 0
End If

If xxre <> "" Then

Me.cmbItmcode = xxre
SendKeys "{enter}"
xxre = ""
End If
End Sub
Private Sub cmbItmcode_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")

Select Case KeyCode
Case vbKeyF3
 LaVolpeButton45.SetFocus
 SendKeys "{enter}"
 Case vbKeyF7
 LaVolpeButton2522.SetFocus
  SendKeys "{enter}"
 Case vbKeyF2
 If Me.kart.Value = True Then
 Me.kart.Value = False
 Else
 Me.kart.Value = True
 End If
  Case vbKeyF9
'VRNIT.SetFocus
'   SendKeys "{enter}"
  Case vbKeyF10
  LaVolpeButton44.SetFocus
   SendKeys "{enter}"
 Case vbKeyF8
 'LaVo2.SetFocus
 ' SendKeys "{enter}"
 Case vbKeyF6
fora = 1
' Me.mii.Visible = True
' Me.mii.text = ""
 
' Me.mii.SetFocus
 
 
Case vbKeyF4
 LaVolpeButton46.SetFocus
 SendKeys "{enter}"
 Case vbKeyA To vbKeyZ
Dim idar As String
zap = Indx
idar = Chr(KeyCode)
   DoSQL "mada", "madasifr", "madanazi", "madanaz1"
       
       
Case Else
    End Select
End Sub



Private Sub Command2_Click()
MsfBill.Col = 4
MsfBill.SetFocus

End Sub

Private Sub Command1_Click()
MsgBox (OSEB)
End Sub

Private Sub desnog_Click()
If Val(Me.Label8.Caption) - 24 > 0 Then
Me.Label8.Caption = Str(Val(Me.Label8.Caption) - 24)
Else
Me.Label8.Caption = "0"
End If
Dim q As Integer
q = Val(Me.Label9.Caption)
Hanb (q)
End Sub

Private Sub Form_Activate()
FillCombo tip_c, "select skladisce from skla"
tip_c.text = Getnazi("select min(skladisce) as skl from skla")
MsfBill.SetFocus
MsfBill.Row = 0
MsfBill.Col = 3
ArrangeTextbox tip_c

If zap <> 0 Then
MsfBill.Row = zap
Else
MsfBill.Row = 1
End If
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx

'txtInvoiceNo.text = GetNewNo("select max(st)+1 from racusif")
nazivv.text = Getnazi("select glava1 from oblikar")
Me.dok.Caption = (tip_dok) & novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='" & tip_dok & "'")) + 1, 6)
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
If ma_ured = 1 Then
napolni
End If
End Sub
Private Sub napolni()
Dim i, stot, fa
 If RS.State = 1 Then RS.Close
   
 RS.Open "select * from glavna where tip_dok='" & frmControlMain.DataGrid1.Columns(0).text & "' and id_dok='" & frmControlMain.DataGrid1.Columns(1).text & "'", myConection, adOpenStatic, adLockOptimistic
Dim c As Integer
For c = 0 To 7
If Not RS.EOF Then
Me.UserControl11(c).BoundDatax = RS.Fields(c + 3)
End If
Next
'MsgBox (aaa)
   If RS.State = 1 Then RS.Close
   
 RS.Open "select tip_dok,id_dok,datum,sifra,kol,cena,mpc,znes,faktor,naziv,pop from nabasif where tip_dok='" & frmControlMain.DataGrid1.Columns(0).text & "' and id_dok='" & frmControlMain.DataGrid1.Columns(1).text & "'", myConection, adOpenStatic, adLockOptimistic

Dim po As Integer
Dim kol As Integer
Dim znes As Double
po = 1
Me.DTPicker1.Value = RS.Fields(2)
Do While Not RS.EOF
If RS.EOF Then
Else

       
       MsfBill.TextMatrix(po, 0) = po
       MsfBill.TextMatrix(po, 1) = RS.Fields(3)
       MsfBill.TextMatrix(po, 2) = RS.Fields(9)
       MsfBill.TextMatrix(po, 3) = RS.Fields(5)
       MsfBill.TextMatrix(po, 4) = RS.Fields(4)
       MsfBill.TextMatrix(po, 5) = RS.Fields(10)
       If Not IsNull(RS.Fields(6)) Then
       MsfBill.TextMatrix(po, 6) = RS.Fields(6)
       End If
       MsfBill.TextMatrix(po, 7) = RS.Fields(7)
      znes = znes + RS.Fields(7)
MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
        '  txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
           FlexgridTotal
po = po + 1
RS.MoveNext
End If
 Loop
 txtTotal.text = Format(znes, "fixed")
 skumi = znes
 zap = Indx
    ind = po
MsfBill.SetFocus
'ArrangeTextbox cmbItmcode
Indx = ind
zap = Indx
 MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
         ' txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
ind = 0

End Sub
Private Sub Form_Load()
ReSizeForm Me

MsfRefresh
'FillCombo cmbItmcode, "select MADASIFR from MADA"
 
End Sub
Private Sub MsfRefresh()
Dim sngVertFactor As Single
    sngVertFactor = getFactor(True)
With MsfBill
      .Cols = 7
      .Rows = 2
      .FormatString = "^No | SIFRA | NAZIV |  PC   | KOL  | POP  | MPCD  | ZNESEK "
       gSlno = 0
       gItemCode = 1
       gItemname = 2
       gQty = 3
       gRate = 4
       gpop = 5
       gce = 6
       gTotal = 7
       .Row = 0
       For Inti = 0 To .Cols - 1
          .Col = Inti
          .CellFontBold = True
       Next
       .ColWidth(gSlno) = 3 * 100 * sngVertFactor
       .ColWidth(gItemCode) = 15 * 100 * sngVertFactor
       .ColWidth(gItemname) = 38 * 100 * sngVertFactor
       .ColWidth(gRate) = 8 * 100 * sngVertFactor
       .ColWidth(gQty) = 15 * 100 * sngVertFactor
       .ColWidth(gpop) = 8 * 100 * sngVertFactor
       .ColWidth(gce) = 15 * 100 * sngVertFactor
       
       .ColWidth(gTotal) = 20 * 100 * sngVertFactor
       .RowHeight(0) = 350 * sngVertFactor
       .RowHeightMin = 350 * sngVertFactor
End With
End Sub

Private Sub ArrangeTextbox(ctrl As Control)
  ctrl.Left = MsfBill.Left + MsfBill.CellLeft
  ctrl.Top = MsfBill.Top + MsfBill.CellTop
  If ctrl.text <> "" Then
  ctrl.text = MsfBill.text
  Else
  ctrl.text = ctrl.text
  End If
  ctrl.Width = MsfBill.ColWidth(MsfBill.Col) - 10
  If TypeOf ctrl Is TextBox Then
  ctrl.Height = MsfBill.RowHeight(MsfBill.Row) - 10
  End If
  ctrl.Visible = True
  'ctrl.text = ""
  ctrl.SetFocus
  ctrl.SelStart = 0
  ctrl.SelLength = Len(ctrl.text)
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub ImgNew_Click()
'Clear frmsalesbill

txtInvoiceNo.text = GetNewNo("select max(invoiceNo)+1 from sales")
MsfBill.SetFocus
MsfBill.Row = 1
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
End Sub

Private Sub ImgSave_Click()
Dim i
Dim TrxType
TrxType = "S"
If MsgBox("Do you want to Save Bill", vbQuestion + vbYesNo + vbDefaultButton1, "Additional security") = vbYes Then
    For i = 1 To MsfBill.Row
     If Len(Trim(MsfBill.TextMatrix(i, 1))) = 0 Then
           MsgBox "Item Code. is Empty Please Enter"
           MsfBill.Row = i
           MsfBill.Col = 1
           Exit Sub
        End If
        If Len(Trim(MsfBill.TextMatrix(i, 4))) = 0 Then
           MsgBox "Qty. is Empty Please Enter"
           MsfBill.Row = i
           MsfBill.Col = 4
           Exit Sub
        End If
        If Len(Trim(MsfBill.TextMatrix(i, 3))) = 0 Then
           MsgBox "Rate is Empty Please Enter"
           MsfBill.Row = i
           MsfBill.Col = 3
           Exit Sub
        End If
        If Val(MsfBill.TextMatrix(i, 3)) = 0 Then
           MsgBox "Cheque Amount is Empty Please Enter"
           MsfBill.Row = i
           MsfBill.Col = 3
           Exit Sub
        End If
    Next
    For i = 1 To MsfBill.Row
        Update1 "Stock", MsfBill.TextMatrix(i, 1), MsfBill.TextMatrix(i, 4) * -1, TrxType, MsfBill.TextMatrix(i, 3)
    Next
    MsgBox "New Bill  details sucessfully Updated", vbInformation
End If
End Sub

Private Sub karto_Click()
C_frmCategory.Show
End Sub

Private Sub LaVo2_Click()
Dim iid As String
fora = 1
jestran = 1
opp = Me.cmbItmcode.Top
oppa = Me.cmbItmcode.Left
'idar = Chr(KeyCode)
ind = Indx
idar = ""
   DoSQL "partner", "sifra", "naziv", ""


End Sub

Private Sub LaVolpeButton1_click()
xopis = "opis"
  xid_dok = Trim(dok.Caption)
  Dialog.Show
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
 If RS.State = 1 Then RS.Close
   RS.Open "select * from swit WHERE [ItemNumber] > 0 AND [Switchboar]=1 order by itemnumber"
      RS.MoveFirst
      Dim aad As Integer
      aad = 0
      If Not RS.EOF Then

       While (Not (RS.EOF))
       aad = aad + 1
         Me("nas" & aad).Caption = RS![ITEMTEXT]
         Me("nas" & aad).Tag = RS![SWITCHBOAR]
            RS.MoveNext
        Wend
      Else
         End If



End Sub

Private Sub LaVolpeButton251_Click()
OSE = Me.Label3.Caption
Form3.Show

End Sub

Private Sub LaVolpeButton2522_Click()
Me.Frame2.Visible = True
Dim i
With ListBox1
For i = 1 To MsfBill.Row
.AddItem presled(MsfBill.TextMatrix(i, 1), 13) & "  " & presled(MsfBill.TextMatrix(i, 2), 17) & "      " & MsfBill.TextMatrix(i, 4)
 Next
End With
Me.ListBox1.SetFocus

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
   If RS.State = 1 Then RS.Close
   
 
RS.Open "select sifra,kol,znesek,datum,ura,stmize from mize", myConection


  
    For intCurrentRow = 0 To ListBox1.ListCount - 1
       
            
    a = Val(Left(ListBox1.Column(0, intCurrentRow), 13))
    b = Val(Right(ListBox1.Column(0, intCurrentRow), 6))
    If ListBox1.Selected(intCurrentRow) Then
    SendKeys a & "{enter}{BS}" & b & "{enter}"
        '
        '  MsfBill.TextMatrix(Indx, 0) = Indx
          
                 'MsfBill.TextMatrix(Indx, 0) = Indx
'MsfBill.TextMatrix(Indx, 1) = Left(ListBox1.Column(0, intCurrentRow), 13)
'MsfBill.TextMatrix(Indx, 2) = Getnazi("select madanazi from mada where madasifr=" & Left(ListBox1.Column(0, intCurrentRow), 13))
'MsfBill.TextMatrix(Indx, 4) = Right(ListBox1.Column(0, intCurrentRow), 6)
'Indx = Indx + 1
'po = po + 1
'MsfBill.Row = po
Else
If stm1 <> 0 Then
If a <> 0 Then
Dim cen As Double
cen = Getnazi("select madampcd from mada where madasifr=" & a)
RS.AddNew
    RS.Fields(0) = a
    RS.Fields(1) = b
    RS.Fields(2) = b * cen 'Val(MsfBill.TextMatrix(i, 5))
    RS.Fields(3) = Date
    RS.Fields(4) = aaa
      RS.Fields(5) = stm1
      RS.Update
End If
End If
    
 
        End If
      
       ' zap = Indx
'          fora = 2
    Next intCurrentRow
RS.Close
'       fora = 2
Me.ListBox1.clear
refr = 1
stm1 = 0
    Me.cmbItmcode.text = ""

Me.Frame2.Visible = False
deln = 0
End Sub

Private Sub LaVolpeButton3_Click()
Hanbt (3)
End Sub

Private Sub LaVolpeButton4_Click()
Hanbt (4)
End Sub

Private Sub LaVolpeButton44_Click()
'End
blagajna = 0
Close
'Form8.Show

Unload Me
End Sub

Private Sub LaVolpeButton45_Click()
Dim stot, fa
Indx = 1

zap = 1
Me.MsfBill.clear
MsfRefresh
MsfBill.SetFocus
If zap <> 0 Then
MsfBill.Row = zap
Else
MsfBill.Row = 1
End If
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
   stot = 0
  fa = Format(stot, "fixed")
txtTotal.text = fa
idstran = 0

Indx = 1
zap = 0

           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          If MsfBill.Col > 1 Then
          txtEnter(MsfBill.Col).Visible = False
          End If
          MsfBill.Col = 1
          ArrangeTextbox cmbItmcode
           
SendKeys "{BS}"
Me.kart.Value = False
skumi = 0
End Sub

Private Sub LaVolpeButton46_Click()
  Dim CNT1 As Integer
myConection.Execute "delete from nabasif where tip_dok='" & Left(Me.dok.Caption, 2) & "' and id_dok='" & Mid(Me.dok.Caption, 3) & "'"
Dim Y As Integer

   SQL = "insert into glavna (tip_dok,id_dok,opis,dod0,dod1,dod2,dod3,dod4,dod5,dod6,dod7) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','" & Me.Text1.text & "','" & Me.UserControl11(0).BoundDatax & "','" & Me.UserControl11(1).BoundDatax & "','" & Me.UserControl11(2).BoundDatax & "','" & Me.UserControl11(3).BoundDatax & "','" & Me.UserControl11(4).BoundDatax & "','" & Me.UserControl11(5).BoundDatax & "','" & Me.UserControl11(6).BoundDatax & "','" & Me.UserControl11(7).BoundDatax & "')"
  ' MsgBox SQL
    myConection.Execute SQL
Dim Rsa As New ADODB.Recordset
   If Rsa.State = 1 Then Rsa.Close

 
Rsa.Open "select tip_dok,id_dok,datum,sifra,kol,cena,mpc,znes,faktor,naziv,pop from nabasif", myConection, adOpenStatic, adLockOptimistic
Dim ddd As Integer

For i = 1 To MsfBill.Row
If Val(MsfBill.TextMatrix(i, 1)) <> 0 Then
Rsa.AddNew
    Rsa.Fields(0) = Left(Me.dok.Caption, 2)
    Rsa.Fields(1) = Mid(Me.dok.Caption, 3)
    Rsa.Fields(2) = Me.DTPicker1.Value
    Rsa.Fields(3) = Val(MsfBill.TextMatrix(i, 1))
    Rsa.Fields(4) = Val(MsfBill.TextMatrix(i, 4))
    Rsa.Fields(5) = Val(MsfBill.TextMatrix(i, 3))
    Rsa.Fields(6) = Val(MsfBill.TextMatrix(i, 6))
    Rsa.Fields(7) = Val(MsfBill.TextMatrix(i, 7))
    Rsa.Fields(8) = Getnazi("select faktor from dokumenti where tip_dok='" & Left(Me.dok.Caption, 2) & "'")
    Rsa.Fields(9) = MsfBill.TextMatrix(i, 2)
    Rsa.Fields(10) = Val(MsfBill.TextMatrix(i, 5))
   SQL = "update mada set madazalo=" & Val(Getnazi("select sum(kol*faktor)  from nabasif where sifra=" & Val(MsfBill.TextMatrix(i, 1)) & " and poknj='K'")) & " where madasifr=" & Val(MsfBill.TextMatrix(i, 1))
 '  MsgBox SQL
    myConection.Execute SQL
 End If
Next
 Rsa.Update
 Rsa.Close
 Set Rsa = Nothing
Call LaVolpeButton45_Click
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
Me.Label8.Caption = Str(Val(Me.Label8.Caption) + 24)
Dim q As Integer
q = Val(Me.Label9.Caption)
Hanb (q)
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

Private Sub mizaa_Click(Index As Integer)
stm1 = Index
If mizaa(Index).BackColor = 14215660 Then
shranimi (Index)
Indx = 1
zap = 0
MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
  Me.cmbItmcode.SetFocus
Else
odprimi (Index)
Dim sSQL As String
    
    'default
    
    
    sSQL = "DELETE * FROM mize WHERE stmize=" & Index
    myConection.Execute sSQL
    mizaa(Index).BackColor = 14215660
'MsfBill.SetFocus
fora = 9
Me.cmbItmcode.SetFocus

End If

End Sub

Private Sub MsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case vbKeyF3
 LaVolpeButton45.SetFocus
 SendKeys "{enter}"
Case vbKeyF4
 LaVolpeButton46.SetFocus
 SendKeys "{enter}"
Case Else
    End Select
End Sub

Private Sub MsfBill_Click()
  If MsfBill.Col = 0 Then
  xopis = MsfBill.Row
  xid_dok = Trim(dok.Caption)
  Dialog.Show
  End If
  If MsfBill.Col = 1 Then
     MsfBill.Col = 1
     ArrangeTextbox cmbItmcode
  ElseIf MsfBill.Col = 2 Then
  
     MsfBill.Col = 2
     ArrangeTextbox txtEnter(2)
  ElseIf MsfBill.Col = 3 Then
     MsfBill.Col = 3
     ArrangeTextbox txtEnter(3)
  ElseIf MsfBill.Col = 4 Then
     MsfBill.Col = 4
     ArrangeTextbox txtEnter(4)
  ElseIf MsfBill.Col = 5 Then
     MsfBill.Col = 5
     ArrangeTextbox txtEnter(5)
  End If
End Sub

Private Sub nas1_Click()
Me.Label8.Caption = "0"
Hanb (1)
Me.Label9.Caption = "1"

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

Private Sub naziv_Change()
MsfBill.text = naziv.text
End Sub
Private Sub naziv_gotfocus()
naziv.text = MsfBill.TextMatrix(MsfBill.Row, 2)
End Sub

Private Sub opiss_Click()
xopis = "opis"
  xid_dok = Trim(dok.Caption)
  Dialog.Show
End Sub

Private Sub pred_Click()

predal
End Sub

Private Sub prija_Click()
Form4.Show
End Sub

Private Sub stev1_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{enter}"
  
     Else
    Me.cmbItmcode.SetFocus
   
    End If
End Sub

Private Sub stev2_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}2"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev3_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}3"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev4_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}4"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev5_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}5"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev6_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}6"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev7_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}7"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev8_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}8"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub stev9_Click(Index As Integer)
 If Me.MsfBill.Col = 4 Then
   Me.txtEnter.SetFocus
    SendKeys "{BS}9"
   SendKeys "{enter}"
     Else
    Me.cmbItmcode.SetFocus
   
    End If

End Sub

Private Sub Timer1_Timer()
Me.Label3.Caption = OSEB
LblDateTime.Caption = Time() & " " & Format(Date, "DDDD")

'txtInvoiceNo.text = GetNewNo("select max(st)+1 from racusif")
If idstran <> 0 Then
Me.stranka.Caption = Getnazi("select naziv from partner where sifra=" & idstran)
Me.lbst.Caption = "Stranka:"
'Me.karto.Visible = True
Else
Me.stranka.Caption = ""
'Me.karto.Visible = False
Me.lbst.Caption = ""
End If
End Sub

Private Sub Timer3_Timer()
If izja = 1 Then
If Getnazi("select tekst from dokm where atribut='opis' and iddo='" & Trim(Me.dok.Caption) & "'") <> "" Then
opiss.BackColor = 255
Else
opiss.BackColor = &HE0E0E0

End If
For i = 1 To MsfBill.Rows - 1
MsfBill.Col = 0
MsfBill.Row = i
If Getnazi("select tekst from dokm where atribut='" & Trim(Str(i)) & "' and iddo='" & Trim(Me.dok.Caption) & "'") <> "" Then

MsfBill.CellBackColor = 255
Else
MsfBill.CellBackColor = &H80000005
End If
Next
izja = 0
End If
End Sub

Private Sub txtEnter_Change(Index As Integer)
 MsfBill.text = txtEnter(Index).text
End Sub
Private Sub txtEnter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case vbKey0 To vbKey9
 If kolik = 1 Then
SendKeys "{BS}"
SendKeys Chr(KeyCode)
kolik = 0
End If
 Case vbKeyA To vbKeyZ

Case Else
    End Select
End Sub

Private Sub txtEnter_gotfocus(Index As Integer)
txtEnter(Index).text = MsfBill.text
End Sub
Private Sub txtEnter_KeyPress(Index As Integer, KeyAscii As Integer)
Indx = MsfBill.Row
If KeyAscii >= 48 And KeyAscii <= 57 Then


End If

If KeyAscii = 13 Then
  If MsfBill.Col = 1 Then
     MsfBill.Col = 2
     ArrangeTextbox txtEnter(Index)
      ElseIf MsfBill.Col = 5 Then
      
      MsfBill.TextMatrix(Indx, 6) = Round(Val(MsfBill.TextMatrix(Indx, 3) / (1 + (Val(MsfBill.TextMatrix(Indx, 5)) / 100))), 2)
      'ArrangeTextbox txtEnter(6)
     MsfBill.Col = 4
     ArrangeTextbox txtEnter(Index)
  ElseIf MsfBill.Col = 2 Then
     MsfBill.Col = 3
     ArrangeTextbox txtEnter(Index)
  ElseIf MsfBill.Col = 3 Then
     MsfBill.Col = 4
     ArrangeTextbox txtEnter(Index)
  ElseIf MsfBill.Col = 4 Then
  
  If MsfBill.TextMatrix(Indx, 4) = "" Then Exit Sub
  
  
Dim asaa As Double
Dim asa As Double
 asa = Val(MsfBill.TextMatrix(Indx, 6))
      asaa = asa * Val(MsfBill.TextMatrix(Indx, 4))

      MsfBill.TextMatrix(Indx, 7) = asaa
      FlexgridTotal
      
      'If MsgBox("Do you want to add Additional Items", vbQuestion + vbYesNo + vbDefaultButton1, "Additional security") = vbYes Then
      If Indx + 1 = MsfBill.Rows Then
     
           MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
           MsfBill.TextMatrix(Indx, 0) = Indx
           txtEnter(Index).Visible = False
           
            MsfBill.TextMatrix(Indx, 1) = ""
            ArrangeTextbox cmbItmcode
      Else
      Indx = MsfBill.Rows - 1
       MsfBill.Col = 1
           MsfBill.Row = Indx
           MsfBill.TextMatrix(Indx, 0) = Indx
           txtEnter(Index).Visible = False
           
      MsfBill.TextMatrix(Indx, 1) = ""
      ArrangeTextbox cmbItmcode
      End If
      '    ImgSave_Click
 ' End If
 
End If

End If
End Sub
Private Sub FlexgridTotal()
Dim stot, fa
stot = 0
For i = 1 To MsfBill.Rows - 1
stot = stot + (Val(MsfBill.TextMatrix(i, 6)) * Val(MsfBill.TextMatrix(i, 4)))
Next
'stot = Val(txtTotal) + Val(MsfBill.TextMatrix(Indx, 5))
fa = Format(stot, "fixed")
txtTotal.text = fa
'txtTotal.Text = sTot
End Sub
Private Function CalculateTotAmount()
 Dim ToTamt
        ToTamt = 0
         For Inti = 1 To MsfBill.Rows - 1
            ToTamt = ToTamt + Val(MsfBill.TextMatrix(Inti, 3))
        Next
        CalculateTotAmount = FormatNumber(Val(ToTamt), 2)
        
End Function

Private Function Hanb(intBtn As Integer)
    trenu = intBtn
    Flistvel veli, "select dim from swit WHERE [command]<>1 AND [Switchboar]=" & Me("nas" & intBtn).Tag & " group by dim order by dim"
    
    If RS.State = 1 Then RS.Close
   If sqlb = "" Then
   RS.Open "select * from swit WHERE [ItemNumber] > " & Val(Me.Label8.Caption) + 1 & " and [command]<>1 AND [Switchboar]=" & Me("nas" & intBtn).Tag & " order by [ItemNumber]"
   Else
   RS.Open sqlb
   'sqlb = ""
   End If
      If RS.EOF Then
      Exit Function
      End If
      RS.MoveFirst
      Dim aad As Integer
      aad = 0
      If Not RS.EOF Then
 Do While Not aad = 24
      aad = aad + 1
      Me("LaVolpeButton" & aad).Tag = ""
      Me("LaVolpeButton" & aad).Visible = True
     
      Loop
      aad = 0
      RS.MoveFirst
       While Not RS.EOF
       aad = aad + 1
       If aad <= 24 Then
       If Not IsNull(RS![ITEMTEXT]) Then
         Me("LaVolpeButton" & aad).Caption = RS![ITEMTEXT]
         Me("LaVolpeButton" & aad).Tag = RS![ARGUMENT]
         End If
       End If
            RS.MoveNext
        Wend
        aad = 0
      Do While Not aad = 24
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
   If Me.MsfBill.Col = 4 Then
   ahha = Me("LaVolpeButton" & intBt).Tag
   stev1_Click (1)
  
'   Hanbtx (intBt)
   Else
    Me.cmbItmcode.SetFocus
  
    Me.cmbItmcode = Me("LaVolpeButton" & intBt).Tag
SendKeys "{enter}"

 If Indx = 1 And MsfBill.TextMatrix(Indx, 4) <> "" Then
 MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
           End If
End If
End Function

Private Function Hanbtx(intBt As Integer)
    'Me.cmbItmcode.SetFocus
  MsgBox (intBt)
    Me.cmbItmcode = Me("LaVolpeButton" & intBt).Tag
SendKeys "{enter}"

 If Indx = 1 And MsfBill.TextMatrix(Indx, 4) <> "" Then
 MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
           End If

End Function


Public Function hh()
Indx = ind
zap = Indx
 MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          'ArrangeTextbox cmbItmcode
ind = 0
'MsfBill.SetFocus
'SendKeys "{BS}"
End Function
Private Function mi()
Dim strsq As String
strsq = "select stmize from mize group by stmize order by stmize"
If RS.State = 1 Then RS.Close
RS.Open strsq, myConection
Dim ss As String
ss = ""
If Not RS.EOF Then
    RS.MoveFirst
    Do While Not RS.EOF
 ss = ss & "," & RS.Fields("stmize")
       Me.mizaa(RS.Fields("stmize")).BackColor = 5609
    RS.MoveNext
    Loop
    'MsgBox (ss)
End If
End Function
Private Function shranimi(stm As Integer)
Dim i, stot, fa
Dim aaa As String
aaa = Left(Time(), 8)
'MsgBox (aaa)
   If RS.State = 1 Then RS.Close
   
 
RS.Open "select sifra,kol,znesek,datum,ura,stmize from mize", myConection, adOpenStatic, adLockOptimistic
For i = 1 To MsfBill.Row
If Val(MsfBill.TextMatrix(i, 1)) <> 0 Then
RS.AddNew
    RS.Fields(0) = Val(MsfBill.TextMatrix(i, 1))
    RS.Fields(1) = Val(MsfBill.TextMatrix(i, 4))
    RS.Fields(2) = Val(MsfBill.TextMatrix(i, 7))
    RS.Fields(3) = Date
    RS.Fields(4) = aaa
      RS.Fields(5) = stm
      
    RS.Update
 End If
    Next
 RS.Close
Indx = 1
zap = 1
Me.MsfBill.clear
MsfRefresh
MsfBill.SetFocus
If zap <> 0 Then
MsfBill.Row = zap
Else
MsfBill.Row = 1
End If
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
  stot = 0
  fa = Format(stot, "fixed")
txtTotal.text = fa
idstran = 0
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
skumi = 0
End Function
Private Function odprimi(stm As Integer)
Dim i, stot, fa
Dim aaa As String
aaa = Left(Time(), 8)
'MsgBox (aaa)
   If RS.State = 1 Then RS.Close
   
 
RS.Open "select sifra,kol, znesek from mize where stmize=" & stm, myConection
Dim po As Integer
Dim kol As Integer
Dim znes As Double
po = 1
Do While Not RS.EOF
If RS.EOF Then
Exit Function
End If
MsfBill.TextMatrix(po, 0) = po
MsfBill.TextMatrix(po, 1) = RS.Fields(0)
MsfBill.TextMatrix(po, 2) = Getnazi("select madanazi from mada where madasifr=" & RS.Fields(0))
MsfBill.TextMatrix(po, 4) = RS.Fields(1)
kol = RS.Fields(1)
znes = znes + RS.Fields(2)
If kol = 0 Then
kol = 1
End If
MsfBill.TextMatrix(po, 3) = RS.Fields(2) / kol
MsfBill.TextMatrix(po, 7) = RS.Fields(2)
MsfBill.Rows = MsfBill.Rows + 1
           Indx = Indx + 1
           MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
           FlexgridTotal
po = po + 1
RS.MoveNext

 Loop
 txtTotal.text = Format(znes, "fixed")
 skumi = znes
 zap = Indx
    ind = po
MsfBill.SetFocus
'ArrangeTextbox cmbItmcode
Indx = ind
zap = Indx
 MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
ind = 0


End Function
Private Sub printrac()
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
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
    cPrint.pPrint
    cPrint.pPrint "Zaposlen: " & Me.Label3.Caption
    cPrint.pPrint
    
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
    cPrint.pPrint Me.txtInvoiceNo.text, 1, True
    cPrint.pPrint "z dne " & Format(Date, "dd/mm/yyyy") & " "
    '& Format(Time(), "hh:mm"), 1.6, True
    
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Naziv                   kol      znesek", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    Dim i, ass
    Dim popu As Double
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim ddv2 As Double
    ddv1 = 0
    ddv2 = 0
    popu = 0
    sku = 0
    For i = 1 To MsfBill.Row
    
   If Getnazi("select madapd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1))) = "20" Then
   ddv1 = ddv1 + Val(MsfBill.TextMatrix(i, 7))
   End If
    If Getnazi("select madapd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1))) = "8.5" Then
   ddv2 = ddv2 + Val(MsfBill.TextMatrix(i, 7))
   End If
    stri = Format(MsfBill.TextMatrix(i, 4), "standard")
    stri1 = Format(MsfBill.TextMatrix(i, 7), "standard")
    sku = sku + Val(MsfBill.TextMatrix(i, 7))
    If stri1 <> "" Then
    'MsgBox (Val(Getnazi("select madampcd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1)))) - (Val(MsfBill.TextMatrix(i, 5)) / Val(MsfBill.TextMatrix(i, 4))))
    'If Val(Getnazi("select madampcd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1)))) <> Val(MsfBill.TextMatrix(i, 5)) / Val(MsfBill.TextMatrix(i, 4)) Then
    popu = popu + Val(Getnazi("select madampcd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1)))) - (Val(MsfBill.TextMatrix(i, 7)) / Val(MsfBill.TextMatrix(i, 4)))
    'End If
    End If
    
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint MsfBill.TextMatrix(i, 2), 0.1, True
    cPrint.pRightJust stri, 3, True
    cPrint.pRightJust stri1, 4, True
    Next
   
    cPrint.pPrint ""
    'cPrint.pPrint ""
    cPrint.pPrint "=======================================", 0.1, False
    'cPrint.pPrint ""
    If popu <> 0 Then
    cPrint.pPrint "Popust vracunan v ceni", 0.1, True
    cPrint.pRightJust Format(popu, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
    cPrint.pPrint "ZA PLACILO EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    
    cPrint.pPrint "SKUPAJ SIT", 0.1, True
    cPrint.pRightJust Format(sku * 239.64, "standard"), 4, True
    zavrnit = sku
    
      cPrint.pPrint
    
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.7, True
    cPrint.pRightJust "20 %", 1.2, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 2, True
    cPrint.pRightJust Format(ddv1, "standard"), 2.8, True
    'cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.8, True
    'cPrint.pRightJust " 20 %", 2, True
    'cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3, True
    'cPrint.pRightJust Format(ddv1, "standard"), 4, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.7, True
    cPrint.pRightJust "8.5 %", 1.2, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 2, True
    cPrint.pRightJust Format(ddv2, "standard"), 2.8, True
    
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
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Placilo: " & plax
    cPrint.pPrint Getnazi("select konec1 from oblikar")
    cPrint.pPrint Getnazi("select konec2 from oblikar")
    cPrint.pPrint Getnazi("select konec3 from oblikar")
    cPrint.pPrint Getnazi("select konec4 from oblikar")
    cPrint.pPrint Getnazi("select konec5 from oblikar")
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
     cPrint.pPrint "", 0.1, False
      cPrint.pPrint "", 0.1, False
       cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
   
   
    cPrint.pPrint Chr(27), 0.1, False
     predal
    odrez
    cPrint.pPrint
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

Private Sub veli_Click()
Me.Label10.Caption = Me.veli.text
If veli.text = "VSE" Then
sqlb = ""
Else
sqlb = "select * from swit WHERE [ItemNumber] > " & Val(Me.Label8.Caption) + 1 & " and [command]<>1 AND [Switchboar]=" & Me("nas" & trenu).Tag & " and dim='" & Me.veli.text & "' order by [ItemNumber]"
End If
Hanb (trenu)
End Sub

Private Sub VRNIT_Click()
Form5.Show
End Sub

Private Sub vst5_Click()
printrac2
Me.vst5.Enabled = False
Me.vst5.ForeColor = 0
Dim i, stot, fa
Dim aaa As String

aaa = Left(Time(), 8)
'MsgBox (aaa)
Dim Rsa As New ADODB.Recordset
   If Rsa.State = 1 Then Rsa.Close

 
Rsa.Open "select sifra,naziv,kol,znesek,datum,ura,st,oseba,doza,vst,placilo,sp from racusif", myConection, adOpenStatic, adLockOptimistic
Dim ddd As Integer
Dim vvv As Integer
vvv = MsfBill.Row
For i = 1 To MsfBill.Row
If Val(MsfBill.TextMatrix(i, 1)) <> 0 Then
Rsa.AddNew
    Rsa.Fields(0) = Val(MsfBill.TextMatrix(i, 1))
    Rsa.Fields(1) = MsfBill.TextMatrix(i, 2)
    Rsa.Fields(2) = Val(MsfBill.TextMatrix(i, 4))
    Rsa.Fields(3) = Round(Val(MsfBill.TextMatrix(i, 7)) / vvv, 2)
    Rsa.Fields(4) = Date
    Rsa.Fields(5) = aaa
    
      Rsa.Fields(6) = Me.txtInvoiceNo.text
        Rsa.Fields(7) = Me.Label3.Caption
       
                Rsa.Fields(10) = 1234
       
If Me.stranka.Caption <> "" Then
ddd = Getnazi("select sifra from partner where naziv='" & Me.stranka.Caption & "'")
Else
ddd = 0
End If
        Rsa.Fields(8) = Val(Getnazi("select madadoza from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1))))
        Rsa.Fields(9) = ddd
 End If
    Next
    Rsa.Update
 Rsa.Close
Indx = 1
zap = 1
Me.MsfBill.clear
MsfRefresh
MsfBill.SetFocus
If zap <> 0 Then
MsfBill.Row = zap
Else
MsfBill.Row = 1
End If
MsfBill.Col = 1
ArrangeTextbox cmbItmcode
Indx = 1
MsfBill.TextMatrix(Indx, 0) = Indx
  stot = 0
  fa = Format(stot, "fixed")
txtTotal.text = fa
idstran = 0
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
Indx = 1
zap = 0
MsfBill.Col = 1
           MsfBill.Row = Indx
          MsfBill.TextMatrix(Indx, 0) = Indx
          txtEnter.Visible = False
          ArrangeTextbox cmbItmcode
          Me.kart.Value = False
          skumi = 0
           LaVolpeButton44_Click
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
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
    cPrint.pPrint
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
    cPrint.pPrint Me.txtInvoiceNo.text, 1, True
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
vss = MsfBill.Row

    For i = 1 To MsfBill.Row
   If Getnazi("select madapd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1))) = "20" Then
   ddv1 = ddv1 + Val(MsfBill.TextMatrix(i, 7)) / vss
   End If
    'If Getnazi("select madapd from mada where madasifr=" & Val(MsfBill.TextMatrix(i, 1))) = "8.5" Then
  ' ddv2 = ddv2 + Val(MsfBill.TextMatrix(i, 5)) / vss
   'End If
    stri = Format(MsfBill.TextMatrix(i, 4), "standard")
    stri1 = Format(v / vss, "standard")
    sku = 15
  
cPrint.pPrint "", 0.1, False
    cPrint.pPrint MsfBill.TextMatrix(i, 2), 0.1, True
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
    
    cPrint.pPrint "SKUPAJ SIT", 0.1, True
    cPrint.pRightJust Format(sku * 239.64, "standard"), 4, True
    zavrnit = sku
      cPrint.pPrint
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(sku / 1.2, "standard"), 1.2, True
    cPrint.pRightJust " 20 %", 1.9, True
    cPrint.pRightJust Format(sku - (sku / 1.2), "standard"), 3, True
    cPrint.pRightJust Format(sku, "standard"), 4, True
    End If
     
    End If
    Dim pl As String
    
  
    pl = "V S T O P N I C A"
   
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Placilo: " & pl
    cPrint.pPrint Getnazi("select konec1 from oblikar")
    cPrint.pPrint Getnazi("select konec2 from oblikar")
    cPrint.pPrint Getnazi("select konec3 from oblikar")
    cPrint.pPrint Getnazi("select konec4 from oblikar")
    cPrint.pPrint Getnazi("select konec5 from oblikar")
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
     cPrint.pPrint "", 0.1, False
      cPrint.pPrint "", 0.1, False
       cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
    cPrint.pPrint Chr(27), 0.1, False
    predal
    odrez
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
Open "be1.txt" For Output As #1
'Print #1, Chr(27) & Chr(105)
Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
Close #1
Call Shell("print /d:LPT1 be1.txt", 6)
   
End Sub
Private Sub odrez()
Open "be1.txt" For Output As #1
Print #1, Chr(27) & Chr(105)
'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
Close #1
Call Shell("print /d:LPT1 be1.txt", 6)
   
End Sub

