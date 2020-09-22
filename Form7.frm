VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{ED53EA70-368C-11D0-AD81-00A0C90DC8D9}#1.0#0"; "SNAPVIEW.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   Caption         =   "IZPISI"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form7"
   ScaleHeight     =   10260
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton LaVolpeButton21 
      Height          =   495
      Left            =   3120
      TabIndex        =   48
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "POTRDI POIZVEDBO"
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form7.frx":0000
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
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   240
      TabIndex        =   46
      Top             =   2400
      Width           =   2535
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton11 
      Height          =   255
      Left            =   5880
      TabIndex        =   45
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Uredi report"
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form7.frx":001C
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   300
      Left            =   3120
      TabIndex        =   43
      Top             =   2400
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   529
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form7.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form7.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form7.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   720
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16842753
      CurrentDate     =   40516
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16842753
      CurrentDate     =   40516
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   7200
      TabIndex        =   8
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   4154
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "ŠIFRANTI"
      TabPicture(0)   =   "Form7.frx":008C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LaVolpeButton14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LaVolpeButton13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LaVolpeButton12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DOKUMENTI"
      TabPicture(1)   =   "Form7.frx":00A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LaVolpeButton15(8)"
      Tab(1).Control(1)=   "LaVolpeButton15(7)"
      Tab(1).Control(2)=   "LaVolpeButton15(6)"
      Tab(1).Control(3)=   "LaVolpeButton15(14)"
      Tab(1).Control(4)=   "LaVolpeButton15(13)"
      Tab(1).Control(5)=   "LaVolpeButton15(12)"
      Tab(1).Control(6)=   "LaVolpeButton15(11)"
      Tab(1).Control(7)=   "LaVolpeButton15(10)"
      Tab(1).Control(8)=   "LaVolpeButton15(9)"
      Tab(1).Control(9)=   "LaVolpeButton15(5)"
      Tab(1).Control(10)=   "LaVolpeButton15(4)"
      Tab(1).Control(11)=   "LaVolpeButton15(3)"
      Tab(1).Control(12)=   "LaVolpeButton15(2)"
      Tab(1).Control(13)=   "LaVolpeButton15(0)"
      Tab(1).Control(14)=   "LaVolpeButton15(1)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "ANALIZE"
      TabPicture(2)   =   "Form7.frx":00C4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LaVolpeButton20"
      Tab(2).Control(1)=   "LaVolpeButton19"
      Tab(2).Control(2)=   "LaVolpeButton18"
      Tab(2).Control(3)=   "LaVolpeButton17"
      Tab(2).Control(4)=   "LaVolpeButton16"
      Tab(2).ControlCount=   5
      Begin LVbuttons.LaVolpeButton LaVolpeButton16 
         Height          =   495
         Left            =   -74760
         TabIndex        =   38
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ANALIZA NABAVE"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":00E0
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
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ARTIKLI"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":00FC
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
         Height          =   495
         Left            =   3600
         TabIndex        =   21
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "PARTNERJI"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0118
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
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "GRUPE"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0134
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
         Height          =   495
         Index           =   1
         Left            =   -74880
         TabIndex        =   23
         Tag             =   "NA"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0150
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
         Height          =   495
         Index           =   0
         Left            =   -74880
         TabIndex        =   24
         Tag             =   "NA"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":016C
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
         Height          =   495
         Index           =   2
         Left            =   -74880
         TabIndex        =   25
         Tag             =   "NA"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0188
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
         Height          =   495
         Index           =   3
         Left            =   -73680
         TabIndex        =   26
         Tag             =   "NA"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":01A4
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
         Height          =   495
         Index           =   4
         Left            =   -73680
         TabIndex        =   27
         Tag             =   "NA"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":01C0
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
         Height          =   495
         Index           =   5
         Left            =   -73680
         TabIndex        =   28
         Tag             =   "NA"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":01DC
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
         Height          =   495
         Index           =   9
         Left            =   -70080
         TabIndex        =   29
         Tag             =   "NA"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":01F8
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
         Height          =   495
         Index           =   10
         Left            =   -70080
         TabIndex        =   30
         Tag             =   "NA"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0214
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
         Height          =   495
         Index           =   11
         Left            =   -70080
         TabIndex        =   31
         Tag             =   "NA"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0230
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
         Height          =   495
         Index           =   12
         Left            =   -71280
         TabIndex        =   32
         Tag             =   "NA"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":024C
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
         Height          =   495
         Index           =   13
         Left            =   -71280
         TabIndex        =   33
         Tag             =   "NA"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0268
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
         Height          =   495
         Index           =   14
         Left            =   -71280
         TabIndex        =   34
         Tag             =   "NA"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0284
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
         Height          =   495
         Index           =   6
         Left            =   -72480
         TabIndex        =   35
         Tag             =   "NA"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":02A0
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
         Height          =   495
         Index           =   7
         Left            =   -72480
         TabIndex        =   36
         Tag             =   "NA"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":02BC
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
         Height          =   495
         Index           =   8
         Left            =   -72480
         TabIndex        =   37
         Tag             =   "NA"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "NABAVA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":02D8
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
         Height          =   495
         Left            =   -73320
         TabIndex        =   39
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ANALIZA PRODAJE"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":02F4
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
         Height          =   495
         Left            =   -74760
         TabIndex        =   40
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "PREJEM IZDAJA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0310
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
         Height          =   495
         Left            =   -71880
         TabIndex        =   41
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ANALIZA ZALOG"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":032C
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
         Height          =   495
         Left            =   -70440
         TabIndex        =   42
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "MATERIALNA KARTICA"
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
         BCOL            =   16315377
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Form7.frx":0348
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
   Begin MSComctlLib.ImageList IML 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   225
      ImageHeight     =   225
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":0364
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":1902
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":2D21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":478B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":4ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":5594
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":5B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":6150
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":67F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":78ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form7.frx":88E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form7.frx":9960
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":997C
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "5"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton8 
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":9998
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form7.frx":99B4
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form7.frx":99D0
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":99EC
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "6"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton6 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":9A08
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "7"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton7 
      Height          =   615
      Left            =   4440
      TabIndex        =   7
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":9A24
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "8"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton9 
      Height          =   615
      Left            =   5160
      TabIndex        =   10
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":9A40
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "9"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton10 
      Height          =   615
      Left            =   5760
      TabIndex        =   11
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "Form7.frx":9A5C
      ALIGN           =   1
      IMGLST          =   "IML"
      IMGICON         =   "10"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSForms.Frame Frame1 
      Height          =   735
      Left            =   0
      OleObjectBlob   =   "Form7.frx":9A78
      TabIndex        =   12
      Top             =   720
      Width           =   2895
   End
   Begin MSForms.Frame Frame2 
      Height          =   735
      Left            =   0
      OleObjectBlob   =   "Form7.frx":A490
      TabIndex        =   15
      Top             =   1440
      Width           =   2895
   End
   Begin MSForms.Frame Frame3 
      Height          =   735
      Left            =   3120
      OleObjectBlob   =   "Form7.frx":AEA8
      TabIndex        =   17
      Top             =   1080
      Width           =   3975
   End
   Begin SnapshotViewerControlCtl.SnapshotViewer SnapshotViewer1 
      Height          =   5415
      Left            =   3120
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9551
      _Version        =   65536
      SnapshotPath    =   ""
      Zoom            =   4
      AllowContextMenu=   -1  'True
      ShowNavigationButtons=   0   'False
   End
   Begin MSForms.Frame Frame4 
      Height          =   735
      Left            =   0
      OleObjectBlob   =   "Form7.frx":B8C0
      TabIndex        =   47
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private zadnjstr As Integer
Private trenut As Integer
Private printord, printgrp As String




Sub osv()
'xrep = Me.Text1.Text
'Dim db As CRAXDRT.Database
Screen.MousePointer = vbHourglass
'If rs.State = 1 Then rs.Close
'MsgBox (printsql)
'rs.Open printsql & printgrp & printord, myConection, adOpenDynamic, adLockReadOnly
' arepo.Text2.SetText "Hello World"
' crRep.Sections(2).ReportObjects(1).SetText "Hello World"
'PRINTSNAP "mada"
Me.SnapshotViewer1.Visible = False
If repor = "nabasif" Then
SNAP "nabasif", "tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp.snp"

Else
SNAP "mada", , , printsql & printord
Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp.snp"

SNAP1 "mada1", , , printsql & printord
SNAP2 "mada2", , , printsql & printord
End If

'Me.SnapshotViewer1.FirstPage
Me.SnapshotViewer1.Visible = True

Me.SnapshotViewer1.NextPage
Me.SnapshotViewer1.FirstPage

Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
If Me.Combo1.Text <> "" Then
printord = " order by " & Trim(Me.Combo1.Text)
Else
printord = ""
End If
osv
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Combo3_Change()
'Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp2.snp"
End Sub

Private Sub Form_Load()
'SQLREP = Me.Text2.Text


Me.DTPicker1.Value = Date
Me.DTPicker2.Value = Date
'osv
osvcon
Me.Timer1.Enabled = True
If rs.State = 1 Then rs.Close
rs.Open "select tip_dok,opis from dokumenti", myConection, adOpenDynamic, adLockOptimistic
rs.MoveFirst
Dim xx As Integer
xx = 0
Do While Not rs.EOF

Me.LaVolpeButton15(xx).Caption = rs.Fields("opis")
Me.LaVolpeButton15(xx).Tag = rs.Fields("tip_dok")
Me.LaVolpeButton15(xx).Visible = True
xx = xx + 1
rs.MoveNext
Loop

End Sub
Private Sub osvcon()
Dim objRs As New ADODB.Recordset
Dim objFields As ADODB.Fields
    Dim intLoop As Integer
    If objRs.State = 1 Then objRs.Close
    
    objRs.Open printsql & printord, myConection, adOpenDynamic, adLockOptimistic
    
    Set objFields = objRs.Fields
    Me.Combo1.clear
    Me.Combo1.AddItem ""
    For intLoop = 0 To (objFields.Count - 1)
     
        Me.Combo1.AddItem objFields.Item(intLoop).Name
    Next



End Sub
Private Sub osvconart()
Dim objRs As New ADODB.Recordset

    Dim intLoop As Integer
    If objRs.State = 1 Then objRs.Close
    
    objRs.Open "select madasifr,madanazi from mada", myConection, adOpenDynamic, adLockOptimistic
    
    Me.Combo3.clear
    'Me.Combo1.AddItem ""
    objRs.MoveFirst
    Do While Not objRs.EOF
     
        Me.Combo3.AddItem objRs.Fields("madanazi")
        objRs.MoveNext
    Loop


End Sub
Private Sub Form_Resize()
Me.SSTab2.Top = Me.SSTab1.Height + 40
Me.SSTab2.Left = Me.List1.Width + 10
'Me.SSTab2.Height = ScaleHeight - Me.SSTab1.Height - 30
Me.SSTab2.Width = ScaleHeight - Me.SSTab1.Height - 30
Me.List1.Top = Me.SnapshotViewer1.Top + 20
Me.List1.Height = ScaleHeight - Me.SSTab1.Height - 30
'Me.SnapshotViewer1.Top = Me.SSTab1.Height + 40
Me.SnapshotViewer1.Width = ScaleWidth - Me.List1.Width - 20
'Me.SnapshotViewer1.Left = Me.List1.Width + 10
Me.SnapshotViewer1.Height = ScaleHeight - Me.SSTab1.Height - 30

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set crystal = Nothing
Set Report = Nothing
End Sub

Private Sub LaVolpeButton1_Click()

Dim x As Long
Dim blRet As Boolean
Dim sPDF As String
Dim sName As String
'sName = "dizp.snp"
sName = JUSTFileName(Me.SnapshotViewer1.SnapshotPath)
If Len(sName & vbNullString) = 0 Then Exit Sub
' let's use the name of the selected Snapshot file
' to name our converted PDF document.



' Debug Stress test
For x = 1 To 1  '1000
sPDF = Mid(sName, 1, Len(sName) - 4)

blRet = ConvertReportToPDF(vbNullString, sName, sPDF & x & ".PDF", False, True, 0, "", "", 0, 1)
Next x
ShellExecute 0&, "open", App.path & sPDF & x & ".PDF", "", "", vbNormalFocus
End Sub

Private Sub LaVolpeButton10_Click()
If SnapshotViewer1.Zoom < 8 Then
SnapshotViewer1.Zoom = SnapshotViewer1.Zoom + 1
'Me.Label2.Caption = Trim(str(SnapshotViewer1.Zoom))
End If
End Sub

Private Sub LaVolpeButton11_Click()
uredirepo "mada"
End Sub

Private Sub LaVolpeButton12_Click()
Me.SSTab2.Tab = 2
Me.SSTab2.Caption = "Kartica"
Me.SSTab2.Tab = 1
Me.SSTab2.Caption = "Razširjen pregled"
Me.SSTab2.Tab = 0
Me.SSTab2.Caption = "Osnovni pregled"

Me.Combo1.Enabled = True

Me.Combo2.Enabled = True
printsql = "select * from mada"
Me.Combo1.Text = ""

repor = "mada"
repor1 = "mada1"
repor2 = "mada2"
osvconart
osvcon
osv
End Sub

Private Sub LaVolpeButton13_Click()
Me.Combo1.Text = ""
Me.Combo1.Enabled = True
Me.Combo2.Enabled = True
printsql = "select * from partner"
osvcon
osv
End Sub

Private Sub LaVolpeButton14_Click()
Me.Combo1.Text = ""
Me.Combo1.Enabled = True
Me.Combo2.Enabled = True
printsql = "select * from grupa"
osvcon
osv
End Sub

Private Sub LaVolpeButton15_Click(Index As Integer)
Me.Combo1.Text = ""
Me.Combo2.Text = ""
Me.Combo1.Enabled = False
Me.Combo2.Enabled = False
doddok (Me.LaVolpeButton15(Index).Tag)
'printsql = "SELECT nabasif.id_dok, PARTNER.SIFRA, PARTNER.NAZIV, nabasif.tip_dok, PARTNER.MESTO, PARTNER.ULICA, PARTNER.POSTA, PARTNER.DAVCNA, nabasif.DATUM, nabasif.SIFRA, nabasif.KOL, nabasif.CENA, nabasif.ZNES, nabasif.pop, nabasif.naziv  FROM   (glavna glavna INNER JOIN nabasif nabasif ON (glavna.id_dok=nabasif.id_dok) AND (glavna.tip_dok=nabasif.tip_dok)) INNER JOIN PARTNER PARTNER ON glavna.dod0=PARTNER.NAZIV where nabasif.id_dok='PA'  ORDER BY nabasif.id_dok,nabasif.pozicija"

End Sub
Private Sub doddok(atip As String)
Dim rst As New ADODB.Recordset
tip_dok = atip

Me.List1.clear
rst.Open "select tip_dok,id_dok ,datum,sum(znes) as znesek from nabasif where tip_dok='" & atip & "' group by tip_dok,id_dok,datum order by id_dok desc", myConection, adOpenDynamic, adLockOptimistic
If Not rst.EOF Then
rst.MoveFirst
Do While Not rst.EOF
List1.AddItem rst.Fields("tip_dok") & rst.Fields("id_dok") & "  " & Format(rst.Fields("datum"), "DD.MM.YYYY") & "  " & FormatNumber(rst.Fields("znesek"), 2)
rst.MoveNext
Loop
End If

'PRINTREP = "dokument.rpt"
End Sub



Private Sub LaVolpeButton2_Click()
'Dim crexp As New Crystalarepo1

End Sub

Private Sub LaVolpeButton21_Click()
SNAP2 "mada2", "sifra='" & Getnazi("select madasifr from mada where madanazi='" & Me.Combo3.Text & "'") & "'", , printsql & printord
Me.SSTab2.Tab = 2

End Sub

Private Sub LaVolpeButton3_Click()
'Dim crexp As New Crystalarepo1

End Sub

Private Sub LaVolpeButton4_Click()
On Error GoTo bbb:
Me.SnapshotViewer1.FirstPage
Me.Label1.Caption = Me.SnapshotViewer1.CurrentPage & "/" & Me.SnapshotViewer1.PageCount
bbb:
End Sub

Private Sub LaVolpeButton5_Click()
On Error GoTo bbb:
Me.SnapshotViewer1.PreviousPage
Me.Label1.Caption = Me.SnapshotViewer1.CurrentPage & "/" & Me.SnapshotViewer1.PageCount
bbb:
End Sub

Private Sub LaVolpeButton6_Click()
On Error GoTo bbb:
Me.SnapshotViewer1.NextPage
Me.Label1.Caption = Me.SnapshotViewer1.CurrentPage & "/" & Me.SnapshotViewer1.PageCount

bbb:
End Sub

Private Sub LaVolpeButton7_Click()
On Error GoTo bbb:
Me.SnapshotViewer1.LastPage
Me.Label1.Caption = Me.SnapshotViewer1.CurrentPage & "/" & Me.SnapshotViewer1.PageCount
bbb:
End Sub

Private Sub LaVolpeButton8_Click()
Me.SnapshotViewer1.PrintSnapshot True
End Sub

Private Sub Text1_Change()
osv
End Sub

Private Sub Text2_LostFocus()
'SQLREP = Me.Text2.Text
'osv
End Sub

Private Sub LaVolpeButton9_Click()
If SnapshotViewer1.Zoom > 0 Then
SnapshotViewer1.Zoom = SnapshotViewer1.Zoom - 1
'Me.Label2.Caption = Trim(str(SnapshotViewer1.Zoom))
End If

'Val(Me.Label2.Caption) + 25)) & " %"
End Sub

Private Sub List1_Click()
xid_dok = Mid(Me.List1.Text, 3, 8)
repor = "nabasif"
printsql = "SELECT * from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Mid(Me.List1.Text, 3, 8) & "'  ORDER BY id_dok,pozicija"
'MsgBox (printsql)
osv
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)

If Me.SSTab2.Tab = 0 Then
' SNAP1 "mada1", , , printsql & printord

Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp.snp"
End If

If Me.SSTab2.Tab = 1 Then
' SNAP1 "mada1", , , printsql & printord
Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp1.snp"
End If

If Me.SSTab2.Tab = 2 Then
'SNAP2 "mada2", , , printsql & printord

Me.SnapshotViewer1.SnapshotPath = App.path & "\dizp2.snp"
End If
'SNAP2 "mada2", , , printsql & printord

End Sub

