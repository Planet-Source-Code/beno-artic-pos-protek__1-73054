VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form pocket 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11070
   ClientLeft      =   15
   ClientTop       =   240
   ClientWidth     =   14520
   ClipControls    =   0   'False
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
   ScaleHeight     =   11070
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton shran 
      Height          =   270
      Left            =   8640
      MaskColor       =   &H8000000F&
      Picture         =   "pocket.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   9120
      Width           =   270
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
      Height          =   450
      Left            =   13080
      TabIndex        =   61
      Text            =   "Uporabnik:"
      Top             =   10440
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
      Height          =   450
      Left            =   13080
      TabIndex        =   60
      Text            =   "FIRMA:"
      Top             =   10440
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   57
      Text            =   "1"
      Top             =   9000
      Width           =   855
   End
   Begin LVbuttons.LaVolpeButton plu 
      Height          =   855
      Left            =   5640
      TabIndex        =   54
      Top             =   9000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
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
      MICON           =   "pocket.frx":00FA
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   53
      Top             =   9000
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   855
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
      Height          =   450
      Left            =   12600
      TabIndex        =   47
      Top             =   10440
      Width           =   2235
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   540
      Left            =   13560
      ScaleHeight     =   480
      ScaleWidth      =   615
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   10200
      Visible         =   0   'False
      Width           =   675
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
         Left            =   0
         TabIndex        =   44
         Top             =   360
         Width           =   3405
      End
   End
   Begin VB.TextBox mii 
      Height          =   435
      Left            =   0
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   0
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
      Height          =   1335
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   1215
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0116
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
         Height          =   1215
         Left            =   0
         TabIndex        =   5
         Top             =   2400
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0132
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
         Height          =   1215
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":014E
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
         Height          =   1215
         Left            =   0
         TabIndex        =   7
         Top             =   3600
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":016A
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
         Height          =   1215
         Left            =   0
         TabIndex        =   8
         Top             =   6000
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0186
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
         Height          =   1215
         Left            =   0
         TabIndex        =   9
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":01A2
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   10
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":01BE
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   11
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":01DA
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   12
         Top             =   2400
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":01F6
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   13
         Top             =   3600
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0212
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   14
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":022E
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
         Height          =   1215
         Left            =   1800
         TabIndex        =   15
         Top             =   6000
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":024A
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   16
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0266
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   17
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0282
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   18
         Top             =   2400
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":029E
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   19
         Top             =   3600
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":02BA
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   20
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":02D6
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
         Height          =   1215
         Left            =   3600
         TabIndex        =   21
         Top             =   6000
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":02F2
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   22
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":030E
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   23
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":032A
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   24
         Top             =   2400
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0346
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   25
         Top             =   3600
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0362
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   26
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":037E
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   27
         Top             =   6000
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":039A
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   95
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":03B6
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   96
         Top             =   1200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":03D2
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   97
         Top             =   2400
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":03EE
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   98
         Top             =   3600
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":040A
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   99
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0426
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
         Height          =   1215
         Left            =   7200
         TabIndex        =   100
         Top             =   6000
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
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
         MICON           =   "pocket.frx":0442
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
      Left            =   9600
      Top             =   10560
   End
   Begin VB.TextBox txtInvoiceNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   10200
      TabIndex        =   1
      Text            =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1515
   End
   Begin LVbuttons.LaVolpeButton nas1 
      Height          =   990
      Index           =   0
      Left            =   1080
      TabIndex        =   29
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":045E
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
      Left            =   13200
      TabIndex        =   30
      Top             =   10080
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
      MICON           =   "pocket.frx":047A
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
      Left            =   4920
      TabIndex        =   31
      Top             =   9960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "STORNO - F3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   27.75
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
      MICON           =   "pocket.frx":0496
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
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":04B2
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
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":04CE
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
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":04EA
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
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   35
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":0506
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
      Height          =   855
      Index           =   5
      Left            =   0
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":0522
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
      Height          =   855
      Index           =   6
      Left            =   0
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":053E
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
      Height          =   855
      Index           =   7
      Left            =   0
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":055A
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
      Height          =   855
      Index           =   8
      Left            =   0
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":0576
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
      Height          =   855
      Index           =   9
      Left            =   0
      TabIndex        =   40
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":0592
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
      Height          =   855
      Index           =   10
      Left            =   0
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LaVolpeButton"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   24
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
      MICON           =   "pocket.frx":05AE
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
      TabIndex        =   46
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
      MICON           =   "pocket.frx":05CA
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
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   52
      Top             =   9960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "ZAKLJUCI - F4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   27.75
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
      MICON           =   "pocket.frx":05E6
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
      Height          =   855
      Left            =   6480
      TabIndex        =   55
      Top             =   9000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
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
      MICON           =   "pocket.frx":0602
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
      Height          =   855
      Left            =   7320
      TabIndex        =   56
      Top             =   9000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
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
      MICON           =   "pocket.frx":061E
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
      DragIcon        =   "pocket.frx":063A
      Height          =   1560
      Left            =   960
      TabIndex        =   58
      Top             =   3840
      Visible         =   0   'False
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   2752
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
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   36
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
      Height          =   990
      Index           =   1
      Left            =   2400
      TabIndex        =   62
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0944
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
      Height          =   990
      Index           =   2
      Left            =   3720
      TabIndex        =   63
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0960
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
      Height          =   990
      Index           =   3
      Left            =   5040
      TabIndex        =   64
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":097C
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
      Height          =   990
      Index           =   4
      Left            =   6360
      TabIndex        =   65
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0998
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
      Height          =   990
      Index           =   5
      Left            =   7680
      TabIndex        =   66
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":09B4
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
      Height          =   990
      Index           =   6
      Left            =   9000
      TabIndex        =   67
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":09D0
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
      Height          =   990
      Index           =   7
      Left            =   10320
      TabIndex        =   68
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":09EC
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
      Height          =   990
      Index           =   8
      Left            =   11640
      TabIndex        =   69
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A08
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
      Height          =   990
      Index           =   9
      Left            =   12960
      TabIndex        =   70
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A24
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
      Height          =   990
      Index           =   10
      Left            =   1080
      TabIndex        =   71
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A40
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
      Height          =   990
      Index           =   11
      Left            =   2400
      TabIndex        =   72
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A5C
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
      Height          =   990
      Index           =   12
      Left            =   3720
      TabIndex        =   73
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A78
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
      Height          =   990
      Index           =   13
      Left            =   5040
      TabIndex        =   74
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0A94
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
      Height          =   990
      Index           =   14
      Left            =   6360
      TabIndex        =   75
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0AB0
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
      Height          =   990
      Index           =   15
      Left            =   7680
      TabIndex        =   76
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0ACC
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
      Height          =   990
      Index           =   16
      Left            =   9000
      TabIndex        =   77
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0AE8
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
      Height          =   990
      Index           =   17
      Left            =   10320
      TabIndex        =   78
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B04
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
      Height          =   990
      Index           =   18
      Left            =   11640
      TabIndex        =   79
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B20
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
      Height          =   990
      Index           =   19
      Left            =   12960
      TabIndex        =   80
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B3C
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
      Height          =   990
      Index           =   20
      Left            =   1080
      TabIndex        =   81
      Top             =   2040
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B58
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
      Height          =   990
      Index           =   21
      Left            =   2400
      TabIndex        =   82
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B74
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
      Height          =   990
      Index           =   22
      Left            =   3720
      TabIndex        =   83
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0B90
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
      Height          =   990
      Index           =   23
      Left            =   5040
      TabIndex        =   84
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0BAC
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
      Height          =   990
      Index           =   24
      Left            =   6360
      TabIndex        =   85
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0BC8
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
      Height          =   990
      Index           =   25
      Left            =   7680
      TabIndex        =   86
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0BE4
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
      Height          =   990
      Index           =   26
      Left            =   9000
      TabIndex        =   87
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0C00
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
      Height          =   990
      Index           =   27
      Left            =   10320
      TabIndex        =   88
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0C1C
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
      Height          =   990
      Index           =   28
      Left            =   11640
      TabIndex        =   89
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0C38
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
      Height          =   990
      Index           =   29
      Left            =   12960
      TabIndex        =   90
      Top             =   1920
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1746
      BTYPE           =   3
      TX              =   ""
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
      COLTYPE         =   2
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "pocket.frx":0C54
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
      Height          =   8415
      Left            =   0
      TabIndex        =   93
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   14843
      BTYPE           =   3
      TX              =   "Mize "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
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
      MICON           =   "pocket.frx":0C70
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   1
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label4 
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
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   495
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
      Height          =   255
      Left            =   8520
      TabIndex        =   91
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label znees 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   9120
      TabIndex        =   59
      Top             =   9480
      Width           =   3735
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
      Left            =   12000
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
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
      Height          =   375
      Left            =   -120
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      Height          =   495
      Left            =   13560
      TabIndex        =   49
      Top             =   10200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   495
      Left            =   10200
      TabIndex        =   48
      Top             =   9360
      Visible         =   0   'False
      Width           =   375
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
      Height          =   375
      Left            =   12240
      TabIndex        =   45
      Top             =   10320
      Width           =   3015
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
      Height          =   375
      Left            =   12240
      TabIndex        =   28
      Top             =   10440
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rac.st.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Menu mnugru 
      Caption         =   "GRUPE"
   End
   Begin VB.Menu mnurac 
      Caption         =   "POGLED RACUNA"
   End
   Begin VB.Menu mnunaj 
      Caption         =   "NAJ"
   End
   Begin VB.Menu mnuizh 
      Caption         =   "IZHOD"
   End
End
Attribute VB_Name = "pocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim gSlno, gItemCode, gItemname, gQty, gRate, gTotal, gpop, Inti, miz, i
Dim Indx
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
prvay = 0
If xzago <> 1 Then
coda852
blagajna = 1
For miz = 1 To 10
mizaa(miz).Caption = miz
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
         nas1(aad).Tag = rs![ARGUMENT]
        
        aad = aad + 1
            rs.MoveNext
        Wend
        aad = 0
      Do While Not aad = 30
      
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

End If
osssv
mnugru_Click
If Left(LTrim(Me.LaVo1.Caption), 6) = "SHRANI" Then
mnurac_Click
End If
End Sub
Private Sub glava()
Dim ii As Integer
Me.text1.FontSize = Me.text1.FontSize * 4
Me.Text2.FontSize = Me.Text2.FontSize * 4
Me.Text3.FontSize = Me.Text3.FontSize * 4
Me.text1.Width = Me.text1.Width * 1.4
Me.Text2.Width = Me.Text2.Width * 1.7
Me.Text3.Width = Me.Text3.Width * 1.4
Me.text1.Height = Me.text1.Height * 1.3
Me.Text2.Height = Me.Text2.Height * 1.3
Me.Text3.Height = Me.Text3.Height * 1.3
Me.Text2.Left = Me.text1.Left + Me.text1.Width
Me.Text3.Left = Me.Text2.Left + Me.Text2.Width
Me.plu.Width = Me.plu.Width * 1.7
Me.min.Width = Me.min.Width * 1.7
Me.bri.Width = Me.bri.Width * 1.7
Me.plu.Height = Me.plu.Height * 1.2
Me.min.Height = Me.min.Height * 1.2
Me.bri.Height = Me.bri.Height * 1.2
Me.plu.Font.Size = Me.plu.Font.Size * 2
Me.min.Font.Size = Me.min.Font.Size * 2
Me.bri.Font.Size = Me.bri.Font.Size * 2

Me.plu.Left = Me.Text3.Left + Me.Text3.Width
Me.min.Left = Me.plu.Left + Me.plu.Width
Me.bri.Left = Me.min.Left + Me.min.Width
Me.shran.Left = Me.bri.Left + Me.bri.Width
For ii = 0 To 29
Me.nas1(ii).Width = Me.nas1(ii).Width * 2
Me.nas1(ii).CaptionOrientation = Horizontal
Me.nas1(ii).Height = Me.nas1(ii).Height * 1.5
Me.nas1(ii).Font.Size = Me.nas1(ii).Font.Size * 1.8
Me.nas1(ii).Font.Bold = True
If ii < 29 Then
Me("LaVolpeButton" & ii + 1).Width = Me.nas1(ii).Width
Me("LaVolpeButton" & ii + 1).Font.Size = Me.nas1(ii).Font.Size
Me("LaVolpeButton" & ii + 1).Font.Bold = True
If ii + 1 > 6 Then
Me("LaVolpeButton" & ii + 1).Left = Me("LaVolpeButton" & ii - 5).Left + Me("LaVolpeButton" & ii - 5).Width
'Me("LaVolpeButton" & ii + 1).Top = Me("LaVolpeButton" & ii - 5).Top + Me("LaVolpeButton" & ii - 5).Height

End If


End If

Next
For ii = 1 To 4
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.nas1(5).Left = Me.nas1(0).Left
Me.nas1(5).Top = Me.nas1(0).Top + Me.nas1(0).Height
For ii = 6 To 9
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.nas1(10).Left = Me.nas1(0).Left
Me.nas1(10).Top = Me.nas1(9).Top + Me.nas1(9).Height
For ii = 11 To 14
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.nas1(15).Left = Me.nas1(0).Left
Me.nas1(15).Top = Me.nas1(14).Top + Me.nas1(14).Height
For ii = 16 To 19
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.nas1(20).Left = Me.nas1(0).Left
Me.nas1(20).Top = Me.nas1(19).Top + Me.nas1(19).Height
For ii = 21 To 24
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.nas1(25).Left = Me.nas1(0).Left
Me.nas1(25).Top = Me.nas1(24).Top + Me.nas1(24).Height

For ii = 26 To 29
Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
Me.nas1(ii).Top = Me.nas1(ii - 1).Top
Next
Me.MSHFlexGrid1.Height = Me.Text3.Top
Me.MSHFlexGrid1.Font.Size = Me.MSHFlexGrid1.Font.Size * 4
'Me.nas1(25).Left = Me.nas1(0).Left
'Me.nas1(25).Top = Me.nas1(24).Top + Me.nas1(24).Height
'For ii = 25 To 29
'Me.nas1(ii).Left = Me.nas1(ii - 1).Left + Me.nas1(ii - 1).Width
'Me.nas1(ii).Top = Me.nas1(ii - 1).Top
'Next
Me.Frame1.Height = Me.Text2.Top
   
Me.znees.Left = Me.LaVolpeButton45.Left + Me.LaVolpeButton45.Width
Me.znees.Top = Me.LaVolpeButton45.Top - (Me.LaVolpeButton45.Height / 4)
Me.znees.Font.Size = Me.znees.Font.Size * 4
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
glava
'MsfRefresh
'FillCombo cmbItmcode, "select MADASIFR from MADA"
 LaVo1_Click

Dim cx As Integer
Me.Frame1.Visible = False
    For cx = 0 To 29
    If nas1(cx).Caption <> "" Then
         
    Me.nas1(cx).Visible = True
    End If
    Next
    Me.MSHFlexGrid1.Visible = False
End Sub




Private Sub Image1_Click()
End
End Sub


Private Sub LaVo1_Click()
If Me.Label11.Caption = "" Then
scamize.Show vbModal
Else
shranimi (Me.Label11.Caption)
Me.MSHFlexGrid1.clear
osssv
End If
End Sub

Private Sub LaVo2_Click()



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
Hanbt (25)
End Sub

'Private Sub LaVolpeButton25_Click()
' If rs.State = 1 Then rs.Close
'   rs.Open "select * from swit WHERE [ItemNumber] > 0 AND [Switchboar]=1 order by itemnumber"
'      rs.MoveFirst
'      Dim aad As Integer
'      aad = 0
'      If Not rs.EOF Then'

       'While (Not (rs.EOF))
       'aad = aad + 1
       '  Me("nas" & aad).Caption = rs![ITEMTEXT]
       '  Me("nas" & aad).Tag = rs![SWITCHBOAR]
       '     rs.MoveNext
       ' Wend
     ' Else
     '    End If



'End Sub
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


  
   
rs.Close
'       fora = 2
'Me.ListBox1.clear
refr = 1
stm1 = 0
   
'Me.Frame2.Visible = False
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

Unload Me
End Sub

Private Sub LaVolpeButton45_Click()
If stalnaprij = 1 Then
Else
prijavljen = ""
End If
'Me.kart.Value = 0
'Me.inter.Value = 0
myConection.Execute ("delete from trenutna where  stdok='" & Pblagajna & "'")
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
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
Indx = 1
zap = 0
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE  "
Me.Label12.Caption = ""
Me.text1.Text = ""
Text2.Text = ""
Text3.Text = 1
text1.SetFocus
'Me.dav.Visible = False
'Me.dav.Text = ""
''Me.davlb.Visible = False
'Me.imes.Visible = False
'Me.imes.Text = ""
'Me.nassl.Visible = False
'Me.nassl.Text = ""
LaVo1_Click
End Sub

Private Sub LaVolpeButton46_Click()
'If Me.MSHFlexGrid1.Col = 4 Then
'End If





    Dim strf As Integer
   
    strf = 0
     plax = "GOTOVINA"
   
    If strf = 0 Then
    
    End If
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='POPPA'") = "D" Then
printrac
Else
printrac2

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
    'If Val(MSHFlexGrid1.TextMatrix(i, 6)) = 0 Then
    'Rsa.Fields("pop") = 0
    'Else
     'Rsa.Fields("pop") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 6), 2)
    'End If
    Rsa.Fields("mpc") = Getcena(MSHFlexGrid1.TextMatrix(i, 0), Date)


    Rsa.Fields("ZNES") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    Rsa.Fields("cena") = FormatNumber(MSHFlexGrid1.TextMatrix(i, 3), 2)
    
    Rsa.Fields("pozicija") = levi_pres(LTrim(str(i)), 4)
    Rsa.Fields("DAT_K") = datum
     Rsa.Fields("skl") = "GOS"
    Rsa.Fields("DATUM") = Date
   
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
        
        
Rsa.Update
Dim sqlll As String
Dim ses_ko As Double
Dim ses_si As String
ses_ko = 0
ses_si = ""
If Getnazi("select sifras from sestavi where sifra=" & Trim(Rsa.Fields("sifra"))) = "" Then
sqlll = "update mada set madazalo=madazalo-" & Replace(Rsa.Fields("KOL") * IIf(xdoz > 0, xdoz, 1), ",", ".") & " where madasifr='" & Trim(Rsa.Fields("sifra")) & "'"
Else
ses_si = LTrim(RTrim(Getnazi("select sifras from sestavi where sifra=" & Trim(Rsa.Fields("sifra")))))
ses_ko = Getnumb("select kol from sestavi where sifra=" & Trim(Rsa.Fields("sifra")))
sqlll = "update mada set madazalo=madazalo-" & Rsa.Fields("KOL") * Replace(ses_ko, ",", ".") & " where madasifr='" & Trim(ses_si) & "'"
End If
myConection.Execute (sqlll)
    
       
 
    Next
    
 myConection.Execute ("insert into glavna (tip_dok,id_dok,skl) values ('PA','" & ddid & "','GOS')")
    
 Rsa.Close
Indx = 1
zap = 1
osssv
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
'Me.UserControl11.Visible = False
mi
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE  "
Me.Label12.Caption = ""
LaVolpeButton45_Click
'LaVo1_Click
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
If mizaa(Index).BackColor = 14215660 Then
shranimi (Index)


Else
odprimi (Index)
Dim sSQL As String
    
    'default
    
    
  '  sSQL = "DELETE * FROM mize WHERE stmize=" & Index
  '  myConection.Execute sSQL
  '  mizaa(Index).BackColor = 14215660
'MSHFlexGrid1.SetFocus
fora = 9


End If

End Sub





Private Sub mnugru_Click()
Dim cx As Integer
Me.Frame1.Visible = False
    For cx = 0 To 29
    
         If nas1(cx).Caption <> "" Then
          Me.nas1(cx).Visible = True
          End If
    Next
    Me.MSHFlexGrid1.Visible = False
End Sub

Private Sub mnuizh_Click()
blagajna = 0

Unload Me
End Sub

Private Sub mnurac_Click()
If Me.MSHFlexGrid1.Visible = False Then
If LTrim(RTrim(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Rows - 1, 0))) = "" Then
Exit Sub
End If
Dim cx As Integer
Me.Frame1.Visible = False
    For cx = 0 To 29
    Me.nas1(cx).Visible = False
    Next

Me.MSHFlexGrid1.Visible = True
Else
Me.MSHFlexGrid1.Visible = False
Me.Frame1.Visible = True
Me.Frame1.Refresh
End If
'Me.MSHFlexGrid1.Font.Size = Me.MSHFlexGrid1.Font.Size * 3
End Sub

Private Sub MSHFlexGrid1_Click()
If Me.MSHFlexGrid1.FixedCols = 1 Then

Else
Me.text1.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)
Me.Text2.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1)
Me.Text3.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
'Me.pop.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)
End If

End Sub



Private Sub MSHFlexGrid1_SelChange()
If Me.MSHFlexGrid1.FixedCols = 1 Then

Else
Me.text1.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)
Me.Text2.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1)
Me.Text3.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
'Me.pop.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)
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
Private Sub osssv()
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

'myConection.Execute ("update trenutna set pop=" & Me.pop.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
'myConection.Execute ("update trenutna set cena=cena*(1-(pop/100)) where x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=cena*(1-(pop/100))*kol where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssv
Me.Text3.Text = 1
Me.Text2.Text = ""
Me.text1.Text = ""
Me.text1.SetFocus

End If
End Sub

Private Sub pop_LostFocus()
'Me.pop.Text = 0
End Sub

Private Sub pred_Click()

predal
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
' LaVolpeButton2522.SetFocus
'  Sendkeys "{enter}"
 Case vbKeyF2
 
  Case vbKeyF9
  Case vbKeyF10
 ' LaVolpeButton44.SetFocus
 '  Sendkeys "{enter}"
 Case vbKeyF8
 'LaVo2.SetFocus
 ' Sendkeys "{enter}"
 Case vbKeyF6
fora = 1
 Me.mii.Visible = True
 Me.mii.Text = ""
 
 Me.mii.SetFocus
 
 
Case vbKeyF4
 zakljucc(0).SetFocus
 Sendkeys "{enter}"

 Case vbKeyA To vbKeyZ

Dim iid As String
vrjenniz = ""
'idar = Chr(KeyCode)
'   DoSQL "mada", "madasifr", "madanazi", "madanaz1"
       iskalni = text1.Text & Chr(KeyCode)
       pritisk = text1.Text & Chr(KeyCode)
      ' DoSQL = ""
      Dim ax As String
      
       ax = DoSQLbe("mada", "madasifr", "madanazi", "madanaz1")
     'MsgBox (ax)
      Me.text1.Text = ax
      If Getnazi("select madanazi from mada where madasifr='" & Trim(ax) & "'") <> "" Then
Me.Text2.Text = Getnazi("select madanazi from mada where madasifr='" & Trim(ax) & "'")
Me.Text3.Text = 1
'Me.pop.Text = 0
Dim rsa1 As New ADODB.Recordset
rsa1.Open "select sifra,naziv,kol,znes,x,cena,pop,stdok from trenutna where  stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
rsa1.AddNew
rsa1.Fields("sifra") = Me.text1.Text
rsa1.Fields("naziv") = Me.Text2.Text
rsa1.Fields("kol") = Me.Text3.Text
'rsa1.Fields("pop") = Me.pop.Text
rsa1.Fields("stdok") = Pblagajna
rsa1.Fields("cena") = Getnazi("select madampcd from mada where madasifr='" & Trim(ax) & "'")
'rsa1.Fields("mpc") = Getcena(Me.Text1.text)

rsa1.Fields("znes") = Getnazi("select madampcd from mada where madasifr='" & Trim(ax) & "'") * Me.Text3.Text
rsa1.Fields("x") = Getnumb("select max(x) as x  from trenutna where  stdok='" & Pblagajna & "'") + 1
rsa1.Update
osssv

Me.Text3.SetFocus
Sendkeys "+{RIGHT}"
End If
Case Else
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not KeyAscii >= 48 And Not KeyAscii <= 57 Then
'Exit Sub
End If
If KeyAscii = 27 Then
Me.text1.Text = ""
End If
If KeyAscii = 13 Then
If Getnazi("select madanazi from mada where madasifr='" & Trim(Me.text1.Text) & "'") <> "" Then
Me.Text2.Text = Getnazi("select madanazi from mada where madasifr='" & Trim(Me.text1.Text) & "'")
Me.Text3.Text = 1
Dim rsa1 As New ADODB.Recordset
rsa1.Open "select sifra,naziv,kol,znes,x,cena,pop,stdok from trenutna where  stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
rsa1.AddNew
rsa1.Fields("sifra") = Me.text1.Text
rsa1.Fields("naziv") = Me.Text2.Text
rsa1.Fields("kol") = Me.Text3.Text
'rsa1.Fields("pop") = Me.pop.Text
rsa1.Fields("stdok") = Pblagajna
rsa1.Fields("cena") = Getnazi("select madampcd from mada where madasifr='" & Trim(Me.text1.Text) & "'") '
'rsa1.Fields("mpc") = Getcena(Me.Text1.text)

rsa1.Fields("znes") = Getnazi("select madampcd from mada where madasifr='" & Trim(Me.text1.Text) & "'") * Me.Text3.Text
rsa1.Fields("x") = Getnumb("select max(x) as x  from trenutna where  stdok='" & Pblagajna & "'") + 1
rsa1.Update
osssv

Me.Text3.SetFocus
Sendkeys "+{RIGHT}"
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Me.Text3.Text = "" Then
Me.Text3.Text = 0

End If
myConection.Execute ("update trenutna set kol=" & Me.Text3.Text & " where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
myConection.Execute ("update trenutna set znes=kol*(1-(pop/100))*cena where stdok='" & Pblagajna & "' and x=" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5))
osssv
Me.Text3.Text = 1
Me.Text2.Text = ""
Me.text1.Text = ""
'Me.pop.Text = 0
Me.text1.SetFocus
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
txtInvoiceNo.Text = novast(Val(Getnazi("select max(id_dok) as iddo from nabasif where tip_dok='PA'")) + 1, 6)
If refr = 1 Then
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
refr = 0
 
End If
End Sub

Private Function Hanb(intBtn As Integer)
    trenu = intBtn
   ' Flistvel veli, "select dim from swit WHERE [command]<>1 AND [Switchboar]=" & nas1(intBtn - 1).Tag & " group by dim order by dim"
   If prvay > 0 Then
    Dim cx As Integer
    For cx = 0 To 29
    Me.nas1(cx).Visible = False
    Next
    Me.Frame1.Visible = True
    Else
    prvay = 1
   Me.Frame1.Visible = False
   End If
   
   
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
 Do While Not aad = 29
      aad = aad + 1
      Me("LaVolpeButton" & aad).Tag = ""
      Me("LaVolpeButton" & aad).Visible = False
     
      Loop
      aad = 0
      rs.MoveFirst
       While Not rs.EOF
       aad = aad + 1
       If aad <= 29 Then
       If Not IsNull(rs![ITEMTEXT]) Then
       Me("LaVolpeButton" & aad).Visible = True
         Me("LaVolpeButton" & aad).Caption = rs![ITEMTEXT]
        If Val(Getnazi("select kontrola from mada where madasifr='" & rs![ARGUMENT] & "'")) <> 0 Then
          Me("LaVolpeButton" & aad).BackColor = Val(Getnazi("select kontrola from mada where madasifr='" & rs![ARGUMENT] & "'"))
         End If
         Me("LaVolpeButton" & aad).Tag = rs![ARGUMENT]
         End If
       End If
            rs.MoveNext
        Wend
        aad = 0
      Do While Not aad = 29
      aad = aad + 1
      If Me("LaVolpeButton" & aad).Tag = "" Then
      Me("LaVolpeButton" & aad).Visible = False
      End If
      Loop
      Else
         End If
        
    ' If no item matches, report the error and exit the function.
    
    Me.Frame1.Refresh
End Function

Private Sub mnunaj_Click()
    'trenu = intBtn
   ' Flistvel veli, "select dim from swit WHERE [command]<>1 AND [Switchboar]=" & nas1(intBtn - 1).Tag & " group by dim order by dim"
   If prvay > 0 Then
    Dim cx As Integer
    For cx = 0 To 29
    Me.nas1(cx).Visible = False
    Next
    Me.Frame1.Visible = True
    Else
    prvay = 1
   Me.Frame1.Visible = False
   End If
   
   
    If rs.State = 1 Then rs.Close
   If sqlb = "" Then
   rs.Open "select top 30 sifra,sum(kol) as koli from nabasif group by sifra order by sum(kol)"
   Else
   rs.Open sqlb
   'sqlb = ""
   End If
      If rs.EOF Then
      Exit Sub
      End If
      rs.MoveFirst
      Dim aad As Integer
      aad = 0
      If Not rs.EOF Then
 Do While Not aad = 29
      aad = aad + 1
      Me("LaVolpeButton" & aad).Tag = ""
      Me("LaVolpeButton" & aad).Visible = False
     
      Loop
      aad = 0
      rs.MoveFirst
       Dim nnnn As String
       
       While Not rs.EOF
       nnnn = Getnazi("select madanazi from mada where madasifr='" & rs.Fields("sifra") & "'")
       aad = aad + 1
       
       If aad <= 29 Then
       If Not IsNull(nnnn) Then
       Me("LaVolpeButton" & aad).Visible = True
         Me("LaVolpeButton" & aad).Caption = nnnn
        If Val(Getnazi("select kontrola from mada where madasifr='" & rs.Fields("sifra") & "'")) <> 0 Then
          Me("LaVolpeButton" & aad).BackColor = Val(Getnazi("select kontrola from mada where madasifr='" & rs.Fields("sifra") & "'"))
         End If
         Me("LaVolpeButton" & aad).Tag = rs.Fields("sifra")
         End If
       End If
            rs.MoveNext
        Wend
        aad = 0
      Do While Not aad = 29
      aad = aad + 1
      If Me("LaVolpeButton" & aad).Tag = "" Then
      Me("LaVolpeButton" & aad).Visible = False
      End If
      Loop
      Else
         End If
        
    ' If no item matches, report the error and exit the function.
    
    Me.Frame1.Refresh
End Sub

Private Function Hanbt(intBt As Integer)
  ' Me.pop.Text = 0
    Me.text1.SetFocus
  
    Me.text1.Text = Me("LaVolpeButton" & intBt).Tag
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
ss = ""
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    If rs.Fields("stmize") <= 10 Then
 ss = ss & "," & rs.Fields("stmize")
       Me.mizaa(rs.Fields("stmize")).BackColor = 5609
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
 myConection.Execute ("insert into mize select sifra,kol,x as mpcd,znes as znesek," & stm & " as stmize,pop as ddva from trenutna where stdok='" & Pblagajna & "'")
 myConection.Execute ("delete from trenutna where stdok='" & Pblagajna & "'")
Indx = 1
zap = 1
osssv
idstran = 0

For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
skumi = 0
  
Me.Label11.Caption = ""
Me.LaVo1.Caption = "MIZE  "
Me.Label12.Caption = ""
If stalnaprij = 1 Then
Else
prijavljen = ""
End If
LaVo1_Click
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

myConection.Execute ("insert into trenutna select sifra,kol,stmize as doza,ddva as pop,mpcd as x,znesek as znes," & Pblagajna & " as stdok  from mize  where stmize=" & stm & " order by mpcd")
 If rs.State = 1 Then rs.Close
rs.Open "select * from trenutna where doza=" & stm, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
rs.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & RTrim(LTrim(rs.Fields("sifra"))) & "'")
'If Not rs.Fields("znes") = 0 Then
rs.Fields("cena") = Getnumb("select madampcd from mada where madasifr='" & rs.Fields("sifra") & "'")
'End If
rs.Update
rs.MoveNext

Loop
End If
osssv

myConection.Execute ("delete from mize where stmize=" & stm)
For miz = 1 To 10
mizaa(miz).Caption = miz
mizaa(miz).BackColor = 14215660

Next
mi
mnurac_Click

End If
End Function
Private Sub printracluk()

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
    
    
    pl = "GOTOVINA"
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
    
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0, True
    cPrint.pPrint txtInvoiceNo.Text, 0.9, True
    cPrint.pPrint " z dne " & Format(Date, "dd/mm/yyyy")
    '& " " & Format(Time(), "hh:mm")
    
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
    cPrint.pRightJust Format(popu, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "----------------------------------------", 0, False
    End If
    cPrint.pPrint "ZA PLACILO EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), 4, True
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
    cPrint.pPrint "----------------------------------------", 0, False
    cPrint.pPrint "Osnova  DDV        Znesek DDV   Vrednost", 0, False
    cPrint.pPrint "----------------------------------------", 0, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.7, True
    cPrint.pRightJust "20 %", tis_e, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3# * tiskdol, True
    cPrint.pRightJust Format(ddv1, "standard"), tis_c, True
    'cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.8, True
    'cPrint.pRightJust " 20 %", 2, True
    'cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3, True
    'cPrint.pRightJust Format(ddv1, "standard"), 4, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.7, True
    cPrint.pRightJust "8.5 %", 1.3, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3#, True
    cPrint.pRightJust Format(ddv2, "standard"), 4, True
    
   ' cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.8, True
   ' cPrint.pRightJust "8.5 %", 2, True
   ' cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3, True
   ' cPrint.pRightJust Format(ddv2, "standard"), 4, True
    End If
    End If
    Dim pl As String
    cPrint.pPrint
    
    pl = "GOTOVINA"
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



Private Sub veli_Click()

Hanb (trenu)
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

Public Sub zakljucc_Click(Index As Integer)
'MsgBox (Index)
Dim rsa1 As New ADODB.Recordset

rsa1.Open "select sifra,naziv,kol,format(cena,'fixed') as cena,format(znes,'fixed') as znesek,X,pop from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic

If Not rsa1.EOF Then
 'If Me.MSHFlexGrid1.Rows > 1 Then
'Me.karto.Visible = False
    Call LaVolpeButton46_Click
    'Me.pop.Text = 0
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
