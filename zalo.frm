VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form zalo 
   Caption         =   "Zaloge"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form7"
   ScaleHeight     =   3090
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "OK"
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
      MICON           =   "zalo.frx":0000
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
   Begin ProsVent.UserControl1 UserControl12 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "skladisce"
      ssql            =   "select * from skla"
      TextLocked      =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Artikel"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Skladišèe"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "zalo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()

Call Print_zal(repor, Me.UserControl11.BoundDatax, Me.UserControl12.BoundDatax)
End Sub
