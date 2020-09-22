VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form VRPO 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Pošta"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form7"
   ScaleHeight     =   5835
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   49545217
      CurrentDate     =   39559
   End
   Begin VB.TextBox id_dok 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   7680
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "SHRANI"
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
      MICON           =   "VRPO.frx":0000
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
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Text            =   "0"
      Top             =   4560
      Width           =   3255
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
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Text            =   "0"
      Top             =   3960
      Width           =   3255
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
      Height          =   1215
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   8775
   End
   Begin ProsVent.UserControl1 UserControl13 
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      Locked          =   0   'False
      polje           =   "naziv"
      ssql            =   "select * from partner"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl12 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      polje           =   "opis"
      ssql            =   "Select * from skla"
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Locked          =   0   'False
      polje           =   "tekst"
      ssql            =   "select tekst from dokm where atribut='POST' order by poz"
      TextLocked      =   0   'False
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   10200
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PREKINI"
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
      MICON           =   "VRPO.frx":001C
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
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ZNESEK  DDV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ZNESEK BREZ DDV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OPIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PARTNER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "STM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "VRSTA POŠTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "VRPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sse As String
Me.DTPicker1 = Date
If idpo = "" Then
Me.id_dok.text = Val(Getnazi("select max(id_dok) as x from posta")) + 1
Else
Me.id_dok.text = idpo
idpo = ""
End If
If RS.State = 1 Then RS.Close
RS.Open "select * from posta where id_dok='" & Me.id_dok.text & "'", myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
Me.id_dok.text = RS.Fields("id_dok")
Me.UserControl11.BoundDatax = RS.Fields("vrpo")
Me.UserControl12.BoundDatax = RS.Fields("stm")
Me.UserControl13.BoundDatax = RS.Fields("partner")

Me.Text1.text = RS.Fields("opis")
Me.Text2.text = RS.Fields("znebr")
Me.Text3.text = RS.Fields("zneddv")
Me.DTPicker1 = RS.Fields("datum")
End If
 End Sub

Private Sub LaVolpeButton1_click()
myConection.Execute ("delete from posta where id_dok='" & Me.id_dok.text & "'")
If RS.State = 1 Then RS.Close
RS.Open "select * from posta where id_dok='" & Me.id_dok.text & "'", myConection, adOpenDynamic, adLockOptimistic

RS.AddNew
RS.Fields("id_dok") = levi_pres(Trim(Me.id_dok.text), 6)
RS.Fields("vrpo") = Me.UserControl11.BoundDatax
RS.Fields("stm") = Me.UserControl12.BoundDatax
RS.Fields("partner") = Me.UserControl13.BoundDatax

RS.Fields("opis") = Me.Text1.text
RS.Fields("znebr") = Me.Text2.text
RS.Fields("zneddv") = Me.Text3.text
 RS.Fields("datum") = Me.DTPicker1.Value
RS.Update
Unload Me

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub

