VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FAX 
   BackColor       =   &H00FFC0C0&
   Caption         =   "FAX"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11595
   LinkTopic       =   "Form7"
   ScaleHeight     =   8640
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   1320
      Width           =   3855
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
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   1800
      Width           =   11535
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      Locked          =   0   'False
      polje           =   "naziv"
      ssql            =   "select * from partner"
      TextLocked      =   0   'False
   End
   Begin LVbuttons.LaVolpeButton Zapis 
      Height          =   735
      Left            =   8280
      TabIndex        =   3
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
      MICON           =   "FAX.frx":0000
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
      Left            =   9960
      TabIndex        =   4
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
      MICON           =   "FAX.frx":001C
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Komu"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tekst"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Prejemnik"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sse As String
If idfx = "" Then
If Getnazi("select max(id_dok) as x from dokm where tip_dok='FX'") = "" Then
Me.Label2.Caption = "000001"
Else
Me.Label2.Caption = novast(Val(Getnazi("select max(id_dok) as x from dokm where tip_dok='FX'")) + 1, 6)
End If
Else
Me.Label2.Caption = idfx
idfx = ""
End If
If RS.State = 1 Then RS.Close
RS.Open "select * from dokm where tip_dok='FX' and id_dok='" & Me.Label2.Caption & "'", myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
'Me.id_dok.text = RS.Fields("id_dok")
Me.UserControl11.BoundDatax = Getnazi("select naziv from partner where sifra=" & RS.Fields("poz"))
Me.Text1.text = RS.Fields("tekst")
End If
Me.Text2.text = Getnazi("select tekst from dokm where tip_dok='FF' and id_dok='" & Me.Label2.Caption & "' and atribut='KOMU'")
End Sub

Private Sub prekin_Click()
Unload Me
End Sub

Private Sub Zapis_Click()
myConection.Execute ("delete from dokm where tip_dok='FX' and id_dok='" & Me.Label2.Caption & "'")
myConection.Execute ("delete from dokm where tip_dok='FF' and id_dok='" & Me.Label2.Caption & "' and atribut='KOMU'")

If RS.State = 1 Then RS.Close
RS.Open "select * from dokm where tip_dok='FX' and id_dok='" & Me.Label2.Caption & "'", myConection, adOpenDynamic, adLockOptimistic

RS.AddNew
RS.Fields("tip_dok") = "FX"
RS.Fields("id_dok") = Trim(Me.Label2.Caption)
RS.Fields("poz") = Val(Getnazi("select sifra from partner where naziv='" & Me.UserControl11.BoundDatax & "'"))
RS.Fields("tekst") = Me.Text1.text
RS.Fields("atribut") = Getnazi("select up from users where username1='" & UPORABNIK & "'")
RS.Update
myConection.Execute ("insert into dokm (atribut,tip_dok,id_dok,tekst) values ('KOMU','FF','" & Trim(Me.Label2.Caption) & "','" & Me.Text2.text & "')")
Unload Me
frmControlMain.osv_Click
End Sub
