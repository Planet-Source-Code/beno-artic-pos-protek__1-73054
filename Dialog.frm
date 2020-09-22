VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Izjave"
   ClientHeight    =   7080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton pobrisi 
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Pobriši vsebino"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
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
      MICON           =   "Dialog.frx":0000
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
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1800
      Top             =   6480
   End
   Begin LVbuttons.LaVolpeButton bri_izj 
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "BRIŠI"
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
      MICON           =   "Dialog.frx":001C
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
   Begin LVbuttons.LaVolpeButton dod_izj 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      MICON           =   "Dialog.frx":0038
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
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   8775
   End
   Begin LVbuttons.LaVolpeButton izj 
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   2
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
      MICON           =   "Dialog.frx":0054
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   8535
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
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   1440
      Width           =   9975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "PREKLIÈI"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "POTRDI"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub bri_izj_Click()
If MsgBox("Ali zbrišem izjavo?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then
 
 myConection.Execute "delete from dokm where atribut='izja' and id_dok='" & Trim(Combo1.text) & "'"
 End If
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub dod_izj_Click()
If Me.Text2.Visible = False Then
Me.Text2.Visible = True
Else
 If Me.Text2.text <> "" Then
 If MsgBox("Ali dodam izjavo?", vbQuestion + vbYesNo + vbDefaultButton1, "Vprašaj") = vbYes Then

 SQL = "Insert into dokm (atribut,id_dok,tekst) values ('izja','" & Trim(Me.Text2.text) & "','" & Me.Text1.text & "')"
         myConection.Execute SQL
         FillCombo Combo1, "select id_dok from dokm where atribut='izja'"
         Me.Text2.Visible = False
 End If
 Else
 MsgBox "VNESI NAZIV IZJAVE ZA DODATI!"
 End If
End If
End Sub

Private Sub Form_Load()
Dim aid, ati As String
ati = Left(xid_dok, 2)
aid = Mid(xid_dok, 3)

Text1.text = Getnazi("select tekst from dokm where atribut='" & xopis & "' and id_dok='" & aid & "' and tip_dok='" & ati & "'")
FillCombo Combo1, "select id_dok from dokm where atribut='izja'"
End Sub

Private Sub izj_Click()
If Me.Combo1.text <> "" Then
Me.Text1.text = Text1 & Getnazi("select tekst from dokm where id_dok='" & Trim(Me.Combo1.text) & "'")
End If
End Sub

Private Sub OKButton_Click()
Dim aid, ati As String
ati = Left(xid_dok, 2)
aid = Mid(xid_dok, 3)
 myConection.Execute "delete from dokm where atribut='" & levi_pres(LTrim((xopis)), 4) & "' and id_dok='" & aid & "' and tip_dok='" & ati & "'"
 SQL = "Insert into dokm (atribut,id_dok,tekst,tip_dok) values ('" & levi_pres(LTrim((xopis)), 4) & "','" & aid & "','" & Me.Text1.text & "','" & ati & "')"
        myConection.Execute SQL
        izja = Val(xopis)
        Unload Me
End Sub

Private Sub pobrisi_Click()
Me.Text1.text = ""
End Sub

Private Sub Timer1_Timer()

If Combo1.text <> "" Then
bri_izj.Visible = True
Else
bri_izj.Visible = False

End If

End Sub
