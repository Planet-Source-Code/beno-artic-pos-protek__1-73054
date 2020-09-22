VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form knji 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Knjiženje"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form8"
   ScaleHeight     =   5625
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   615
      Left            =   11880
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "POTRDI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
      MICON           =   "knji.frx":0000
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Izberi vse"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "knji.frx":001C
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
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "NE izberi vseh"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "knji.frx":0038
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
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      VariousPropertyBits=   746589211
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "24156;8281"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "Courier New"
      FontHeight      =   240
      FontCharSet     =   238
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "knji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim ssqqss, pp As String
If RS.State = 1 Then RS.Close
pp = Getnazi("select dod0 from glavna where tip_dok='" & tip_dok & "' and id_dok='" & knjiz & "'")
ssqqss = " select nabasif.tip_dok, nabasif.id_dok, nabasif.DATUM, glavna.dod0,grupa.grupa,nabasif.skl" & _
       " from  (glavna INNER JOIN (nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR) ON (glavna.id_dok = nabasif.id_dok) AND (glavna.tip_dok = nabasif.tip_dok)) LEFT JOIN GRUPA ON mada.MADAGRUP = GRUPA.SIFRA " & _
       " where  nabasif.tip_dok='" & tip_dok & "' and  glavna.dod0='" & pp & "' and isnull(nabasif.poknj)" & _
       " GROUP BY grupa.grupa,nabasif.tip_dok, nabasif.id_dok, nabasif.DATUM, glavna.dod0,nabasif.skl" & _
       " order by grupa.grupa,nabasif.skl"
 'MsgBox ssqqss
RS.Open ssqqss, myConection, adOpenStatic, adLockOptimistic
'" FROM nabasif INNER JOIN glavna ON (nabasif.id_dok = glavna.id_dok) AND (nabasif.tip_dok = glavna.tip_dok) "
If Not RS.EOF Then

RS.MoveFirst

End If
Dim i
With ListBox1
If Not RS.EOF Then
RS.MoveFirst
End If
Do While Not RS.EOF
.AddItem presled(Trim(RS.Fields(0)) & Trim(RS.Fields(1)), 13) & " " & presled(Left(RS.Fields(2), 18), 20) & "   " & Trim(RS.Fields(3)) & "  " & Left(RS.Fields(4), 8) & "  " & RS.Fields(5)
RS.MoveNext
Loop
End With
End Sub

Private Sub LaVolpeButton1_Click()
Dim i
With ListBox1
For i = 0 To .ListCount - 1
If .Selected(i) = False Then
.Selected(i) = True
End If
Next
End With
End Sub

Private Sub LaVolpeButton2_Click()
Dim i
With ListBox1
For i = 0 To .ListCount - 1
If .Selected(i) = True Then
.Selected(i) = False
End If
Next
End With
End Sub

Private Sub LaVolpeButton3_Click()
Dim iidd, xidx As String
iidd = ""
xidx = ""
If frmControlMain.Combo1.text <> "" Then
With ListBox1
For i = 0 To .ListCount - 1
If .Selected(i) = True Then
If iidd = "" Then
iidd = "'" & Mid(.Column(0, i), 3, 10) & "'"
Else
iidd = iidd & ",'" & Mid(.Column(0, i), 3, 10) & "'"
End If
End If
Next
End With
xidx = Replace(iidd, "','", ",")
Dim dss, adss As String
dss = "insert into trenutna select 'XS' as tip_dok,'AAA' as id_dok,sum(kol) as kol, sifra,min(naziv) as naziv,count(sifra) as pozicija from nabasif where tip_dok='" & tip_dok & "' and id_dok in (" & iidd & ") group by sifra"
adss = "insert into dokm (atribut,tip_dok,id_dok,tekst) values ('TREN','XS','AAA'," & xidx & ")"
myConection.Execute (adss)
'MsgBox dss
ma_ured = "1"
ma_ko = 1
dtip_dok = Trim(frmControlMain.Combo1.text)
myConection.Execute (dss)
frmblag.Show
Unload Me
End If
End Sub
