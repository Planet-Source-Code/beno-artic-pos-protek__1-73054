VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form nar_web 
   Caption         =   "Naroèila web"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form7"
   Moveable        =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   975
      Left            =   9000
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "POTRDITEV NAROCILA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   49152
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "nar_web.frx":0000
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
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8160
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   975
      Left            =   11640
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "PREKLIC NAROCILA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   192
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "nar_web.frx":001C
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
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   8400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "ZAPRI NAROCILA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16761024
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "nar_web.frx":0038
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
Attribute VB_Name = "nar_web"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim qaha As String
Dim xni As Integer

qaha = "select id_dok,max(stdok) as tel,sum(znes) as znes from nabasif where tip_dok='NK' and isnull(poknj) and x=0 group by id_dok order by id_dok desc"
If rst.State = 1 Then rst.Close
rst.Open qaha, myConection, adOpenDynamic, adLockOptimistic
If rst.EOF Then
'MsgBox ("Ne najdem nobenega podatka ki ustreza pogoju!")
Exit Sub
End If
rst.MoveFirst
List1.clear
Dim bec As String
Do While Not rst.EOF
If Len(Trim(rst.Fields("tel"))) = 9 Then
bec = "!@@@ @@@-@@@"
Else
bec = "!@@@ @@-@@@@"
End If
List1.AddItem Left(rst.Fields("id_dok"), 6) & "  " & presled(Format$(rst.Fields("tel"), bec), 15) & " " & levi_pres(FormatNumber(rst.Fields("znes"), 2), 10)

rst.MoveNext
Loop
ReSizeForm Me
Me.Width = Me.LaVolpeButton2.Left + Me.LaVolpeButton2.Width
Me.Height = Me.LaVolpeButton2.Top + Me.LaVolpeButton2.Height
Me.Text1.Height = Me.List1.Height
Me.Left = 0

End Sub
Sub narocil()

Text1.Text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & vbCrLf
Text1.Text = Text1.Text & "NAROÈILO STEVILKA: " & Left(Me.List1.Text, 6) & vbCrLf & vbCrLf

Text1.Text = Text1.Text & "Stranka: " & vbCrLf
Text1.Text = Text1.Text & Getnazi("select dod0 from glavna where tip_dok='PA' and id_dok='" & Left(Me.List1.Text, 6) & "'") & _
vbCrLf & Getnazi("select dod1 from glavna where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'") & _
vbCrLf & Getnazi("select dod2 from glavna where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'") & _
vbCrLf & Getnazi("select dod3 from glavna where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'") & vbCrLf


Text1.Text = Text1.Text & "========================================" & vbCrLf
Text1.Text = Text1.Text & "Naziv                                   " & vbCrLf
Text1.Text = Text1.Text & "========================================" & vbCrLf

Dim rst As New ADODB.Recordset
rst.Open "select * from nabasif where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'", myConection, adOpenDynamic, adLockOptimistic
If rst.EOF Then
Exit Sub
End If
rst.MoveFirst
Dim ZNESE As Double

Do While Not rst.EOF
Text1.Text = Text1.Text & Left(rst.Fields("naziv"), 40) & vbCrLf
'& levi_pres(FormatNumber(rst.Fields("kol"), 2), 7) & levi_pres(FormatNumber(rst.Fields("pop"), 2), 5) & levi_pres(FormatNumber(rst.Fields("znes"), 2), 8) & vbCrLf
ZNESE = ZNESE + rst.Fields("ZNES")

rst.MoveNext
Loop

Text1.Text = Text1.Text & "========================================" & vbCrLf


Text1.Text = Text1.Text & "ZA PLACILO EUR " & levi_pres(FormatNumber(ZNESE, 2), 25) & vbCrLf
    
   Text1.Text = Text1.Text & Getnazi("select dod5 from glavna where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'") & vbCrLf
End Sub

Private Sub LaVolpeButton1_Click()
Dim aaaa, stnr As String
stnr = "NK" & Left(Me.List1.Text, 6)
If stnr <> "" Then
aaaa = "insert into trenutna select (tip_dok+id_dok) as kopija,sifra,naziv,cena,kol,pop, x, znes,'" & Pblagajna & "' as stdok  from nabasif  where tip_dok+id_dok='" & stnr & "' order by pozicija"
'MsgBox (aaaa)
stevnaro = stnr
myConection.Execute (aaaa)

frmsalesbill.osssv
Unload Me
End If

End Sub

Private Sub LaVolpeButton2_Click()
myConection.Execute ("update nabasif set poknj='K' where tip_dok='NK' and id_dok='" & Left(Me.List1.Text, 6) & "'")
Unload Me
End Sub

Private Sub LaVolpeButton3_Click()
Unload Me
End Sub

Private Sub List1_Click()
Me.LaVolpeButton1.Visible = True
Me.LaVolpeButton2.Visible = True

narocil
End Sub

Private Sub Text1_GotFocus()
List1.SetFocus
End Sub
