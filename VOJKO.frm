VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "fm20.dll"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form VOJKO 
   Caption         =   "Uredi"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10485
   LinkTopic       =   "Form7"
   ScaleHeight     =   10080
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox novi 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7560
      TabIndex        =   8
      Text            =   "0"
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox zne 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   6
      Text            =   "0"
      Top             =   9000
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dato 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   50659329
      CurrentDate     =   40188
   End
   Begin LVbuttons.LaVolpeButton plu 
      Height          =   975
      Left            =   9360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
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
      MICON           =   "VOJKO.frx":0000
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
      Height          =   975
      Left            =   9360
      TabIndex        =   3
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
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
      MICON           =   "VOJKO.frx":001C
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
      Height          =   975
      Left            =   9480
      TabIndex        =   10
      Top             =   9000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
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
      MICON           =   "VOJKO.frx":0038
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
   Begin MSForms.ComboBox Combo1 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   720
      Width           =   6255
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "11033;873"
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
      Object.Width           =   "7055"
   End
   Begin VB.Label Label5 
      Caption         =   "Nov:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Star:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Menjam z artiklom"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "DATUM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin MSForms.ListBox ListBox1 
      Height          =   7215
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   8655
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "15266;12726"
      MatchEntry      =   0
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "VOJKO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dato_Change()
Dim dess
Me.ListBox1.clear
uda = Format(Me.dato.Value, "dd.mm.yyyy")

dess = Mid(uda, 4, 2) & "/" & Left(uda, 2) & "/" & Mid(uda, 7, 4)
 
If rs.State = 1 Then rs.Close
If TABLEExist("delna") Then
myConection.Execute (" insert into delna select sifra,min(naziv) as naziv,0 as cen,sum(kol) as kol,sum(znes) as znesek,count(sifra) as stetje,sum(kol) as nova,0 as nov_zn from nabasif where tip_dok='PA' and datum=#" & dess & "# and placilo<>9999 group by sifra")
Else
rs.Open "select sifra,min(naziv) as naziv,0 as cen,sum(kol) as kol,sum(znes) as znesek,count(sifra) as stetje,sum(kol) as nova,0 as nov_zn into delna from nabasif where tip_dok='PA' and datum=#" & dess & "# and placilo<>9999 group by sifra", myConection, adOpenStatic, adLockOptimistic
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from delna", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then

rs.MoveFirst
End If
Dim i
With ListBox1
If Not rs.EOF Then
rs.MoveFirst
End If
.clear
Do While Not rs.EOF

.AddItem presled(rs.Fields(0), 13) & " " & presled(Left(rs.Fields(1), 18), 20) & "   " & rs.Fields(3) & "   " & rs.Fields(4)
If rs.Fields("kol") <> 0 Then
rs.Fields("cen") = Round(rs.Fields("znesek") / rs.Fields("kol"), 2)
rs.Update
End If
rs.MoveNext
Loop
End With
Me.zne.Text = Getnumb("select sum(znesek) as xx from delna")
End Sub

Private Sub Form_Load()
Dim des
 'Call Ficombo(Combo1, "select madasifr,madanazi from mada")
 If rs.State = 1 Then rs.Close

rs.Open "select madasifr,madanazi from mada", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
    Me.Combo1.clear
    rs.MoveFirst
    Do While Not rs.EOF
        With rs
            Combo1.AddItem presled(.Fields(0), 10) & .Fields(1)
        End With
    rs.MoveNext
    Loop
End If
 Me.Combo1.Text = "200        Nogavice"
Me.dato.Value = frmControlMain.DATOD.Value
uda = Format(frmControlMain.DATOD.Value, "dd.mm.yyyy")

des = Mid(uda, 4, 2) & "/" & Left(uda, 2) & "/" & Mid(uda, 7, 4)
 If obstaja("delna") Then
 myConection.Execute ("DROP TABLE delna")
 End If
If rs.State = 1 Then rs.Close

rs.Open "select sifra,min(naziv) as naziv,0 as cen,sum(kol) as kol,sum(znes) as znesek,count(sifra) as stetje,sum(kol) as nova,0 as nov_zn into delna from nabasif where tip_dok='PA' and datum=#" & des & "# and placilo<>9999 group by sifra", myConection, adOpenStatic, adLockOptimistic
If rs.State = 1 Then rs.Close
rs.Open "select * from delna", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then

rs.MoveFirst
End If
Dim i
With ListBox1
If Not rs.EOF Then
rs.MoveFirst
End If
Do While Not rs.EOF
.AddItem presled(rs.Fields(0), 13) & " " & presled(Left(rs.Fields(1), 18), 20) & "   " & rs.Fields(3) & "   " & rs.Fields(4)
'rs.Fields("cen") = Round(rs.Fields("znesek") / rs.Fields("kol"), 2)
rs.Update
rs.MoveNext
Loop
End With
Me.zne.Text = Getnumb("select sum(znesek) as xx from delna")
'Me.ListBox1.SetFocus

End Sub


Private Sub LaVolpeButton1_Click()
myConection.Execute ("delete from delna where kol=nova")
If rs.State = 1 Then rs.Close
rs.Open "select * from delna", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
Dim kvl, bbb, dexl As Integer
kvl = 0
rs.MoveFirst
Dim desx, udar
udar = Format(frmControlMain.DATOD.Value, "dd.mm.yyyy")

desx = Mid(udar, 4, 2) & "/" & Left(udar, 2) & "/" & Mid(udar, 7, 4)

Dim trs As New ADODB.Recordset

Do While Not rs.EOF
kvl = rs.Fields("nova") - rs.Fields("kol")
If trs.State = 1 Then trs.Close
trs.Open "select * from nabasif where tip_dok='PA' and sifra='" & rs.Fields("sifra") & "' and datum=#" & desx & "# and placilo<>9999 order by kol", myConection, adOpenStatic, adLockOptimistic
If Not trs.EOF Then
'MsgBox ""
trs.MoveFirst
bbb = 0
Do While Not bbb >= kvl
dexl = trs.Fields("kol")
trs.Fields("sifra") = Left(Me.Combo1.Text, 10)
trs.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & Left(Me.Combo1.Text, 10) & "'")
trs.Fields("kol") = 1
trs.Fields("cena") = Getnumb("select madampcd from mada where madasifr='" & Left(Me.Combo1.Text, 10) & "'")
trs.Fields("znes") = Getnumb("select madampcd from mada where madasifr='" & Left(Me.Combo1.Text, 10) & "'")
trs.Update
trs.MoveNext
bbb = bbb + dexl
Loop
End If
rs.MoveNext
Loop
End If
Unload Me
End Sub

Private Sub min_Click()
myConection.Execute ("update delna set kol=kol-1 where sifra='" & Left(Me.ListBox1.Text, 10) & "'")
myConection.Execute ("update delna set znesek=kol*cen where sifra='" & Left(Me.ListBox1.Text, 10) & "'")
If Getnazi("select naziv from delna where sifra='" & Left(Me.Combo1.Text, 10) & "'") <> "" Then
myConection.Execute ("update delna set kol=kol+1,nova=nova+1 where sifra='" & Left(Me.Combo1.Text, 10) & "'")
myConection.Execute ("update delna set znesek=kol*cen where sifra='" & Left(Me.Combo1.Text, 10) & "'")

Else
Dim koul As String
koul = Replace(Getnazi("select madampcd from mada where madasifr='" & Left(Me.Combo1.Text, 10) & "'"), ",", ".")
'MsgBox koul
myConection.Execute ("insert into delna (sifra,naziv,kol,cen,znesek,nova) values ('" & Left(Me.Combo1.Text, 10) & "','" & Mid(Me.Combo1.Text, 10) & "',1," & koul & "," & koul & ",1)")
 
End If
Me.novi.Text = Getnumb("select sum(znesek) as zne from delna")
osvr (Left(Me.ListBox1.Text, 10))
End Sub


Private Sub plu_Click()
myConection.Execute ("update delna set kol=kol+1 where sifra='" & Left(Me.ListBox1.Text, 10) & "'")
myConection.Execute ("update delna set znesek=kol*cen where sifra='" & Left(Me.ListBox1.Text, 10) & "'")

If Getnazi("select naziv from delna where sifra='" & Left(Me.Combo1.Text, 10) & "'") <> "" Then
myConection.Execute ("update delna set kol=kol-1,nova=nova-1 where sifra='" & Left(Left(Me.Combo1.Text, 10), 10) & "'")
myConection.Execute ("update delna set znesek=kol*cen where sifra='" & Left(Left(Me.Combo1.Text, 10), 10) & "'")

End If
Me.novi.Text = Getnumb("select sum(znesek) as zne from delna")
osvr (Left(Me.ListBox1.Text, 10))
End Sub
Private Sub osvr(xsir As String)
Me.ListBox1.clear
If rs.State = 1 Then rs.Close

If rs.State = 1 Then rs.Close
rs.Open "select * from delna", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then

rs.MoveFirst
End If
Dim i As Integer
With ListBox1
If Not rs.EOF Then
rs.MoveFirst
End If
i = 1
Do While Not rs.EOF
.AddItem presled(rs.Fields(0), 13) & " " & presled(Left(rs.Fields(1), 18), 20) & "   " & rs.Fields(3) & "   " & rs.Fields(4)
If rs.Fields(0) = LTrim(RTrim(xsir)) Then
.Selected(i - 1) = True
End If
rs.MoveNext
i = i + 1
Loop
End With
Me.Combo1.Enabled = False
End Sub
