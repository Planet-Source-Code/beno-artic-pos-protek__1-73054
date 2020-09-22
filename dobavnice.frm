VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form dobavnice 
   Caption         =   "dobavnice"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16260
   LinkTopic       =   "Form7"
   ScaleHeight     =   8055
   ScaleWidth      =   16260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   9480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   6735
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "NOVA STRANKA"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dobavnice.frx":0000
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
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1560
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   855
      Left            =   6600
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "SHRANI STRANKO"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dobavnice.frx":001C
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
      Left            =   10440
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "SHRANI IN IZPIŠI  DOBAVNICO"
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
      COLTYPE         =   2
      BCOL            =   8454016
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dobavnice.frx":0038
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
      Left            =   13200
      TabIndex        =   6
      Top             =   6480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "IZHOD"
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
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dobavnice.frx":0054
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
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   7200
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "FAKTURIRAJ"
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
      COLTYPE         =   2
      BCOL            =   8454143
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dobavnice.frx":0070
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
   Begin MSComDlg.CommonDialog cdd 
      Left            =   7200
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "dobavnice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ReSizeForm Me
Filist List1, "select poz,tekst from dokm where tip_dok='VV' and atribut='XDOX'"

If Getnumb("select nivo from users where username1='" & UPORABNIK & "'") = 1 Then
'Me.LaVolpeButton5.Visible = True
End If
End Sub

Private Sub LaVolpeButton1_Click()
idtipk = 5
FKeyboard.Show vbModal
Me.Text2(5).Visible = True
Me.LaVolpeButton1.Visible = False
Me.LaVolpeButton2.Visible = True
End Sub

Private Sub LaVolpeButton2_Click()
Me.Text2(5).Visible = False
Me.LaVolpeButton1.Visible = True
Me.LaVolpeButton2.Visible = False
Dim asg As Integer
asg = Getnumb("select max(poz) as ss from dokm where tip_dok='VV' and atribut='XDOX'") + 1
myConection.Execute ("insert into dokm (tip_dok,poz,tekst,atribut) values('VV'," & asg & ",'" & Trim(Me.Text2(5).Text) & "','XDOX')")
Filist List1, "select poz,tekst from dokm where tip_dok='VV' and atribut='XDOX'"

End Sub
Private Function Filist(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
        dolg = presled(Trim(str(rs.Fields(0))), 5)
       ' zalo = Str(RS.Fields(3))
        With rs
        
            Cmbl.AddItem dolg & " | " & presled(Left(Trim(.Fields(1)), 20), 20) & " |" & Trim(Getnazi("select sum(znes) as ss from dobavn where faktura='.' and stranka='" & Trim(.Fields(1)) & "'"))
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function

Private Sub LaVolpeButton3_Click()
Dim stdov As Long
stdov = Getnumb("select max(st) as xx from dobavn") + 1
  Dim dass
    Dim datumx As String
    
dass = Format(Now, "dd.mm.yyyy")
datumx = Mid(dass, 4, 2) & "/" & Left(dass, 2) & "/" & Mid(dass, 7, 4)

If Me.Text1.Text <> "" Then
cdd.Copies = 1
cdd.PrinterDefault = True
cdd.ShowPrinter
Printer.Print Text1.Text
Printer.EndDoc
End If

'MsgBox ("insert into dobavn select sifra,kol,znes,'.' as faktura,#" & datumx & "# as datum ," & stdov & " as st,'" & Trim(Mid(Me.List1.Text, 8)) & "' as stranka from trenutna where stdok='" & Pblagajna & "'")
myConection.Execute ("insert into dobavn select sifra,kol,znes,'.' as faktura,#" & datumx & "# as datum ," & stdov & " as st,'" & Trim(Mid(Me.List1.Text, 8, 20)) & "' as stranka from trenutna where stdok='" & Pblagajna & "'")
myConection.Execute ("delete from trenutna where stdok='" & Pblagajna & "'")

Unload Me
frmsalesbill.LaVolpeButton45_Click

End Sub

Private Sub LaVolpeButton4_Click()
Unload Me
End Sub

Private Sub LaVolpeButton5_Click()
Dim aaaa As String
aaaa = "insert into trenutna select sifra,sum(kol) as kol,sum(znes/kol) as cena,min(naziv) as naziv,sum(znes) as znes,'" & Pblagajna & "' as stdok  from dobavn  where stranka='" & Trim(Mid(Me.List1.Text, 8, 20)) & "' and faktura='.' group by sifra order by sifra"
'MsgBox (aaaa)
myConection.Execute (aaaa)
If rs.State = 1 Then rs.Close
rs.Open "select * from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF

rs.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & rs.Fields("sifra") & "'")
rs.Update
rs.MoveNext
Loop
jedobavnica = Trim(Mid(Me.List1.Text, 8, 20))
End If
Unload Me
frmsalesbill.osssv
'myConection.Execute ("delete from mize where stmize=" & stm)
End Sub

Private Sub List1_Click()
If Getnumb("select nivo from users where username1='" & UPORABNIK & "'") = 1 Then
Me.LaVolpeButton5.Visible = True
End If
Dim stdov As Long
 Dim dass
   
    
dass = Format(Now, "dd.mm.yyyy")
'datumx = Mid(dass, 4, 2) & "/" & Left(dass, 2) & "/" & Mid(dass, 7, 4)
stdov = Getnumb("select max(st) as xx from dobavn") + 1
Text1.Text = Getnazi("select glava1 from oblikar") & _
vbCrLf & Getnazi("select glava2 from oblikar") & _
vbCrLf & Getnazi("select glava3 from oblikar") & _
vbCrLf & Getnazi("select glava4 from oblikar") & _
vbCrLf & Getnazi("select glava5 from oblikar") & _
vbCrLf & vbCrLf & "Datum: " & dass & vbCrLf & vbCrLf
'vbCrLf & "Time:" & vbCrLf

Text1.Text = Text1.Text & "Stranka: " & vbCrLf
Text1.Text = Text1.Text & Mid(Me.List1.Text, 8, 20) & vbCrLf

Text1.Text = Text1.Text & "DOBAVNICA STEVILKA: " & Trim(str(stdov)) & vbCrLf
Text1.Text = Text1.Text & "Prodajalec:" & Getnazi("select username1 from users where up='" & Getnazi("select uporabnik from nabasif where tip_dok='PA' and id_dok='" & Left(Me.List1.Text, 6) & "'") & "'") & vbCrLf & vbCrLf
Text1.Text = Text1.Text & "========================================" & vbCrLf
Text1.Text = Text1.Text & "Naziv                   kol  pop  znesek" & vbCrLf
Text1.Text = Text1.Text & "========================================" & vbCrLf
 'myConection.Execute ("insert into mize select sifra,kol,x as mpcd,znes as znesek," & stm & " as stmize,pop as ddva from trenutna where stdok='" & Pblagajna & "'")
 'myConection.Execute ("delete from trenutna where stdok='" & Pblagajna & "'")
Dim rst As New ADODB.Recordset
rst.Open "select * from trenutna where stdok='" & Pblagajna & "'", myConection, adOpenDynamic, adLockOptimistic
If rst.EOF Then
Exit Sub
End If
rst.MoveFirst
Dim ZNESE As Double
Dim ddva As Double
Dim ddvb As Double
Dim placi  As Integer
placi = 0
ZNESE = 0
ddva = 0
ddvb = 0
Do While Not rst.EOF
Text1.Text = Text1.Text & presled(Left(rst.Fields("naziv"), 20), 20) & levi_pres(FormatNumber(rst.Fields("kol"), 2), 7) & levi_pres(FormatNumber(rst.Fields("pop"), 2), 5) & levi_pres(FormatNumber(rst.Fields("znes"), 2), 8) & vbCrLf
ZNESE = ZNESE + rst.Fields("ZNES")
If Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) = 20 Then
ddva = ddva + rst.Fields("ZNES")
End If
If Val(Getnazi("select madapd from mada where madasifr='" & rst.Fields("sifra") & "'")) = 8.5 Then
ddvb = ddvb + rst.Fields("ZNES")
End If
placi = rst.Fields("placilo")


rst.MoveNext
Loop

Text1.Text = Text1.Text & "========================================" & vbCrLf


Text1.Text = Text1.Text & "ZA PLACILO EUR " & levi_pres(FormatNumber(ZNESE, 2), 25) & vbCrLf
     If ddva <> 0 Or ddvb <> 0 Then
        Text1.Text = Text1.Text & "----------------------------------------" & vbCrLf
        Text1.Text = Text1.Text & "Osnova  DDV        Znesek DDV   Vrednost" & vbCrLf
        Text1.Text = Text1.Text & "----------------------------------------" & vbCrLf
       
        If ddva <> 0 Then
    
         Text1.Text = Text1.Text & presled(Format(ddva / 1.2, "standard"), 8) & "20 %" & levi_pres(Format(ddva - (ddva / 1.2), "standard"), 14) & levi_pres(Format(ddva, "standard"), 14) & vbCrLf
   
        End If
        If ddvb <> 0 Then
         Text1.Text = Text1.Text & presled(Format(ddvb / 1.085, "standard"), 8) & "8,5 %" & levi_pres(Format(ddvb - (ddvb / 1.085), "standard"), 14) & levi_pres(Format(ddvb, "standard"), 14) & vbCrLf
        End If
       Text1.Text = Text1.Text & "----------------------------------------" & vbCrLf
    End If
  
    Text1.Text = Text1.Text & vbCrLf & "Plaèilo : DOBAVNICA" & pl & vbCrLf
    Text1.Text = Text1.Text & vbCrLf & "Podpis:_______________ " & pl & vbCrLf
      'cPrint.pPrint " Placilo: " & plax, 0.1, False
      
      Text1.Text = Text1.Text & vbCrLf
    If Getnazi("select konec1 from oblikar") <> "" Then
    Text1.Text = Text1.Text & Getnazi("select konec1 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    Text1.Text = Text1.Text & Getnazi("select konec2 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
   Text1.Text = Text1.Text & Getnazi("select konec3 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec4 from oblikar") <> "" Then
    Text1.Text = Text1.Text & Getnazi("select konec4 from oblikar") & vbCrLf
    End If
    If Getnazi("select konec5 from oblikar") <> "" Then
     Text1.Text = Text1.Text & Getnazi("select konec5 from oblikar") & vbCrLf
    End If
     Text1.Text = Text1.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Me.LaVolpeButton3.Visible = True

End Sub
