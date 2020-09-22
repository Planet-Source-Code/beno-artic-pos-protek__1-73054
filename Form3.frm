VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form3"
   ScaleHeight     =   3795
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   180
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   375
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing... Please wait"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   7
         Top             =   360
         Width           =   3405
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "DELNI ZAKLJUCEK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
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
      MICON           =   "Form3.frx":0000
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
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """€""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "ZAKLJUCEK DNEVA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
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
      MICON           =   "Form3.frx":001C
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "KON.STANJE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ZAC.STANJE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_KeyDown(KeyCode As Integer, _
     Shift As Integer)
      
    intCtrlDown = Shift
     If Shift = 2 Then
     LaVolpeButton1.Caption = "ODKNJIŽI"
   
     End If
End Sub
Private Sub LaVolpeButton1_KeyUP(KeyCode As Integer, _
     Shift As Integer)

    intCtrlDown = 0
If Shift <> 2 Then
     LaVolpeButton1.Caption = "DELNI ZAKLJUCEK"
    
     End If

End Sub
Private Sub Form_Load()

Me.Label3.Caption = UPORABNIK
Me.Text1.text = Format(Me.Text1.text, "0.00")
If Me.Text1.text = "" Then
Me.Text1.text = "0"
End If
If Me.Text2.text = "" Then
Me.Text2.text = "0"
End If
If Getnazi("select id_dok from nabasif  where tip_dok='PA' and isnull(poknj)") = "" Then
Me.LaVolpeButton2.Enabled = False
End If

End Sub

Private Sub LaVolpeButton1_Click()
If intCtrlDown = 2 Then
Dim bb
Dim aaa As String
bb = Format(Date, "dd/mm/yyyy")
aaa = Replace(bb, ".", "/")

myConection.Execute ("update nabasif set poknj=null where tip_dok='PA' and datum=#" & aaa & "#")
Else
If Me.Text1.text = "" Then
Me.Text1.text = "0"
End If
If Me.Text2.text = "" Then
Me.Text2.text = "0"
End If

Call delnizak
Unload Me
End If
End Sub

Private Sub LaVolpeButton2_Click()
Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
   ' frmPrinterSetUp.Show vbModal
   ' If QuitCommand Then
   '     Set cPrint = Nothing
   '     Exit Sub
   ' End If

    
SendToPrinter:
   picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
    Dim datex As String
    
    datex = RTrim(LTrim(Getnazi("select datum from nabasif where tip_dok='PA' and isnull(poknj)")))
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Rekapitulacija za dan:" & Format(Date, "dd/mm/yyyy") & " "
        cPrint.pPrint "", 0.1, False
   ' cPrint.pPrint "Zaposlen:", 0.1, True
    
   ' cPrint.pPrint Me.Label3.Caption, 1, True
    If rs.State = 1 Then rs.Close
   
 Dim das, des

rs.Open "select znes,sifra,sifrapart,placilo from nabasif  where  isnull(poknj) and tip_dok='PA'", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If

Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim orr As Double
Dim hrana As Double
Dim pijaca As Double
Dim cig As Double
Dim vsto As Double

Dim kart As Double
Dim gotov As Double
gotov = 0
kart = 0
zne = 0
ddva = 0

ddvb = 0
hrana = 0
pijaca = 0
cig = 0
storitve = 0
vsto = 0
Dim davek As Double
Dim vrsta As Integer
storitve = Getnumb("SELECT  Sum(nabasif.ZNES) AS vv FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR WHERE isnull(poknj) and tip_dok='PA' and (((mada.tip_art)='STO'))")
Do While Not rs.EOF

If rs.Fields("sifrapart") <> 0 Then
orr = orr + rs.Fields(0)
End If
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

If IsNull(rs.Fields("sifra")) Then
rs.Fields("sifra") = 0
rs.Update
End If
vrsta = Getnumb("select madagrup from mada where madasifr='" & rs.Fields("sifra") & "'")
If Getnumb("select vr from grupa where sifra=" & vrsta) = 0 Then
pijaca = pijaca + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 1 Then
hrana = hrana + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 2 Then
cig = cig + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 3 Then
'storitve = storitve + rs.Fields(0)
End If
If Getnumb("select vr from grupa where sifra=" & vrsta) = 4 Then
vsto = vsto + rs.Fields(0)
End If

'storitve = Getnumb("SELECT  Sum(nabasif.ZNES) AS vv FROM nabasif LEFT JOIN mada ON nabasif.SIFRA = mada.MADASIFR WHERE isnull(poknj) and tip_dok='PA' and (((mada.tip_art)='STO'))")


zne = zne + rs.Fields(0)
If Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'") = "20" Then
ddva = ddva + rs.Fields(0)
End If
If Replace(Getnazi("select madapd from mada where madasifr='" & (rs.Fields("sifra")) & "'"), ",", ".") = "8.5" Then
ddvb = ddvb + rs.Fields(0)
End If

rs.MoveNext
Loop
Dim aa As Double
Dim bb As Double
'aa = Format(Me.Text1.text, "0.00")
'bb = Format(Me.Text2.text, "0.00")

If rs.State = 1 Then rs.Close
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where  isnull(poknj) and tip_dok='PA'", myConection, adOpenStatic, adLockOptimistic
cPrint.pPrint "", 0.1, False
Dim ee As Long

Dim ff As Long
ee = 0
ff = 0

If Not rs.EOF Then

If IsNull(rs.Fields(0)) Then
ee = 0
Else
ee = rs.Fields(0)
End If
If IsNull(rs.Fields(1)) Then

ff = 0
Else
ff = rs.Fields(1)
End If



End If
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Zacetna st.rac. : " & ee, 0.1, False
    cPrint.pPrint "Konèna st.rac.  : " & ff, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Skupaj izdano raèunov : " & ff - ee + 1, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Skupaj znesek prodaje : " & FormatNumber(zne, 2), 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
'   If pijaca <> 0 Then
'   cPrint.pPrint "Skupaj znesek pijace : " & pijaca, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If hrana <> 0 Then
'   cPrint.pPrint "Skupaj znesek hrane : " & hrana, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
'   If cig <> 0 Then
'   cPrint.pPrint "Skupaj znesek cigaretov : " & cig, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False
'   End If
 If storitve <> 0 Then
   cPrint.pPrint "Skupaj znesek storitev : " & storitve, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   
   If vsto <> 0 Then
  ' cPrint.pPrint "Skupaj znesek vstopnic : " & vsto, 0.1, False
  '  cPrint.pPrint "=======================================", 0.1, False
   End If
   
   
   If orr <> 0 Then
   cPrint.pPrint "Skupaj znesek Orginalov : " & orr, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   If kart <> 0 Then
   cPrint.pPrint "Skupaj znesek kartic : " & kart, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   If gotov <> 0 Then
   cPrint.pPrint "Skupaj znesek gotovine : " & gotov, 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
   End If
   
  cPrint.pPrint
  
  
      If ddva <> 0 Or ddvb <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddva <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddva / 1.2, "standard"), tis_e, True
    cPrint.pRightJust " 20 %", tis_a, True
    cPrint.pRightJust Format(ddva - (ddva / 1.2), "standard"), tis_b, True
    cPrint.pRightJust Format(ddva, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
     If ddvb <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddvb / 1.085, "standard"), tis_e, True
    cPrint.pRightJust "8.5 %", tis_a, True
    cPrint.pRightJust Format(ddvb - (ddvb / 1.085), "standard"), tis_b, True
    cPrint.pRightJust Format(ddvb, "standard"), tis_c, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
    End If
cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    
    Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
'odrez
    
    cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
    Dim datxx As String
    Dim dattx As String
    
    dattx = Format(Getnazi("select datum from nabasif where tip_dok='PA' and isnull(poknj)"), "mm.dd.yyyy")
    datxx = Replace(dattx, ".", "/")
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CEZPO'") = "D" Then
myConection.Execute ("update nabasif set datum=#" & datxx & "# where tip_dok='PA' and isnull(poknj)")
End If
myConection.Execute ("update nabasif set poknj='K' where tip_dok='PA' and isnull(poknj)")
myConection.Execute ("update mada set nabrddv=0")
If rs.State = 1 Then rs.Close
rs.Open ("select sum(kol) as koli,sifra from nabasif where tip_dok='PA' group by sifra")
Do While Not rs.EOF
myConection.Execute ("update mada set nabrddv=" & rs.Fields("koli") & " where madasifr='" & rs.Fields("sifra") & "'")
rs.MoveNext
Loop
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Me.Text1.text = Format(Me.Text1.text, "0.00")
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LaVolpeButton1.SetFocus
End If
End Sub
Private Sub Text2_LostFocus()
Me.Text2.text = Format(Me.Text2.text, "0.00")
End Sub
Private Sub delnizak()
 Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    'tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
   ' frmPrinterSetUp.Show vbModal
   ' If QuitCommand Then
   '     Set cPrint = Nothing
   '     Exit Sub
   ' End If

    
SendToPrinter:
   picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
   ' cPrint.pPrint
   
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Rekapitulacija za dan:" & Format(Date, "dd/mm/yyyy") & " "
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Zaposlen:", 0.1, True
    
    cPrint.pPrint Me.Label3.Caption, 1, True
    If rs.State = 1 Then rs.Close
   
' MsgBox "select znes,sifra from nabasif where isnull(poknj) and tip_dok='PA' and uporabnik='" & LTrim(Getnazi("select up from users where username1='" & Me.Label3.Caption & "'")) & "'"
rs.Open "select znes,sifra,placilo from nabasif where  isnull(poknj) and tip_dok='PA' and uporabnik='" & LTrim(Getnazi("select up from users where username1='" & Me.Label3.Caption & "'")) & "'", myConection, adOpenStatic, adLockOptimistic
If Not rs.EOF Then

rs.MoveFirst
End If
Dim zne As Double
Dim ddva As Double
Dim ddvb As Double
Dim kart As Double
Dim gotov As Double
gotov = 0
kart = 0
zne = 0
ddva = 0
ddvb = 0
Dim davek As Double

Do While Not rs.EOF
If rs.Fields("placilo") = 0 Then
gotov = gotov + rs.Fields("znes")
Else

kart = kart + rs.Fields("znes")
End If

zne = zne + rs.Fields(0)
If Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'") = "20" Then
ddva = ddva + rs.Fields(0)
End If
If Replace(Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'"), ",", ".") = "8.5" Then
ddvb = ddvb + rs.Fields(0)
End If

rs.MoveNext
Loop
Dim aa As Double
Dim bb As Double
aa = Format(Me.Text1.text, "0.00")
bb = Format(Me.Text2.text, "0.00")

If rs.State = 1 Then rs.Close
   
rs.Open "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where isnull(poknj) and tip_dok='PA' and uporabnik='" & Getnazi("select up from users where username1='" & Me.Label3.Caption & "'") & "'", myConection, adOpenStatic, adLockOptimistic
cPrint.pPrint "", 0.1, False
'Dim ssss
'ssss = "select min(id_dok) as minst, max(id_dok) as maxst from nabasif where poknj='' and uporabnik='" & Getnazi("select up from users where username1='" & Me.Label3.Caption & "'") & "'"
Dim ee As Long
Dim ff As Long
'MsgBox ssss
ee = 0
ff = 0
If Not rs.EOF Then
If Not IsNull(rs.Fields("minst")) Then
ee = rs.Fields("minst")
ff = rs.Fields("maxst")
End If
End If
cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Prijava ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ee, 6) & "'"), 0.1, False
    cPrint.pPrint "Odjava  ob : " & Getnazi("select dat_k from nabasif where tip_dok='PA' and id_dok='" & novast(ff, 6) & "'"), 0.1, False

    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Zacetna st.rac. : " & ee, 0.1, False
    cPrint.pPrint "Koncna st.rac.  : " & ff, 0.1, False
'    cPrint.pPrint "=======================================", 0.1, False

'cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "", 0.1, False
  '  cPrint.pPrint "Zacetno stanje : " & aa, 0.1, False
   ' cPrint.pPrint "Konèno stanje  : " & bb, 0.1, False
    'cPrint.pPrint "=======================================", 0.1, False
 '   cPrint.pPrint "Stanje blagajne: " & bb - aa, 0.1, False
    
    cPrint.pPrint "Skupaj znesek prodaje : " & FormatNumber(zne, 2), 0.1, True
        cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
 cPrint.pPrint Getnacindan(Format(Date, "dd/mm/yyyy"), Getnazi("select up from users where username1='" & UPORABNIK & "'"))
     cPrint.pPrint "=======================================", 0.1, False
    'If zne > (bb - aa) Then
'     cPrint.pPrint "Manjka ti : " & zne - (bb - aa), 0.1, False
    ' End If
  cPrint.pPrint
  
  
  If rs.State = 1 Then rs.Close
   
rs.Open "select sifra,naziv,kol from nabasif where  isnull(poknj) and tip_dok='PA' and uporabnik='" & LTrim(Getnazi("select up from users where username1='" & Me.Label3.Caption & "'")) & "'", myConection, adOpenStatic, adLockOptimistic
'
'If Not rs.EOF Then
'rs.MoveFirst
'End If
'Dim ciga As Integer
'ciga = Getnazi("select sifra from grupa where grupa='CIGARETI'")
'ciga = 0
'Do While Not rs.EOF
'If Val(Getnazi("select madagrup from mada where madasifr='" & rs.Fields("sifra") & "'")) = ciga Then
'cPrint.pPrint rs.Fields("naziv"), 0.1, True
'cPrint.pRightJust rs.Fields("kol"), 2.5, True
  cPrint.pPrint
'End If
 ' rs.MoveNext

'Loop
  
      If ddva <> 0 Or ddvb <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddva <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddva / 1.2, "standard"), tis_e, True
    cPrint.pRightJust " 20 %", tis_a, True
    cPrint.pRightJust Format(ddva - (ddva / 1.2), "standard"), tis_b, True
    cPrint.pRightJust Format(ddva, "standard"), tis_c, True
    End If
     If ddvb <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddvb / 1.085, "standard"), tis_e, True
    cPrint.pRightJust "8.5 %", tis_a, True
    cPrint.pRightJust Format(ddvb - (ddvb / 1.085), "standard"), tis_b, True
    cPrint.pRightJust Format(ddvb, "standard"), tis_c, True
    End If
    End If
    cPrint.pPrint
    cPrint.pPrint
      cPrint.pPrint
 cPrint.pPrint Getnacindancig(Format(Date, "dd/mm/yyyy"), Getnazi("select up from users where username1='" & UPORABNIK & "'"))
        cPrint.pPrint
cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    cPrint.pPrint ""
    
        
    cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
            cPrint.pPrint "", 0.1, False
                cPrint.pPrint "", 0.1, False
  'cPrint.pPrint Chr(27) & Chr(105), 0.1, False
   
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
  'myConection.Execute "Update racusif set org=-1 where org=0 and oseba='" & Getnazi("select up from users where username1='" & Me.Label3.Caption & "'") & "'"
End Sub

Private Sub odrez()
Open "c:\be.txt" For Output As #1
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
'Print #1, ""
Print #1, Chr(27) & Chr(105)
'Print #1, Chr(7)

'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)
'Close #1
Call Shell("print /d:lpt1 c:\be.txt", 6)
   
End Sub


Private Sub Timer1_Timer()

End Sub
