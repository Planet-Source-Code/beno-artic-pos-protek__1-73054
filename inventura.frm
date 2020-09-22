VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form inventura 
   Caption         =   "inventura"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10845
   LinkTopic       =   "Form9"
   ScaleHeight     =   8055
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "inventura.frx":0000
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
   Begin ProsVent.xcKeypad xcKeypad2 
      Height          =   3855
      Left            =   7080
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6800
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin ProsVent.UserControl2 UserControl21 
      Height          =   1095
      Left            =   3120
      Top             =   3000
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "PRERACUNAJ"
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
      MICON           =   "inventura.frx":001C
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
   Begin ProsVent.xcKeypad xcKeypad1 
      Height          =   3855
      Left            =   7320
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6800
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "NEPREGLEDANI"
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
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "inventura.frx":0038
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
      DragIcon        =   "inventura.frx":0054
      Height          =   6120
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   10795
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
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "PREGLEDANI"
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "inventura.frx":035E
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
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "VSI"
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
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "inventura.frx":037A
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
      Caption         =   "TEŽA EMB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "POPIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "inventura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function oss()
End Function

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
Dim dat_i As String
Dim dat_ii As Date
dat_ii = Date
 dat_i = RTrim(LTrim(str(Month(dat_ii)))) & "/" & RTrim(LTrim(str(Day(dat_ii)))) & "/" & RTrim(LTrim(str(Year(dat_ii))))
'MsgBox ("select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent where dat=#" & dat_i & "#")
rs.Open "select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent where dat=#" & dat_i & "#", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Me.MSHFlexGrid1.Visible = False
Me.MSHFlexGrid1.clear
  Set Me.MSHFlexGrid1.DataSource = rs
Me.MSHFlexGrid1.Refresh


AutosizeGridColumns MSHFlexGrid1, 50, 3000
Me.MSHFlexGrid1.Row = 1
Me.MSHFlexGrid1.Visible = True
'MSHFlexGrid1_Click
End If
End Sub

Private Sub Form_Resize()
'Me.MSHFlexGrid1.Left = Me.Left
Me.MSHFlexGrid1.Height = Me.Height - Me.MSHFlexGrid1.Top - (Me.MSHFlexGrid1.Top / 4)
Me.MSHFlexGrid1.Width = Me.Width
End Sub

Private Sub Label4_Click()
Me.Text2.Enabled = True
Me.Text2.SetFocus
End Sub

Private Sub LaVolpeButton1_Click()
Me.LaVolpeButton2.BackColor = 16315377
Me.LaVolpeButton3.BackColor = 16315377
Me.LaVolpeButton1.BackColor = 255


If rs.State = 1 Then rs.Close
rs.Open "select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent where preg=0", myConection, adOpenDynamic, adLockOptimistic

'rs.Open "select madasifr,madanazi,format(doza,'fixed') as doza,madaenme,otvoritev,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis from invent where PREG=0", myConection, adOpenDynamic, adLockOptimistic
Me.MSHFlexGrid1.Visible = False
Me.MSHFlexGrid1.clear
If Not rs.EOF Then

  Set Me.MSHFlexGrid1.DataSource = rs
Me.MSHFlexGrid1.Refresh


AutosizeGridColumns MSHFlexGrid1, 50, 3000
Me.MSHFlexGrid1.Row = 1
Me.MSHFlexGrid1.Visible = True
MSHFlexGrid1_Click
End If
End Sub

Private Sub LaVolpeButton2_Click()
Me.LaVolpeButton1.BackColor = 16315377
Me.LaVolpeButton3.BackColor = 16315377
Me.LaVolpeButton2.BackColor = 255


If rs.State = 1 Then rs.Close
rs.Open "select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent where preg=1", myConection, adOpenDynamic, adLockOptimistic

'rs.Open "select madasifr,madanazi,format(doza,'fixed') as doza,madaenme,otvoritev,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis from invent where PREG=1", myConection, adOpenDynamic, adLockOptimistic
Me.MSHFlexGrid1.Visible = False
Me.MSHFlexGrid1.clear
If Not rs.EOF Then
Set Me.MSHFlexGrid1.DataSource = rs
Me.MSHFlexGrid1.Refresh


AutosizeGridColumns MSHFlexGrid1, 50, 3000

Me.MSHFlexGrid1.Row = 1
Me.MSHFlexGrid1.Visible = True
MSHFlexGrid1_Click
End If
End Sub

Private Sub LaVolpeButton3_Click()
Me.LaVolpeButton2.BackColor = 16315377
Me.LaVolpeButton1.BackColor = 16315377
Me.LaVolpeButton3.BackColor = 255


If rs.State = 1 Then rs.Close
'rs.Open "select madasifr,madanazi,format(doza,'fixed') as doza,madaenme,otvoritev,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis from invent", myConection, adOpenDynamic, adLockOptimistic
rs.Open "select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent", myConection, adOpenDynamic, adLockOptimistic

Me.MSHFlexGrid1.Visible = False
Me.MSHFlexGrid1.clear
If Not rs.EOF Then

  Set Me.MSHFlexGrid1.DataSource = rs
Me.MSHFlexGrid1.Refresh


AutosizeGridColumns MSHFlexGrid1, 50, 3000
Me.MSHFlexGrid1.Row = 1
Me.MSHFlexGrid1.Visible = True
MSHFlexGrid1_Click
End If
End Sub

Private Sub LaVolpeButton4_Click()
Dim vsia As New ADODB.Recordset
vsia.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
vsia.MoveFirst
Xvs = 1
Do While Not vsia.EOF
Xvs = Xvs + 1
vsia.MoveNext
Loop
vsia.MoveFirst
Yvs = 1
Me.UserControl21.opentime
Me.UserControl21.Visible = True

Dim rsa1 As New ADODB.Recordset
Dim rsa1x As New ADODB.Recordset
 If rs.State = 1 Then rs.Close
 If obstaja("invent") Then
 'myConection.Execute ("DROP TABLE invent")
 rs.Open "select * FROM invent", myConection, adOpenDynamic, adLockOptimistic
 Else

 rs.Open "select madazacd as dat,madasifr,madanazi,format(madadoza,'fixed') as doza,madaenme,madazacd as otvoritev,tezaemb,madazalo*0 as nabava,madazalo*0 as prodaja,madazalo*0 as zaloga,madazalo*0 as POPIS,madazalo*0 as PREG,space(1) as poknj into invent from mada", myConection, adOpenDynamic, adLockOptimistic
 End If
rsa1x.Open "select * from invent", myConection, adOpenDynamic, adLockOptimistic
rsa1x.MoveFirst
Do While Not rsa1x.EOF
DoEvents
zall (rsa1x.Fields("madasifr"))
'MsgBox (rsa1x.Fields("madasifr"))
rsa1x.MoveNext
Yvs = Yvs + 1
Loop
Me.LaVolpeButton1.BackColor = &H8080FF
Me.LaVolpeButton2.BackColor = &HF8F3F1
Me.LaVolpeButton3.BackColor = &HF8F3F1

If rs.State = 1 Then rs.Close
rs.Open "select format(dat,'dd.mm.yyyy') as datum,madasifr,madanazi,format(doza,'fixed') as doza,madaenme,format(otvoritev,'dd.mm.yyyy') as otvo,tezaemb,format(nabava,'fixed') as nabava,format(prodaja,'fixed') as prodaja,format(zaloga,'fixed') as zaloga,format(POPIS,'fixed') as popis  from invent", myConection, adOpenDynamic, adLockOptimistic
Me.MSHFlexGrid1.Visible = False
Me.MSHFlexGrid1.clear
  Set Me.MSHFlexGrid1.DataSource = rs
Me.MSHFlexGrid1.Refresh


AutosizeGridColumns MSHFlexGrid1, 50, 3000
Me.MSHFlexGrid1.Row = 1
Me.MSHFlexGrid1.Visible = True
Me.UserControl21.closetime
Me.UserControl21.Visible = False
MSHFlexGrid1_Click
End Sub

Private Sub LaVolpeButton5_Click()
If Me.Text2.Enabled = True Then
myConection.Execute ("update mada set tezaemb=" & Replace(Me.Text2.Text, ",", ".") & " where madasifr='" & Me.Label1.Caption & "'")
myConection.Execute ("update invent set tezaemb=" & Replace(Me.Text2.Text, ",", ".") & " where madasifr='" & Me.Label1.Caption & "'")
Me.MSHFlexGrid1.Refresh
Me.Text2.Enabled = False

End If
End Sub

Private Sub MSHFlexGrid1_Click()
If Me.xcKeypad1.Visible = True Then
Me.xcKeypad1.Visible = False
End If
If Me.xcKeypad2.Visible = True Then
Me.xcKeypad2.Visible = False
End If

If Me.MSHFlexGrid1.FixedCols = 1 Then

Else
Me.Label1.Caption = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1)
Me.Label2.Caption = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
Me.Text2.Text = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)

End If

End Sub

Private Sub Text1_Click()

Me.xcKeypad1.Visible = True
Me.Text1.Text = ""

End Sub



Private Sub Text2_DblClick()
If Me.Label1.Caption <> "" Then
MsgBox ("")
If Me.xcKeypad1.Visible = True Then
Me.xcKeypad1.Visible = False
End If

Me.Text2.Enabled = True
Me.xcKeypad2.Visible = True
Me.Text2.Text = ""
End If
End Sub

Private Sub xcKeypad1_Key(KeyPressed As Variant)
 Select Case KeyPressed
        Case "BS"
            If Len(Me.Text1.Text) > 0 Then
                Me.Text1.Text = Left$(Me.Text1.Text, Len(Me.Text1.Text) - 1)
            End If
        Case "Clear"
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            Me.Text1.Text = ""
           Case "ZAPRI"
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            Me.xcKeypad1.Visible = False
        Case Else
            Me.Text1.Text = Me.Text1.Text & KeyPressed
    End Select
End Sub
Function zall(artik As String)
 Dim skuu, skuup, zz As Long
skuup = 0
skuu = 0
zz = 0
 

Dim rsta As New ADODB.Recordset

Dim tString  As String
Dim rstx As New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
Dim RSt1 As New ADODB.Recordset
 
RSt1.Open "select madasifr,madanazi,madazalo,madadoza,madagrup,madasest,madanabc from mada where madasifr='" & artik & "' order by madagrup,madasifr", myConection, adOpenStatic, adLockOptimistic
If Not RSt1.EOF Then
RSt1.MoveFirst
End If
Dim dat_i, dat_x As String
Dim zal_i As Double
Dim dat_ii As Date
Dim dat_ix As Date
dat_ii = Date - 30000
Dim zalo As Double
  Dim grpa As Integer
  grpa = 0
  Dim VREDZ, dozz, nab, prod As Double

  zal_i = 0
  dat_ii = Date - 30000
    dat_ix = Date
   If Getnazi("select datum from INVENT where sifra='" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K'") = "" Then
   Else
   'dat_ii = Format(Getnazi("select datum  from nabasif where sifra='" & ltrim(rtrim(rst1.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K' order by datum desc"), "dd/mm/yyyy")
   'zal_i = Getnazi("select kol  from nabasif where sifra='" & ltrim(rtrim(rst1.Fields("madasifr"))) & "' and tip_dok='IN' and poknj='K' order by datum desc")
   End If
 dat_i = RTrim(LTrim(str(Month(dat_ii)))) & "/" & RTrim(LTrim(str(Day(dat_ii)))) & "/" & RTrim(LTrim(str(Year(dat_ii))))
 dat_x = RTrim(LTrim(str(Month(dat_ix)))) & "/" & RTrim(LTrim(str(Day(dat_ix)))) & "/" & RTrim(LTrim(str(Year(dat_ix))))
  
  VREDZ = 0
    If Getnazi("select sum(kol*faktor) as xx from nabasif where sifra='" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "' and datum>#" & dat_i & "#") = "" Then
    VREZD = 0
    Else
    dozz = RSt1.Fields("madadoza")
    
    nab = (Getnumb("select sum(kol*faktor) as xx from nabasif where tip_dok='NA' and sifra='" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "' and datum>#" & dat_i & "#"))
    prod = (Getnumb("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "' and datum>#" & dat_i & "#")) * dozz
    VREDZ = FormatNumber(nab + prod, 2)
    End If
     If Getnazi("select sifra from sestavi where sifra=" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "") <> "" Then
     VREDZ = 0
     RSt1.Fields("madasest") = "D"
     Else
     RSt1.Fields("madasest") = ""
     End If
     
     If rstx.State = 1 Then rstx.Close
  rstx.Open "select * from sestavi where sifras=" & LTrim(RTrim(RSt1.Fields("madasifr"))) & "", myConection, adOpenDynamic, adLockOptimistic
  If Not rstx.EOF Then
  rstx.MoveFirst
  Do While Not rstx.EOF
   prod = prod + ((Val(Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & rstx.Fields("sifra") & "' and datum>#" & dat_i & "#")) * rstx.Fields("kol")))
   VREDZ = zal_i + VREDZ + ((Val(Getnazi("select sum(kol*faktor) as xx from nabasif where tip_dok='PA' and sifra='" & rstx.Fields("sifra") & "' and datum>#" & dat_i & "#")) * rstx.Fields("kol")))
  rstx.MoveNext
  Loop
  End If
     'MsgBox ("update invent set prodaja=" & Replace(FormatNumber(prod, 2), ",", ".") & ",nabava=" & FormatNumber(nab, 2) & ", zaloga=" & FormatNumber(VREDZ, 2) & " where madasifr='" & artik & "'")
 myConection.Execute ("update invent set dat='" & Date & "',otvoritev='" & dat_ii & "',prodaja=" & Replace((prod), ",", ".") & ",nabava=" & Replace((nab), ",", ".") & ", zaloga=" & Replace((VREDZ), ",", ".") & " where madasifr='" & artik & "'")
End Function

Private Sub xcKeypad2_Key(KeyPressed As Variant)
Select Case KeyPressed
        Case "BS"
            If Len(Me.Text2.Text) > 0 Then
                Me.Text2.Text = Left$(Me.Text2.Text, Len(Me.Text2.Text) - 1)
            End If
        Case "Clear"
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            Me.Text2.Text = ""
           Case "ZAPRI"
            'Debug.Print txtTyping.Text
            'Me.Text1.Text = txtTyping.Text
            Me.xcKeypad2.Visible = False
        Case Else
            Me.Text2.Text = Me.Text2.Text & KeyPressed
    End Select
End Sub
