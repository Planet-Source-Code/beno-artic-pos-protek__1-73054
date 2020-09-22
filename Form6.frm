VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uredi"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   9120
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H80000005&
      Height          =   60
      Left            =   11640
      ScaleHeight     =   0
      ScaleWidth      =   135
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   195
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing... Please wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   13
         Top             =   360
         Width           =   3405
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "PREGLED POBR"
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
      MICON           =   "Form6.frx":0000
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
      Height          =   375
      Left            =   10920
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9120
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   975
      Left            =   10320
      TabIndex        =   3
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "POTRDI"
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
      MICON           =   "Form6.frx":001C
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
      Left            =   10440
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form6.frx":0038
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
      Left            =   10440
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "Form6.frx":0054
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
      Height          =   7935
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   7815
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "13785;13606"
      MatchEntry      =   0
      MultiSelect     =   1
      FontName        =   "Courier New"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   10680
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Kolicina"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cena"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ARTIKEL"
      Height          =   255
      Left            =   9600
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_LostFocus()
Dim c As Double
Dim a As Integer
a = Getnazi("select madasifr from mada where madanazi='" & Me.Combo1.text & "'")
c = Getnazi("select madampcd from mada where madanazi='" & Me.Combo1.text & "'")
Me.Text2.text = 1
Me.Text1.text = Format(c, "0.00")
Me.Label4.Caption = a
End Sub

Private Sub Form_Load()
Dim des
uda = Format(frmControlMain.datod.Value, "dd.mm.yyyy")

des = Mid(uda, 4, 2) & "/" & Left(uda, 2) & "/" & Mid(uda, 7, 4)
 Call CMB1("mada", "madanazi", Combo1)
If rs.State = 1 Then rs.Close
' If dara = 1 Then
rs.Open "select id_dok,min(datum) as datum,sum(znes) as znesek,min(placilo) as placilo from nabasif where tip_dok='PA' and datum=#" & des & "# and placilo<>9999 group by id_dok", myConection, adOpenStatic, adLockOptimistic
' Me.Label6.Caption = dara
' Else
'  Me.Label6.Caption = dara
'RS.Open "select sifra,min(naziv) as naziv,sum(kol) as kol,sum(znes) as znesek from nabasif where  id_dok=" & uredira & " group by sifra", myConection, adOpenStatic, adLockOptimistic
'End If
If Not rs.EOF Then

rs.MoveFirst
End If
Dim i
With ListBox1
If Not rs.EOF Then
rs.MoveFirst
End If
Do While Not rs.EOF
.AddItem presled(rs.Fields(0), 13) & " " & presled(Left(rs.Fields(1), 18), 20) & "   " & rs.Fields(2) & "   " & rs.Fields(3)
rs.MoveNext
Loop
End With
'Me.ListBox1.SetFocus

End Sub

Private Sub LaVolpeButton1_click()
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
Dim i
Dim a As String
Dim sss As Integer
With ListBox1
sss = 0
For i = 0 To .ListCount - 1
If .Selected(i) = True Then
 a = Trim(Left(ListBox1.Column(0, i), 13))
 'MsgBox (dara)
 If rs.State = 1 Then rs.Close
rs.Open "select * from nabasif where tip_dok='PA' and id_dok='" & a & "'", myConection, adOpenStatic, adLockOptimistic
If Me.Combo1.text = "" Then
Me.Combo1.SetFocus
MsgBox ("Artikel je obvezen!")
Exit Sub
Else
myConection.Execute ("update nabasif set stdok='A' where tip_dok='PA' and id_dok='" & a & "'")
myConection.Execute ("insert into storn select * from nabasif where tip_dok='PA' and id_dok='" & a & "' order by pozicija")
rs.MoveFirst
rs.Fields("sifra") = Label4.Caption
rs.Fields("naziv") = Combo1.text

rs.Fields("kol") = Text2.text
rs.Fields("cena") = Text1.text
rs.Fields("znes") = FormatNumber(Text1.text * Text2.text, 2)
rs.Fields("pop") = 0
rs.Fields("stdok") = ""
rs.Update

End If
End If
Next
End With
myConection.Execute ("delete from nabasif where tip_dok='PA' and stdok='A'")

uredira = 0
dara = 0
Unload Me
End Sub

Private Sub LaVolpeButton4_Click()
  Dim tString  As String
  
    tString = "This" & vbCrLf & "is" & vbLf & vbCr & "a" & vbCrLf & "Test"
    
    Set cPrint = New clsMultiPgPreview
    
    'frmPrint.Show vbModal
    If QuitCommand Then
        Set cPrint = Nothing
        Exit Sub
    End If

    
SendToPrinter:
    picPrinting.Visible = True
    
    cPrint.pStartDoc
    'cPrint.pHeader "PREGLED", , False
    cPrint.FontSize = 12
    cPrint.CurrentY = 1
    
    
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint " Naziv                             kol ", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    Dim i, ass
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim ddv2 As Double
    ddv1 = 0
    ddv2 = 0
    sku = 0
    SQL = "select sifra,sum(kol) as kol from storn where stdok='A' group by sifra"
    If rs.State = 1 Then rs.Close
    rs.Open SQL, myConection, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
    rs.MoveFirst
   Do While Not rs.EOF
cPrint.pPrint "", 0.1, False
    'cPrint.pPrint Format(RS.Fields("datum"), "dd/mm/yyyy"), 0.1, True
    cPrint.pPrint Getnazi("select madanazi from mada where madasifr='" & rs.Fields("sifra")) & "'", 0.1, True
    cPrint.pRightJust rs.Fields("kol"), 3.5, True
    rs.MoveNext
   Loop
   End If
    cPrint.pPrint ""
    'cPrint.pPrint ""
   
    
   
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
   
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "", 0.1, False
     cPrint.pPrint "", 0.1, False
      cPrint.pPrint "", 0.1, False
       cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
        cPrint.pPrint "", 0.1, False
    cPrint.pPrint Chr(27), 0.1, False
    
    cPrint.pPrint
    picPrinting.Visible = False
   ' cPrint.pFooter
    cPrint.pEndDoc
     ' cPrint.SendToPrinter = True
   ' cPrint.Orientation = Printer.Orientation
      If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
 
  If MsgBox("Pobrišem bazico z saj veš èem? ", vbInformation + vbYesNo) = vbYes Then
       
        SQL = "Delete from storn Where stdok='A'"
        myConection.Execute SQL
    End If
   
End Sub

Private Sub ListBox1_Click()
Dim a As Integer
Dim xxxsql As String
a = Val(Left(ListBox1.Column(0, ListCount), 13))
  xxxsql = "sELECT sifra FROM nabasif WHERE id_dok=" & a
           Filllist List1, xxxsql
End Sub

Private Sub listbox1_ItemCheck(Item As Integer)

End Sub

Private Sub listbox1_Scroll()
Dim a As Integer
Dim xxxsql As String
a = Val(Left(ListBox1.Column(0, ListCount), 13))
  xxxsql = "sELECT sifra FROM nabasif WHERE id_dok=" & a
           Filllist List1, xxxsql
End Sub

Private Sub ListBox2_Click()

End Sub
