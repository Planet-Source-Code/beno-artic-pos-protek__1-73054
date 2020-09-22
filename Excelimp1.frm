VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Excelimp 
   Caption         =   "Uvozi Excel"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form7"
   ScaleHeight     =   9525
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7080
      TabIndex        =   35
      Text            =   "2"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7800
      TabIndex        =   32
      Text            =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6240
      TabIndex        =   30
      Text            =   "KOS"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4920
      TabIndex        =   28
      Text            =   "20"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3480
      TabIndex        =   26
      Text            =   "TRG"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kreiraj artikle"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   8
      Left            =   2520
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   7320
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   2520
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   6720
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   2520
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   6120
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   2520
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   5520
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   2520
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   4920
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   2520
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   4320
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   2520
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   3720
      Width           =   6855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   2520
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   3120
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10560
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox polja 
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      Top             =   720
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   10800
      TabIndex        =   4
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "..."
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Excelimp.frx":0000
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
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   8055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   2520
      Width           =   6855
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   8400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "UVOZI"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Excelimp.frx":001C
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
   Begin VB.Label Label16 
      Caption         =   "Zaèni pri vrstici:"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "GRUPA:"
      Height          =   255
      Left            =   6960
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "EM:"
      Height          =   255
      Left            =   5760
      TabIndex        =   31
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "DDV:"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "TIP_ART:"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Uvozni List"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "MPC"
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
      Left            =   960
      TabIndex        =   13
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "PC"
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
      Left            =   960
      TabIndex        =   12
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Marza"
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
      Left            =   960
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Rabat"
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
      Left            =   960
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Cena"
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
      Left            =   960
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Kolicina"
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
      Left            =   960
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Naziv"
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
      Left            =   960
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "EAN"
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
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Datoteka"
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
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Sifra"
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
      Left            =   960
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Excelimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'VB6 MENU - PROJECT , REFERENCES, set a refernce to:
'Microsoft Excel 10.0 Object Library


Const cMyExcel = "Excel.xls"
Dim cPath As String
Private dokume As String
Private appExcel As Excel.Application
Private Workbook As Object

Private Sub Check1_Click()
If Me.Label12.Visible = True Then
Me.Label12.Visible = False
Me.Label13.Visible = False
Me.Text2.Visible = False
Me.Text3.Visible = False
Me.Label14.Visible = False
Me.Label15.Visible = False
Me.Text4.Visible = False
Me.Text5.Visible = False
Else
Me.Label12.Visible = True
Me.Label13.Visible = True
Me.Text2.Visible = True
Me.Text3.Visible = True
Me.Label14.Visible = True
Me.Label15.Visible = True
Me.Text4.Visible = True
Me.Text5.Visible = True

End If

End Sub

Private Sub Command1_Click()
 Dim row As Integer
  Dim cel As Integer
  Dim wb As Object
  Dim ws As Worksheet
  
  Set wb = appExcel.Application.ActiveWorkbook
  Set ws = wb.ActiveSheet
   For row = 0 To 8
  Combo1(row).clear
  '  MsgBox ws.Rows.Cells(row, 1)
  Next
  'DISPLAY VALUES IN EXCEL CELLS
  Dim ccc
  ccc = ""
  For cel = 1 To 20
    'MsgBox ws.Rows.Cells(1, cel)
    If ws.Rows.Cells(1, cel) <> "" Then
        ccc = Trim(str(cel)) + " - STOLPEC"
       For row = 0 To 8
          Combo1(row).AddItem ccc
  
         Next
    End If
  Next
  
 
End Sub
Public Sub odpri(bdok As String)
dokume = bdok
Me.Show vbModal

End Sub
Private Sub LaVolpeButton1_Click()
Dim filelocation As String
CommonDialog1.DialogTitle = "Izberi si datoteko.."
CommonDialog1.Filter = "Microsoft Excel Workbooks (*.xls)*.xls"
Dim i  As Integer
Dim strName As String
CommonDialog1.ShowOpen
   
    filelocation = CommonDialog1.FileName
    Me.Text1.Text = filelocation
    
   Set appExcel = New Excel.Application
Dim sheetcount As Integer
With appExcel
    .Workbooks.Open Text1.Text
    sheetcount = .Worksheets.Count
polja.clear

    For i = 1 To sheetcount
        strName = .Worksheets(i).Name
        polja.AddItem strName
    Next i
    polja.Refresh
    End With
    Check1_Click
End Sub
Private Sub LaVolpeButton2_Click()
Dim tip, EM As String
tip = Me.Text2.Text
EM = Me.Text4.Text
Dim gr, ddd As Integer
gr = Me.Text5.Text
ddd = Me.Text3.Text
 Dim row As Integer
  Dim cel As Integer
  Dim wb As Object
  Dim ws As Worksheet
  Dim xpo As Integer
  Dim tii, idd As String
  Dim sif, ean, naz As String
  Dim kol, cen, mar, rab, pc, mpc As Double
  Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, xpx As Integer
  xpx = 0
  p1 = Val(Left(Combo1(0).Text, 2))
  If p1 > xpx Then
  xpx = p1
  End If
  p2 = Val(Left(Combo1(1).Text, 2))
  If p2 > xpx Then
  xpx = p2
  End If
  p3 = Val(Left(Combo1(2).Text, 2))
  If p3 > xpx Then
  xpx = p3
  End If
  p4 = Val(Left(Combo1(3).Text, 2))
  If p4 > xpx Then
  xpx = p4
  End If
  p5 = Val(Left(Combo1(4).Text, 2))
  If p5 > xpx Then
  xpx = p5
  End If
  p6 = Val(Left(Combo1(5).Text, 2))
  If p6 > xpx Then
  xpx = p6
  End If
  p7 = Val(Left(Combo1(6).Text, 2))
  If p7 > xpx Then
  xpx = p7
  End If
  p8 = Val(Left(Combo1(7).Text, 2))
  If p8 > xpx Then
  xpx = p8
  End If
  p9 = Val(Left(Combo1(8).Text, 2))
  If p9 > xpx Then
  xpx = p9
  End If
  
  
tii = Left(dokume, 2)
idd = Mid(dokume, 3)
  Dim rsta As New ADODB.Recordset
  If rsta.State = 1 Then rsta.Close
  rsta.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
xpo = 1
  Set wb = appExcel.Application.ActiveWorkbook
  Set ws = wb.ActiveSheet
myConection.Execute "delete from trenutna where tip_dok='" & tii & "' and id_dok='" & idd & "'"
   For row = Val(Me.Text6.Text) To 1000
   sif = ""
   naz = ""
   ean = ""
   kol = 0
   cen = 0
   mar = 0
   rab = 0
   pc = 0
   mpc = 0
    For cel = 1 To xpx
      If cel = p1 Then
      sif = Trim(ws.Rows.Cells(row, cel))
      End If
      If cel = p2 Then
      ean = Trim(ws.Rows.Cells(row, cel))
      End If
      If cel = p3 Then
      naz = Trim(ws.Rows.Cells(row, cel))
      End If
      If cel = p4 Then
      kol = ws.Rows.Cells(row, cel)
      End If
      If cel = p5 Then
      cen = ws.Rows.Cells(row, cel)
      End If
      If cel = p6 Then
      rab = ws.Rows.Cells(row, cel)
      End If
      If cel = p7 Then
      mar = ws.Rows.Cells(row, cel)
      End If
      If cel = p8 Then
      pc = ws.Rows.Cells(row, cel)
      End If
      If cel = p9 Then
      mpc = ws.Rows.Cells(row, cel)
      End If
    Next
      If Me.Label12.Visible = True Then
      If Getnazi("select madasifr from mada where madaean='" & ean & "'") = "" Then
      Dim rs As New ADODB.Recordset
      If rs.State = 1 Then rs.Close
        Dim sifr As String
    sifr = Trim(str(Val(Getnazi("SELECT MAX(val(MADASIFR)) AS CC FROM MADA")) + 1))

        rs.Open "select * from mada", myConection, adOpenDynamic, adLockOptimistic
        rs.AddNew
        rs.Fields("madasifr") = sifr
       's.Fields("postava") = post
        rs.Fields("madanazi") = naz
        'rs.Fields("madanaz1") = naziv1.Text
       ' rs.Fields("dobavit_id") = dob_ide
        rs.Fields("madaean") = ean
        rs.Fields("madaenme") = EM
        rs.Fields("madagrup") = gr
        rs.Fields("madadoza") = 1
        rs.Fields("madaminz") = 0
        rs.Fields("tip_art") = tip
        rs.Fields("madapdv") = ddd
        rs.Fields("madapd") = ddd
        rs.Fields("madanabc") = cen
        rs.Fields("madampcd") = mpc
        rs.Fields("kontrola") = 0
        rs.Fields("odure") = ""
        rs.Fields("doure") = ""
        rs.Fields("happy") = 0
        rs.Fields("madazalo") = 0
        rs.Fields("madaemba") = 1
        rs.Fields("madaZACS") = 0
        rs.Update
        End If
      End If
      If sif = "" Then
      If ean <> "" Then
      sif = Getnazi("select madasifr from mada where madaean='" & ean & "'")
      End If
      End If
If sif <> "" Then

     rsta.AddNew
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = sif
rsta.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & sif & "'")
rsta.Fields("cena") = cen
rsta.Fields("kol") = kol
rsta.Fields("znes") = cen * kol
rsta.Fields("datum") = frmblag.DTPicker1.Value
rsta.Fields("pozicija") = levi_pres(LTrim(str(xpo)), 4)
rsta.Fields("x") = mar
rsta.Fields("y") = mpc


'rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
End If
  '   MsgBox ws.Rows.Cells(row, cel)
   
    xpo = xpo + 1
    If sif = "" Then
    If ean = "" Then
      MsgBox ("Urejeno")
  frmblag.beref
  Unload Me
    Exit Sub
    End If
    End If
  Next
 
  MsgBox ("Urejeno")
  frmblag.beref
Unload Me
End Sub


Private Sub polja_LostFocus()
Command1_Click
End Sub
