VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form kosovnica 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Kosovnica"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form3"
   ScaleHeight     =   4335
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Tag             =   "0"
   Begin VB.TextBox cmbcat 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Tag             =   "1"
      Top             =   2640
      Width           =   1695
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Preklici"
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
      MICON           =   "kosovnica.frx":0000
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
   Begin LVbuttons.LaVolpeButton shran 
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Tag             =   "2"
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   2
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "kosovnica.frx":001C
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
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kolicina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Naziv"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sifra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kosovnica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   330
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "kosovnica.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7410
   End
End
Attribute VB_Name = "kosovnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbcat_lostfocus()
If xxre <> "" Then

Me.cmbcat = xxre
'SendKeys "{enter}"

Label1.Caption = Getnazi("select madanazi from mada where madasifr='" & LTrim(xxre)) & "'"
    Label2.Caption = Getnazi("select madaenme from mada where madasifr='" & LTrim(xxre) & "'")
Me.Text1.text = ""
Me.Text1.SetFocus
xxre = ""
End If
Label1.Caption = Getnazi("select madanazi from mada where madasifr='" & Me.cmbcat & "'")
    Label2.Caption = Getnazi("select madaenme from mada where madasifr='" & Me.cmbcat & "'")
    
End Sub


Private Sub cmbcat_KeyUp(KeyCode As Integer, Shift As Integer)


If xxre <> "" Then

Me.cmbcat = xxre
'SendKeys "{enter}"

Label1.Caption = Getnazi("select madanazi from mada where madasifr='" & xxre) & "'"
    Label2.Caption = Getnazi("select madaenme from mada where madasifr='" & xxre & "'")
Me.Text1.text = ""
Me.Text1.SetFocus
xxre = ""
End If
End Sub
Private Sub cmbcat_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")
Select Case KeyCode

 Case vbKeyA To vbKeyZ

Dim iid As String
zap = Indx
opp = Me.Top
oppa = Me.Left
'idar = Chr(KeyCode)
'   DoSQL "mada", "madasifr", "madanazi", "madanaz1"
        iskalni = cmbcat.text & Chr(KeyCode)
       pritisk = cmbcat.text & Chr(KeyCode)
      ' DoSQL = ""
       ax = DoSQL("mada", "madasifr", "madanazi", "madanaz1")
      Me.cmbcat.text = ax
      SendKeys "{TAB}"
Case Else
    End Select
End Sub

Private Sub Form_Activate()



    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from sestavi where sifra=" & siff & " and sifras=" & izbrko)

    If Rs1.RecordCount <> 0 Then
        cmbcat.text = Str(Rs1!sifras)
        Text1.text = FormatNumber(Rs1!kol, 4)
        
        
    End If
    Label1.Caption = Getnazi("select madanazi from mada where madasifr='" & izbrko & "'")
    Label2.Caption = Getnazi("select madaenme from mada where madasifr='" & izbrko & "'")
    Set Rs1 = Nothing
    Set DCON = Nothing
    
Me.cmbcat.SetFocus
    
End Sub

Private Sub shran_Click()
If Me.Text1.text = "" Then
Me.Text1.SetFocus
Else
Dim rst As ADODB.Recordset
If UREDI = 1 Then
 Set rst = myConection.Execute("delete * from sestavi where sifra=" & siff & " and sifras=" & izbrko)
  UREDI = 0
 End If
 If RS.State = 1 Then RS.Close
   
 
RS.Open "select sifra,sifras,kol from sestavi", myConection

If Me.cmbcat.text <> "" Then
RS.AddNew
    RS.Fields(0) = siff
    RS.Fields(1) = Val(Me.cmbcat.text)
    RS.Fields(2) = FormatNumber(Me.Text1.text, 4)
    RS.Update
 RS.Close
 End If
    
 
izbrko = 0
siff = 0
     
     Unload Me
End If
End Sub

Private Sub LaVolpeButton2_Click()
 
izbrko = 0
siff = 0
Unload Me
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

 Case vbKeyReturn
 Me.shran.SetFocus
 Case vbKeyEscape
 Me.cmbcat.text = ""
    Me.cmbcat.SetFocus
Case Else
    End Select
End Sub
