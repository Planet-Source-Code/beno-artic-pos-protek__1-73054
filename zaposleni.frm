VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form zaposleni 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Zaposleni"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form7"
   ScaleHeight     =   6285
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox rezi 
      Height          =   375
      Left            =   8760
      MaxLength       =   50
      TabIndex        =   26
      Text            =   "1"
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox otrok 
      Height          =   375
      Left            =   7800
      MaxLength       =   50
      TabIndex        =   24
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox km 
      Height          =   375
      Left            =   6840
      MaxLength       =   50
      TabIndex        =   22
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox emso 
      Height          =   375
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox davcna 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   18
      Top             =   4440
      Width           =   1815
   End
   Begin ProsVent.UserControl1 stmm 
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      ssql            =   "select * from skla"
      polje           =   "skladisce"
      textlocked      =   0
      locked          =   0
   End
   Begin MSComCtl2.DTPicker dat_p_zap 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46071809
      CurrentDate     =   39518
   End
   Begin MSComCtl2.DTPicker dat_roj 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46071809
      CurrentDate     =   39518
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "PREKLICI"
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
      MICON           =   "zaposleni.frx":0000
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
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "SHRANI"
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
      MICON           =   "zaposleni.frx":001C
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
   Begin VB.TextBox naslov 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      MaxLength       =   200
      TabIndex        =   6
      Top             =   3000
      Width           =   8775
   End
   Begin VB.TextBox Priimek 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox ime 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox sifra 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker Dat_zap 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46071809
      CurrentDate     =   39518
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rezident"
      Height          =   195
      Left            =   8640
      TabIndex        =   27
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Št.otrok"
      Height          =   195
      Left            =   7800
      TabIndex        =   25
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KM"
      Height          =   195
      Left            =   6840
      TabIndex        =   23
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMSO"
      Height          =   195
      Left            =   2520
      TabIndex        =   21
      Top             =   4200
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Davcna"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STM"
      Height          =   195
      Left            =   7440
      TabIndex        =   17
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datum  zaposlitve pri nas"
      Height          =   195
      Left            =   4920
      TabIndex        =   15
      Top             =   2040
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datum odjave"
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datum rojstva"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Naslov"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priimek"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ime"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Šifra"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   315
   End
End
Attribute VB_Name = "zaposleni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sse As String
If zaposle = "" Then
Me.sifra.text = Val(Getnazi("select max(val(sifra)) as x from zaposleni")) + 1
Else
Me.sifra.text = zaposle
zaposle = ""
End If
  For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    If comman.Name <> "sifra" Then
    sse = "select " & comman.Name & " from zaposleni where sifra='" & Me.sifra.text & "'"
    comman.text = Getnazi(sse)
    End If
   
    End If
    
 Next
 If Not Getnazi("select dat_p_zap from zaposleni where sifra='" & Me.sifra.text & "'") = "" Then
 Me.stmm.BoundDatax = Getnazi("select stmm from zaposleni where sifra='" & Me.sifra.text & "'")
 Me.dat_p_zap.Value = Getnazi("select dat_p_zap from zaposleni where sifra='" & Me.sifra.text & "'")
    Me.Dat_zap.Value = Getnazi("select dat_zap from zaposleni where sifra='" & Me.sifra.text & "'")
    Me.dat_roj.Value = Getnazi("select dat_roj from zaposleni where sifra='" & Me.sifra.text & "'")
    Me.stmm.BoundDatax = Getnazi("select stmm from zaposleni where sifra='" & Me.sifra.text & "'")
End If
End Sub

Private Sub LaVolpeButton1_click()
myConection.Execute ("delete from zaposleni where sifra ='" & Me.sifra.text & "'")
If RS.State = 1 Then RS.Close
RS.Open "select * from zaposleni where sifra='" & Me.sifra.text & "'", myConection, adOpenDynamic, adLockOptimistic

RS.AddNew
If Me.km.text = "" Then
Me.km.text = "0"
End If
If Me.otrok.text = "" Then
Me.otrok.text = "0"
End If
 For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    RS.Fields(comman.Name) = comman.text
        'comman.text = Getnazi(sse)
    End If
   
 Next
RS.Update

myConection.Execute ("update zaposleni set dat_roj='" & Me.dat_roj.Value & "' where sifra='" & Me.sifra.text & "'")
myConection.Execute ("update zaposleni set dat_zap='" & Me.Dat_zap.Value & "' where sifra='" & Me.sifra.text & "'")
myConection.Execute ("update zaposleni set dat_p_zap='" & Me.dat_p_zap.Value & "' where sifra='" & Me.sifra.text & "'")
myConection.Execute ("update zaposleni set stmm='" & Me.stmm.BoundDatax & "' where sifra='" & Me.sifra.text & "'")

Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub

