VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form skladisce 
   Caption         =   "skladisca"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4710
   LinkTopic       =   "Form7"
   ScaleHeight     =   3570
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox skladisce 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox sifra 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox opis 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
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
      MICON           =   "skladisce.frx":0000
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
      Left            =   960
      TabIndex        =   3
      Top             =   2640
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
      MICON           =   "skladisce.frx":001C
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skladisce"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Šifra"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opis"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   315
   End
End
Attribute VB_Name = "skladisce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If MODIFYID <> "" And xskladd = "" Then
xskladd = MODIFYID
End If
Dim sse As String
If xskladd = "" Then
Me.sifra.text = Val(Getnazi("select max(val(sifra)) as x from skla")) + 1
Else
Me.sifra.text = xskladd
xskladd = ""
End If
  For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    If comman.Name <> "sifra" Then
    sse = "select " & comman.Name & " from skla where sifra='" & Me.sifra.text & "'"
    comman.text = Getnazi(sse)
    End If
    End If
 Next
End Sub

Private Sub LaVolpeButton1_click()
myConection.Execute ("delete from skla where sifra ='" & Me.sifra.text & "'")
If RS.State = 1 Then RS.Close
RS.Open "select * from skla where sifra='" & Me.sifra.text & "'", myConection, adOpenDynamic, adLockOptimistic

RS.AddNew
 For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    RS.Fields(comman.Name) = comman.text
        'comman.text = Getnazi(sse)
    End If
 Next
RS.Update
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub




