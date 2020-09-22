VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Vr_po 
   Caption         =   "Vrsta poste"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   LinkTopic       =   "Form7"
   ScaleHeight     =   3165
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tekst 
      Height          =   375
      Left            =   600
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox poz 
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      MaxLength       =   30
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
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
      MICON           =   "Vr_po.frx":0000
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
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
      MICON           =   "Vr_po.frx":001C
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vrsta pošte"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Šifra"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   315
   End
End
Attribute VB_Name = "Vr_po"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sse As String
If MODIFYID = "" Then
Me.poz.text = Val(Getnazi("select max(poz) as x from dokm where atribut='POST'")) + 1
Else
Me.poz.text = MODIFYID
MODIFYID = ""
End If
  For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    If comman.Name <> "poz" Then
    sse = "select " & comman.Name & " from dokm where atribut='POST' and poz=" & Me.poz.text
    comman.text = Getnazi(sse)
    End If
    End If
 Next
End Sub

Private Sub LaVolpeButton1_Click()
myConection.Execute ("delete from dokm where atribut='POST' and poz=" & Me.poz.text)
If RS.State = 1 Then RS.Close
RS.Open "select * from dokm where atribut='POST' and poz=" & Me.poz.text, myConection, adOpenDynamic, adLockOptimistic

RS.AddNew
RS.Fields("atribut") = "POST"
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



