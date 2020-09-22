VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form ured_izp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Urejanje izpisov"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13095
   LinkTopic       =   "Form7"
   ScaleHeight     =   10500
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   735
      Left            =   3360
      TabIndex        =   22
      Top             =   9000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Shrani"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
      MICON           =   "uredi_izp.frx":0000
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
   Begin VB.TextBox skup3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   21
      Text            =   "IDK"
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox skup2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Text            =   "IDK"
      Top             =   8880
      Width           =   1335
   End
   Begin VB.TextBox skup1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Text            =   "IDK"
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox ddv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   18
      Text            =   "IDK"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox znes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   17
      Text            =   "IDK"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox pop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Text            =   "IDK"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox cena 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Text            =   "IDK"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox me 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Text            =   "IDK"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox kol 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Text            =   "IDK"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox opis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Text            =   "IDK"
      Top             =   7800
      Width           =   2775
   End
   Begin VB.TextBox ident 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Text            =   "IDK"
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox zap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Text            =   "IDK"
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox izdel 
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Text            =   "IDK"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox dobav 
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Text            =   "IDK"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox prod 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Text            =   "IDK"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox naz_do 
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
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Text            =   "IME"
      Top             =   6240
      Width           =   3495
   End
   Begin VB.TextBox dat 
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Text            =   "IDK"
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox IDP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Text            =   "IDK"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox IDK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "IDK"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox glava3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "glava"
      Top             =   3120
      Width           =   11895
   End
   Begin VB.TextBox glava2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "glava"
      Top             =   2880
      Width           =   11895
   End
   Begin VB.TextBox glava1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "glava"
      Top             =   2640
      Width           =   11895
   End
   Begin VB.Line Line5 
      X1              =   12600
      X2              =   12600
      Y1              =   240
      Y2              =   9960
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   12600
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   9960
   End
   Begin VB.Line Line2 
      X1              =   12600
      X2              =   240
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   600
      X2              =   12360
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   600
      X2              =   12360
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   480
      X2              =   12360
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   600
      Width           =   5205
   End
End
Attribute VB_Name = "ured_izp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Image1.Picture = LoadPicture(App.path & "\gaber.jpg")
If RS.State = 1 Then RS.Close
RS.Open "Select * from izpisi where naziv='" & repor & "'", myConection, adOpenDynamic, adLockOptimistic
Dim comman As Control
Dim sse As String

    For Each comman In Me.Controls
       If TypeOf comman Is TextBox Then
    sse = "select " & comman.Name & " from izpisi where  naziv='" & repor & "'"
    comman.text = Getnazi(sse)
'    MsgBox (sse)
    End If
   Next
 End Sub

Private Sub LaVolpeButton1_Click()
    For Each comman In Me.Controls
   If TypeOf comman Is TextBox Then
    myConection.Execute ("update izpisi set " & comman.Name & "='" & comman.text & "'  where naziv='" & repor & "'")
    
'    MsgBox (sse)
    End If
   Next
Unload Me
End Sub
