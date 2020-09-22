VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form nastavitve 
   Caption         =   "nastavitve"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6135
   LinkTopic       =   "Form7"
   ScaleHeight     =   9105
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NACA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   33
      Text            =   "N"
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox VNMA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   31
      Text            =   "N"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox PARA2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   29
      Text            =   "N"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox WEBDA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   27
      Text            =   "N"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox stalnada 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Text            =   "N"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox SPLETDA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Text            =   "N"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox BOLDDA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Text            =   "N"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox MENJA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox GEND 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Text            =   "01.01.2010"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox SKRIST 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Text            =   "N"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox ZAKPA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Text            =   "D"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox CENAPA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Text            =   "D"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox CEZPO 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Text            =   "D"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox POPPA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Text            =   "D"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox OPISPO 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Text            =   "D"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox AVTOOP 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Text            =   "D"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox PLDNI 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Vnos nacin placila avtomatsko"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   7920
      Width           =   3615
   End
   Begin VB.Label Label16 
      Caption         =   "Vnos cene + marze"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Label Label15 
      Caption         =   "IZPISEM 2X PARAGONCA?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label Label14 
      Caption         =   "WEB PICE?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   6480
      Width           =   3615
   End
   Begin VB.Label Label13 
      Caption         =   "Skrijem gumb stalna prijava?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "Spletna postaja?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label11 
      Caption         =   "Bold tipke?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label Label10 
      Caption         =   "Menjalni artikel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   4560
      Width           =   3615
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   8400
      Width           =   1095
      Caption         =   "POTRDI"
      Size            =   "1931;873"
      FontHeight      =   165
      FontCharSet     =   238
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label9 
      Caption         =   "Generalni datum?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label8 
      Caption         =   "Na PA skrijem storno pozicij?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Nivo 2 dela zakljuèek?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Prikaz cene na gumbuh Paragonca"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "Delovni èas èez polnoè?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Izpis popusta pri PA?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Prikaz  opisa ARTIKLA na POZ dokumenta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Avtomatski prikaz opisa pozicij "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Plaèilo dni za vse partnerje:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "nastavitve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text3_Change()

End Sub

Private Sub CommandButton1_Click()
Dim ssee As String

  For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    myConection.Execute ("delete from dokm where tip_dok='XX' and id_dok='" & comman.Name & "'")
  myConection.Execute ("insert into dokm (tip_dok,id_dok,tekst)  values ('XX','" & comman.Name & "','" & comman.Text & "')")

      
    End If
 Next
 Unload Me
End Sub

Private Sub Form_Load()
Dim sse As String

  For Each comman In Me.Controls
      
    If TypeOf comman Is TextBox Then
    
    sse = "select tekst from dokm where tip_dok='XX' and id_dok='" & comman.Name & "'"
    comman.Text = Getnazi(sse)
   
    End If
 Next
End Sub

