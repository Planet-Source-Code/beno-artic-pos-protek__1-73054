VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   Caption         =   "Prijavi"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form4"
   ScaleHeight     =   5415
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "odjava"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form4.frx":0000
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   4680
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   735
      Left            =   6600
      TabIndex        =   1
      Top             =   4440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "NOVA PRIJAVA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form4.frx":001C
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
   Begin VB.ListBox Cmbl 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4020
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   8640
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim xxxsql As String
   xxxsql = "sELECT * from users where prijav=1"
     
     If RS.State = 1 Then RS.Close
RS.Open xxxsql, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
If Not RS.EOF Then
    Cmbl.clear
    RS.MoveFirst
    'MsgBox ("")
    Do While Not RS.EOF
    
    dolg = ""
    
        dolg = RS.Fields("username1")
        
        With RS
         
            Cmbl.AddItem dolg
            
        End With
    RS.MoveNext
    Loop
End If
End Sub

Private Sub LaVolpeButton1_click()
frmLogin.Show
End Sub

Private Sub LaVolpeButton2_Click()
 myConection.Execute "Update users set prijav=0 where username1='" & Cmbl _
                 & "'"
End Sub

Private Sub Timer1_Timer()

Call Form_Load
End Sub
