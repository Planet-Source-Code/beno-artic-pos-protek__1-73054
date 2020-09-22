VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form UPORABNIKFO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uporabnik"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11850
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleMode       =   0  'User
   ScaleWidth      =   12805.26
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   10440
      Top             =   9360
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10215
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   1095
      Left            =   10320
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "GOR"
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
      MICON           =   "UPORABNIK.frx":0000
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
      Height          =   1095
      Left            =   10320
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "DOL"
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
      MICON           =   "UPORABNIK.frx":001C
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   1095
      Left            =   10320
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "Izberi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
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
      MICON           =   "UPORABNIK.frx":0038
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
      Height          =   1095
      Left            =   10320
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "pocket"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
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
      MICON           =   "UPORABNIK.frx":0054
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
      BackColor       =   &H00FFC0C0&
      Caption         =   "IMAŠ NOVA WEB NAROCILA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2566;873"
      Value           =   "0"
      Caption         =   "STALNA"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   238
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "UPORABNIKFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If stalnaprij = 0 Then
stalnaprij = 1
Else
stalnaprij = 0

End If

End Sub

Private Sub Command1_Click()
'vnoscen.cene "NA000010", "   1"
Excelimp.odpri "NK000022"

End Sub

Private Sub Form_Activate()

LaVolpeButton2_Click
If stalnaprij = 0 Then
Me.CheckBox1.Value = 0
Else
Me.CheckBox1.Value = 1
End If
Me.Label1.Caption = "Blagajna številka :" & stblagg()
End Sub

Private Sub Form_Load()
Filix List1, "select up,username1 from users order by up"
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='stalnada'") = "D" Then
Me.CheckBox1.Visible = False

Else
Me.CheckBox1.Visible = True
End If
'ReSizeForm Me
End Sub
Private Function Filix(Cmbl As ListBox, strSQl As String)
Call GetNewConnection2
Dim xrs As New Recordset
'MsgBox (strSQl)
xrs.Open strSQl, DCON, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not xrs.EOF Then
    Cmbl.clear
    xrs.MoveFirst
    'MsgBox ("")
    Do While Not xrs.EOF
    dd = Len(xrs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
        dolg = presled(Trim((xrs.Fields(0))), 6)
        With xrs
         
            Cmbl.AddItem dolg & " | " & presled(Trim(.Fields(1)), 20)
            
        End With
    xrs.MoveNext
    Loop
End If
End Function

Private Sub Label1_DblClick()
Me.Label1.Caption = "Blagajna številka : " & savekirablg()
End Sub

Private Sub LaVolpeButton1_Click()
Me.List1.SetFocus

Sendkeys "{UP}"

End Sub

Private Sub LaVolpeButton2_Click()
Me.List1.SetFocus

Sendkeys "{DOWN}"



End Sub

Private Sub LaVolpeButton4_Click()
List1_DblClick
End Sub

Private Sub List1_DblClick()
prijavljen = Left(Me.List1.Text, 6)
Unload Me
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List1_DblClick
End If

End Sub

Private Sub Timer1_Timer()
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='WEBDA'") = "D" Then

    If Getnazi("select max(id_dok) from nabasif where tip_dok='NK' and isnull(poknj) and x=0") <> "" Then
        Me.Label2.Visible = True
     Else
        Me.Label2.Visible = False
    End If
 Else
 Timer1.Enabled = False
End If
End Sub
