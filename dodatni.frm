VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form dodatni 
   Caption         =   "Dodatni"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form7"
   ScaleHeight     =   7560
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   9960
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "DODAJ"
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
      COLTYPE         =   2
      BCOL            =   8454016
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dodatni.frx":0000
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   9960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "BRIŠI"
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
      COLTYPE         =   2
      BCOL            =   8421631
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "dodatni.frx":001C
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
End
Attribute VB_Name = "dodatni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Fili List1, "select naziv from dodatni where sifra='" & dodatni_ar & "'"
If DOD_AR = "art" Or DOD_AR = "OPOMBA" Then
Else
Me.Width = Me.List1.Width
Me.Height = Me.List1.Height
Me.LaVolpeButton1.Visible = False
End If
End Sub
Private Function Fili(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
'MsgBox (strSQl)
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
        dolg = Trim(rs.Fields(0))
        With rs
         
            Cmbl.AddItem dolg
            
        End With
    rs.MoveNext
    Loop
End If
End Function

Private Sub LaVolpeButton1_Click()
Dim aca As String
aca = UCase(InputBox("dodaj"))
myConection.Execute ("insert into dodatni (sifra,naziv) values ('" & dodatni_ar & "','" & aca & "')")
Fili List1, "select naziv from dodatni where sifra='" & dodatni_ar & "'"
End Sub

Private Sub LaVolpeButton2_Click()
myConection.Execute ("delete from dodatni where sifra='" & dodatni_ar & "' and naziv='" & Me.List1.Text & "'")
Fili List1, "select naziv from dodatni where sifra='" & dodatni_ar & "'"
End Sub

Private Sub List1_DblClick()
If Me.LaVolpeButton1.Visible = False Then
dodatni_ar = Me.List1.Text
Unload Me
End If
End Sub
