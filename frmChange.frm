VERSION 5.00
Begin VB.Form frmChange 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pos"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   FillColor       =   &H00DC705C&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "l"
      TabIndex        =   12
      Top             =   1100
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmChange.frx":0000
      Left            =   1800
      List            =   "frmChange.frx":000A
      TabIndex        =   10
      Text            =   "1"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Preklici"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmChange.frx":0014
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MouseIcon       =   "frmChange.frx":031E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3400
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   2010
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "l"
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lbluser 
      Caption         =   "Label6"
      Height          =   135
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   15
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   705
      TabIndex        =   11
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spremeni uporabniške nastavitve"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F9F0EB&
      Height          =   270
      Left            =   100
      TabIndex        =   9
      Top             =   240
      Width           =   4125
   End
   Begin VB.Image imgTop 
      Height          =   840
      Left            =   0
      Picture         =   "frmChange.frx":0628
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Staro geslo"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novo geslo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   7
      Top             =   2130
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uporabnik"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Potrdi geslo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F8D9CB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   3000
      Left            =   240
      Top             =   960
      Width           =   3900
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call GetNewConnection2

Set Rs1 = New Recordset

        
DCON.Execute "Update users Set username1='" & Text4.text & "'" _
                                & " where username1='" & lbluser.Caption & "'"

    If Text1.text <> "" Then
             Set Rs1 = DCON.Execute("Select * from users where username1='" & lbluser.Caption _
                 & "' And password1='" & Text1.text & "'")
        If Rs1.RecordCount > 0 Then
        
        
        If Text2.text <> "" Then
            If Text2.text = Text3.text Then
                DCON.Execute "Update users Set password1='" & Text2.text & "',nivo=" & Me.Combo1.text _
                                & " where username1='" & lbluser.Caption & "'"
                Text1.text = ""
                Text2.text = ""
                Text3.text = ""
                MsgBox "Password je spremenjen", vbInformation
                
            Else
                MsgBox "Pretipkaj password.   ", vbInformation, "Password"
            End If
        Else
            MsgBox "Vnesi password.   ", vbInformation, "Password"
        End If
        
        ElseIf Rs1.RecordCount = 0 Then
            MsgBox "Pretipkaj star password.   ", vbInformation, "Old Password"
        End If
        
    Else
        MsgBox "Vnesi star Password.    ", vbInformation, "User Name"
    End If
    



Set Rs1 = Nothing

Set DCON = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text4.text = CurUser
lbluser.Caption = CurUser
Me.Text1.text = Getnazi("select password1 from users where username1='" & CurUser & "'")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub
