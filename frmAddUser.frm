VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAddUser.frx":0000
      Left            =   1920
      List            =   "frmAddUser.frx":000A
      TabIndex        =   9
      Text            =   "1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   2
      Top             =   1755
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   900
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1335
      Width           =   2055
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
      Left            =   2280
      MouseIcon       =   "frmAddUser.frx":0014
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2640
      Width           =   855
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
      Left            =   3240
      MouseIcon       =   "frmAddUser.frx":031E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dodaj uporabnika"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2160
   End
   Begin VB.Image imgTop 
      Height          =   720
      Left            =   0
      Picture         =   "frmAddUser.frx":0628
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
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
      Left            =   795
      TabIndex        =   7
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upor. ime"
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
      Left            =   960
      TabIndex        =   6
      Top             =   930
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Geslo"
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
      Left            =   1215
      TabIndex        =   5
      Top             =   1410
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EEECE8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   120
      Top             =   720
      Width           =   4290
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo PSDERR

If RS.State = 1 Then RS.Close
RS.Open "select * from users", myConection, adOpenDynamic, adLockOptimistic



Dim upo As String
upo = novast(Val(Getnazi("select max(up) as iddo from users ")) + 1, 2)

    If Text1.text <> "" Then
        If Text2.text <> "" Then
            If Text2.text = Text3.text Then
            If Getnazi("select username1 from users where username1='" & Text1.text & "'") = "" Then
                myConection.Execute "Insert into users (USERNAME1,PASSWORD1,NIVO,up) values ('" & Text1.text & "','" _
                                & Text2.text & "'," & Combo1.text & ",'" & upo & "')"
                Text1.text = ""
                Text2.text = ""
                Text3.text = ""
                MsgBox "Nov uporabnik dodan", vbInformation
             End If
                Exit Sub
            Else
                MsgBox "Prosim pretipkaj password.   ", vbInformation, "Password"
            End If
        Else
            MsgBox "Vnesi password.   ", vbInformation, "Password"
        End If
    Else
        MsgBox "Vnesi up.ime.    ", vbInformation, "User Name"
    End If
    




Exit Sub


PSDERR:
    MsgBox "To ime je ze v uporabi.   ", vbInformation, "User Name"
    

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If

End Sub
