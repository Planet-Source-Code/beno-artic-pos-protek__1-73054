VERSION 5.00
Begin VB.Form frmuser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pos"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
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
      Left            =   2400
      MouseIcon       =   "frmuser.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1920
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
      Left            =   3480
      MouseIcon       =   "frmuser.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spremeni uporabnika"
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
      TabIndex        =   6
      Top             =   120
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   450
      Picture         =   "frmuser.frx":0614
      Top             =   1845
      Width           =   780
   End
   Begin VB.Image imgTop 
      Height          =   720
      Left            =   0
      Picture         =   "frmuser.frx":0C20
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trenutni uporabnik"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Geslo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label lbluser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EEECE8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   1890
      Left            =   240
      Top             =   780
      Width           =   4290
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo PSDERR
Call GetNewConnection2
Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from users where username1='" & lbluser.Caption & "'")
        If Text1.text <> "" Then
            DCON.Execute "Update users Set username1='" & Text1.text & "'" _
                                & " where username1='" & lbluser.Caption & "'"
                CurUser = Text1.text
                lbluser.Caption = CurUser
                Text1.text = ""
                MsgBox "User Name Change", vbInformation
                
                
            Else
                MsgBox "Please input your user name.   ", vbInformation, "User Name"
            End If
Set Rs1 = Nothing
Set DCON = Nothing
Exit Sub
PSDERR:
    MsgBox "Uporabnik je že prijavljen.   ", vbInformation, "User Name"
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    lbluser.Caption = CurUser
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If

End Sub
