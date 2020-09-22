VERSION 5.00
Begin VB.Form C_frmBank 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Banking Information"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   7230
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6165
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5205
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label lblError 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   3600
      X2              =   3600
      Y1              =   840
      Y2              =   3120
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3720
      TabIndex        =   2
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banks Name"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Bank Information Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "C_frmBank.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7530
   End
End
Attribute VB_Name = "C_frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
    lblError.BackColor = vbWhite
End Sub

Private Sub lbl_crlimit_Click()
Dim cls As CDbase
Set cls = New CDbase


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
