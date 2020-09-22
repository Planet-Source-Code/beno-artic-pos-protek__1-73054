VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form C_frmEngMod 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customers Information"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6285
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5325
      TabIndex        =   19
      Top             =   3960
      Width           =   855
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
      TabIndex        =   21
      Top             =   4080
      Width           =   4365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   3600
      X2              =   3600
      Y1              =   840
      Y2              =   3960
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   9
      Left            =   1440
      TabIndex        =   4
      Tag             =   "crdays"
      Top             =   3360
      Width           =   735
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "1296;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   8
      Left            =   240
      TabIndex        =   3
      Tag             =   "credit"
      Top             =   3360
      Width           =   735
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "1296;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   7
      Left            =   2520
      TabIndex        =   5
      Tag             =   "discount"
      Top             =   3360
      Width           =   735
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "1296;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   18
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Discounts:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   17
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label lbl_crlimit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Days:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   15
      Top             =   3120
      Width           =   825
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   4
      Left            =   240
      TabIndex        =   2
      Tag             =   "phone"
      Top             =   2640
      Width           =   3255
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "5741;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Tag             =   "add"
      Top             =   1920
      Width           =   3255
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "5741;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      TabIndex        =   14
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Company"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   13
      Top             =   960
      Width           =   1065
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
      TabIndex        =   12
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Title:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3720
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label LBL_HEAD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customers Information Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   270
      Left            =   240
      TabIndex        =   10
      Top             =   90
      Width           =   3120
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "name"
      Top             =   1200
      Width           =   3255
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "5741;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Tag             =   "Nationality"
      Top             =   1200
      Width           =   3375
      VariousPropertyBits=   753944603
      MaxLength       =   50
      BorderStyle     =   1
      Size            =   "5953;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   330
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Tag             =   "Remark"
      Top             =   1920
      Width           =   3375
      VariousPropertyBits=   753944603
      BorderStyle     =   1
      Size            =   "5953;582"
      BorderColor     =   -2147483637
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Customers Information Sheet"
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
      TabIndex        =   8
      Top             =   360
      Width           =   3210
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "C_frmEngMod.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10530
   End
End
Attribute VB_Name = "C_frmEngMod"
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
