VERSION 5.00
Begin VB.Form S_frmUOM 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UOM"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3660
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   3375
   End
   Begin VB.Frame frameSUOM 
      BackColor       =   &H80000009&
      Caption         =   "Simple Units Of Measure"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lbl_cust 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Formal Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   990
      End
   End
   Begin VB.Frame frameCUOM 
      BackColor       =   &H80000009&
      Caption         =   "Compound Units Of Measure"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lbl_cust 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Second Unit"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_cust 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conversion"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lbl_cust 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "First Unit"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2685
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1725
      TabIndex        =   2
      Top             =   3240
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
      TabIndex        =   4
      Top             =   2910
      Width           =   3285
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Units of Measures"
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
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   0
      Picture         =   "S_frmUOM.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3690
   End
End
Attribute VB_Name = "S_frmUOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbType_Click()
If cmbType.text = "Compound" Then
        frameCUOM.Visible = True
        frameSUOM.Visible = False
Else
        frameSUOM.Visible = True
        frameCUOM.Visible = False
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()


Exit Sub
End Sub

Private Sub Form_Load()
    lblError.BackColor = vbWhite
    cmbType.AddItem "Compound"
    cmbType.AddItem "Simple"
    cmbType.ListIndex = 1
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
