VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form C_fmAutoCompany 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AutoCompany Information"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5370
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
   ScaleHeight     =   5025
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   4290
      TabIndex        =   11
      Top             =   1695
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   4290
      TabIndex        =   10
      Top             =   2130
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4305
      TabIndex        =   9
      Top             =   2550
      Width           =   735
   End
   Begin VB.TextBox txtnumber 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtseries 
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtyear 
      Height          =   330
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtname 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3525
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4425
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   1440
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine Model Number"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Series"
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   1440
      Width           =   435
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
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   3285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Company Name"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Auto Company Information"
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
      Top             =   240
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "C_frmAutoCompany.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5370
   End
End
Attribute VB_Name = "C_fmAutoCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AUTOID As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOk_Click()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String


Call GetNewConnection(CIns)
Set CDb = CIns


CustID = CIns.AUTONUM(CDb.OpenDb, "AutoCompany", "AutoCompanyID", "AUTO")

CDb.TableName = "autocompany"

CIns.FieldVal CustID, CText
CIns.FieldVal txtname, CText


CIns.Insert
End Sub

Private Sub Command1_Click()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String


Call GetNewConnection(CIns)
Set CDb = CIns


CustID = CIns.AUTONUM(CDb.OpenDb, "enginemodel", "EngineModelID", "ENGI")

CDb.TableName = "enginemodel"

CIns.FieldVal CustID, CText
CIns.FieldVal AUTOID, CText
CIns.FieldVal txtnumber, CText
CIns.FieldVal txtseries, CText
CIns.FieldVal txtyear, CText


CIns.Insert
End Sub

Private Sub Form_Activate()
Dim CDb As CDbase
Dim CIns As New CInsert


Call GetNewConnection(CIns)
Set CDb = CIns


AUTOID = CIns.AUTONUM(CDb.OpenDb, "AutoCompany", "AutoCompanyID", "AUTO")

End Sub

Private Sub Form_Load()
    lblError.BackColor = vbWhite
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
