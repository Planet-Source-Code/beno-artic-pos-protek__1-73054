VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form12 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POS "
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67633153
      CurrentDate     =   38578
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67633153
      CurrentDate     =   38578
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox cmbtitle 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pregled"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmTestPaymentx.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4530
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Suppliers Information Sheet"
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
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dobavitelj"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   705
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
      TabIndex        =   4
      Top             =   1680
      Width           =   645
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
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EEECE8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   1920
      Left            =   60
      Top             =   825
      Width           =   4155
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustID As Integer

Private Sub cmbtitle_Click()

Select Case rptState
Case "SalesRegistry"
    SQL = "Select min(sifrapart)as sifrapart,stdok,sum(nabcena) as nabcena from nabasif where sifrapart=" & cmbtitle & " group by stdok "
Case "PurchaseRegistry"
SQL = "Select * from PurchaseRegistryHeader where supplierid='" & cmbtitle & "'"

End Select
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)

If Rs1.RecordCount <> 0 Then
    CustID = Rs1.Fields(0)
End If

Set Rs1 = Nothing
Set DCON = Nothing
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Select Case rptState
Case "SalesRegistry"

SQL = "Select stdok, min(datum) as datum, min(sifrapart) as sifrapart,sum(nabcena) as nabcenas from nabasif where (sifrapart=" & CustID & ") And [datum] between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# group by stdok"
Call frmControlMain.CreateDataPage(SQL, "Pregled nabave")

Case "PurchaseRegistry"
    
SQL = "Select * from PurchaseRegistryHeader where (PurchaseRegistryID='" & CustID & "') And [date] between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "#"
Call frmControlMain.CreateDataPage(SQL, "Purchase Registry Information")

End Select

Unload Me

End Sub

Private Sub Form_Load()

Select Case rptState
Case "SalesRegistry"

Call CMB3("nabasif", "sifrapart", cmbtitle, , True)

Case "PurchaseRegistry"
    
Call CMB3("PurchaseRegistryHeader", "SupplierID", cmbtitle, , True)

End Select


End Sub
