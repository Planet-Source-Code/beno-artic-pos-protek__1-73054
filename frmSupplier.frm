VERSION 5.00
Begin VB.Form C_frmSupplier 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Supplier's Information"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7305
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
   ScaleHeight     =   4440
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttitle 
      Height          =   375
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtdays 
      Height          =   330
      Left            =   1440
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtlimit 
      Height          =   330
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtremarks 
      Height          =   1215
      Left            =   3720
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtcontact 
      Height          =   375
      Left            =   3720
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtphone 
      Height          =   375
      Left            =   240
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtaddress 
      Height          =   375
      Left            =   240
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   240
      MaxLength       =   25
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6285
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5325
      TabIndex        =   6
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
      TabIndex        =   17
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
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   16
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lbl_crlimit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Limit:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
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
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
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
      TabIndex        =   13
      Top             =   2400
      Width           =   510
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   1155
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers Name"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1095
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
      TabIndex        =   8
      Top             =   120
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmSupplier.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7410
   End
End
Attribute VB_Name = "C_frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
If ADDING = True Then
    
If Trim(txtname.text) <> "" Then

  Call SuppSave
Else
    MsgBox "Please Input atleast Supplier name", vbInformation
End If

Else
    Call SuppUpdate
    
End If
Call GridRefresh



End Sub
Private Sub SuppUpdate()
Dim CDb As CDbase
Dim CUpd As New CUpdate
Dim CustID As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from supplier where businessname='" & txtname.text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CUpd)
Set CDb = CUpd

CDb.TableName = "Supplier"
CDb.ClauseStatement = "Where suppliersID='" & MODIFYID & "'"

Call CUpd.FieldVal(Trim(txtname), Trim(txtaddress), Trim(txtphone), Trim(txtcontact), Trim(txttitle), Trim(txtremarks))

Call CUpd.Update("businessname", "Address", "Phone", "contact", "Position1", "Note1")

MsgBox "Record has been Updated", vbInformation

Unload Me

Set CUpd = Nothing


    Else
          
            If Rs1!suppliersid <> MODIFYID Then
            
                MsgBox "The Record was already exist", vbInformation
            Else
            Call GetNewConnection(CUpd)
            Set CDb = CUpd

            CDb.TableName = "Supplier"
                CDb.ClauseStatement = "Where suppliersID='" & MODIFYID & "'"

           Call CUpd.FieldVal(Trim(txtname), Trim(txtaddress), Trim(txtphone), Trim(txtcontact), Trim(txttitle), Trim(txtremarks))

              
            Call CUpd.Update("businessname", "Address", "Phone", "contact", "Position1", "Note1")

            MsgBox "Record has been Updated", vbInformation

                Unload Me

            Set CUpd = Nothing

          
            End If
End If

Set Rs1 = Nothing
Set DCON = Nothing
End Sub
Private Sub SuppSave()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Supplier where BusinessName='" & txtname.text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CIns)
Set CDb = CIns


CustID = CIns.AUTONUM(CDb.OpenDb, "Supplier", "SuppliersID", "VND")

CDb.TableName = "Supplier"

CIns.FieldVal CustID, CText
CIns.FieldVal txtname, CText
CIns.FieldVal txtcontact, CText
CIns.FieldVal txttitle, CText
CIns.FieldVal txtaddress, CText
CIns.FieldVal txtphone, CText
CIns.FieldVal txtremarks, CText
CIns.FieldVal 1, CNum
CIns.FieldVal 1, CNum


CIns.Insert
MsgBox "Record has been saved", vbInformation
For Each Control In C_frmSupplier
    If TypeOf Control Is TextBox Then
        Control.text = ""
    End If
Next

Set CIns = Nothing

Else
    MsgBox "The Supplier Name was exist", vbInformation, "Supplier"

End If

Set Rs1 = Nothing
Set DCON = Nothing

End Sub

Private Sub Form_Activate()
On Error Resume Next

If ADDING = False Then
    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from supplier where suppliersID='" & MODIFYID & "'")

    If Rs1.RecordCount <> 0 Then
        txtname.text = Rs1!businessname
        txtaddress.text = Rs1!Address
        txtphone.text = Rs1!Phone
        txtcontact.text = Rs1!contact
        txttitle.text = Rs1!Position1
        txtremarks.text = Rs1!note1
        
    End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    lblError.BackColor = vbWhite
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))


End Sub

Private Sub txtcontact_KeyPress(KeyAscii As Integer)
Dim sS As String

KeyAscii = Asc(UCase(Chr(KeyAscii)))

sS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑ" & Chr(vbKeySpace) & Chr(vbKeyBack)

Call offDefine(KeyAscii, txtcontact, sS)
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim sS As String

KeyAscii = Asc(UCase(Chr(KeyAscii)))

sS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑ" & Chr(vbKeySpace) & Chr(vbKeyBack)

Call offDefine(KeyAscii, txtname, sS)


End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
Dim sS As String
sS = "1234567890-#"
Call offDefine(KeyAscii, txtphone, sS)


End Sub

Private Sub txttitle_KeyPress(KeyAscii As Integer)
Dim sS As String

KeyAscii = Asc(UCase(Chr(KeyAscii)))

sS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑ" & Chr(vbKeySpace) & Chr(vbKeyBack)

Call offDefine(KeyAscii, txttitle, sS)
End Sub
