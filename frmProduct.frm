VERSION 5.00
Begin VB.Form C_frmProduct 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Information"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7635
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
   ScaleHeight     =   5100
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbcat 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox cmblocation 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtserial 
      Height          =   360
      Left            =   120
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      ForeColor       =   &H8000000E&
      Height          =   1995
      Left            =   3480
      TabIndex        =   22
      Top             =   1680
      Width           =   2565
      Begin VB.TextBox txtselling 
         Height          =   330
         Left            =   1635
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox txtcost 
         Height          =   330
         Left            =   1635
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtlevel 
         Height          =   330
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   6
         Top             =   885
         Width           =   855
      End
      Begin VB.TextBox txtqty 
         Height          =   330
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   5
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox txtstock 
         Height          =   330
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   4
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Units Cost Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Units Selling Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   26
         Top             =   1605
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reorder Quantity"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   25
         Top             =   525
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Units In Stock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   11
         Left            =   60
         TabIndex        =   24
         Top             =   165
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ReorderLevel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   23
         Top             =   885
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Applications"
      Height          =   3750
      Left            =   1320
      TabIndex        =   17
      Top             =   6600
      Width           =   3495
      Begin VB.TextBox txtengine 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000009&
         Caption         =   "Add"
         Height          =   375
         Left            =   2580
         MaskColor       =   &H80000006&
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000009&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2595
         MaskColor       =   &H80000006&
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Model Number"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   435
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Engine Models Applicable to Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2910
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   3915
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3915
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
      Left            =   -240
      TabIndex        =   16
      Top             =   3960
      Width           =   4365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   2880
      Y1              =   1680
      Y2              =   3600
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "SerialNumber"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   660
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fill The Product Information Sheet"
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
      TabIndex        =   11
      Top             =   240
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmProduct.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "C_frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CategoryID As String
Private EngID As String






Private Sub cmbengine_Click()
Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from EngineModel where EngineModelNo='" & cmbengine.text & "'")

If Rs1.RecordCount <> 0 Then
    EngID = Rs1!EngineModelID
End If


Set Rs1 = Nothing
Set DCON = Nothing

End Sub

Private Sub cmbcat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub cmblocation_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub ProdUpdate()
Dim CDb As CDbase
Dim CUpd As New CUpdate
Dim CustID As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from mada where madanazi='" & txtname.text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CUpd)
Set CDb = CUpd

CDb.TableName = "mada"
CDb.ClauseStatement = "Where madasifr=" & Val(MODIFYID)

Call CUpd.FieldVal(Trim(MADANAZI), Trim(madasifr), Trim(madagrup), Trim(madapd), Trim(madazalo), Trim(madanabc), Trim(madazacs), Trim(MADAMPCD))

     
        
Call CUpd.Update("Name", "SerialNumber", "category", "LocationID", "UnitsInStock", "ReorderQuantity", "ReorderLevel", "UnitCostPrice", "UnitSellingPrice")

MsgBox "Record has been Updated", vbInformation

Unload Me

Set CUpd = Nothing


    Else
          
            If Rs1!madagrup <> MODIFYID Then
            
                MsgBox "The Category was already exist", vbInformation
            Else
            Call GetNewConnection(CUpd)
            Set CDb = CUpd

            CDb.TableName = "mada"
                CDb.ClauseStatement = "Where madasifr=" & Val(MODIFYID)

           'Call CUpd.FieldVal(Trim(txtname), Trim(txtserial), Trim(cmbcat), Trim(cmblocation), Trim(txtstock), Trim(txtQTY), Trim(txtlevel), Trim(txtcost))
Call CUpd.FieldVal(Trim(MADANAZI), Trim(madasifr), Trim(madagrup), Trim(madapd), Trim(madazalo), Trim(madanabc), Trim(madazacs), Trim(MADAMPCD))

     
              
            Call CUpd.Update("company", "Address", "Phone", "contact", "Position1", "Note1")

            MsgBox "Record has been Updated", vbInformation

                Unload Me

            Set CUpd = Nothing

          
            End If
End If

Set Rs1 = Nothing
Set DCON = Nothing

End Sub
Private Sub ProdSave()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim ProdID As String
Dim i As Integer
Call GetNewConnection2


Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from mada where madasifr=" & Val(txtserial.text) & " Or madanazi='" & txtname.text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CIns)
Set CDb = CIns
ProdID = CIns.AUTONUM(CDb.OpenDb, "mada", "madasifr", "PRD")

CDb.TableName = "mada"

CIns.FieldVal ProdID, CText
CIns.FieldVal cmblocation, CText
CIns.FieldVal CategoryID, CText
CIns.FieldVal txtname, CText
CIns.FieldVal txtserial, CText
CIns.FieldVal "0", CText
CIns.FieldVal txtstock, CNum
CIns.FieldVal txtstock, CText
CIns.FieldVal txtlevel, CNum
CIns.FieldVal txtQTY, CNum
CIns.FieldVal txtselling, CNum
CIns.FieldVal txtcost, CNum
CIns.FieldVal False, CBoolean

CIns.Insert

MsgBox "Record has been saved", vbInformation

For Each Control In C_frmProduct
    If TypeOf Control Is TextBox Then
        Control.text = ""
    End If
Next



Set CIns = Nothing

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from grupa where grupa='" & cmbcat.text & "'")

If Rs1.RecordCount = 0 Then

    DCON.Execute "Insert into Category values('" & cmbcat.text & "')"

End If

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from grupa where grupa='" & cmblocation.text & "'")

If Rs1.RecordCount = 0 Then

    DCON.Execute "Insert into grupa(grupa) values('" & cmblocation.text & "')"

End If

'For i = 0 To List1.ListCount - 1
'
'    DCON.Execute "Insert into itemToEngine values('" & ProdID & "','" & List1.List(i) & "')"
'
'Next
List1.Clear
'cmblocation.ListIndex = -1
'cmbcat.ListIndex = -1

cmblocation.text = ""
cmbcat.text = ""

Else
    MsgBox "The Product was exist", vbInformation, "Product"
    
End If

Set Rs1 = Nothing
Set DCON = Nothing

End Sub


Private Sub cmdOk_Click()
If ADDING = True Then
If Trim(txtname.text) <> "" Then

    Call ProdSave
Else
    MsgBox "Please Input atleast product name", vbInformation
End If

Else
    Call ProdUpdate
    
End If

End Sub

Private Sub Command3_Click()
If List1.text <> "" Then
    List1.RemoveItem List1.ListIndex
End If

End Sub

Private Sub Command4_Click()
If txtengine.text <> "" Then
    List1.AddItem txtengine.text
    txtengine.text = ""
End If

End Sub

Private Sub Form_Activate()


If ADDING = False Then
    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from mada where madasifr=" & Val(MODIFYID))

    If Rs1.RecordCount <> 0 Then
        txtname.text = Rs1!MADANAZI & " "
        txtserial.text = Rs1!madasifr & " "
        cmbcat.text = Rs1!madagrup & " "
        cmblocation.text = Rs1!madapd & " "
        txtstock.text = Rs1!madazalo & " "
        txtQTY.text = Rs1!madanabc & " "
       ' txtlevel.text = Rs1!ReorderLevel & " "
       ' txtcost.text = Rs1!UnitCostPrice & " "
        txtselling.text = Rs1!MADAMPCD & " "
        
    End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
End If


    
End Sub

Private Sub Form_Load()

    lblError.BackColor = vbWhite
    
    Call CMB1("grupa", "grupa", cmbcat)
    Call CMB1("tarife", "naziv", cmblocation)
    

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub txtcost_KeyPress(KeyAscii As Integer)
Call Decimals(KeyAscii, txtcost, 2)

End Sub

Private Sub txtlevel_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtlevel)

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
Dim ss As String

KeyAscii = Asc(UCase(Chr(KeyAscii)))

ss = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑ" & Chr(vbKeySpace) & Chr(vbKeyBack)

Call offDefine(KeyAscii, txtname, ss)

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtQTY)

End Sub

Private Sub txtselling_KeyPress(KeyAscii As Integer)
Call Decimals(KeyAscii, txtselling, 2)

End Sub

Private Sub txtselling_Validate(KEEPFOCUS As Boolean)
If Val(txtselling.text) < Val(txtcost.text) Then
    MsgBox "Unit Cost Price is greater than the Selling Price", vbInformation
    KEEPFOCUS = True
End If

End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtstock)




End Sub
