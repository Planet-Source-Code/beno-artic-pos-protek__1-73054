VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSalesReturn 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "s"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   27
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   585
      Left            =   6240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7200
      Width           =   840
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   11850
      TabIndex        =   4
      Top             =   0
      Width           =   11910
      Begin VB.ComboBox cmbdate 
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbReg 
         Height          =   315
         Left            =   9345
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1590
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtQTY 
         Height          =   375
         Left            =   2910
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Height          =   375
         Left            =   4065
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox text5 
         Height          =   375
         Left            =   5625
         TabIndex        =   19
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtDis 
         Height          =   375
         Left            =   5025
         TabIndex        =   18
         Top             =   1680
         Width           =   495
      End
      Begin VB.ComboBox cmbCust 
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   405
         Left            =   7125
         TabIndex        =   1
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales RegistyID"
         Height          =   195
         Left            =   7920
         TabIndex        =   24
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   4080
         TabIndex        =   17
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales ReturnID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   7755
         TabIndex        =   16
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lipa Solid Auto Supply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Left            =   7920
         TabIndex        =   5
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Order ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   9510
         TabIndex        =   8
         Top             =   255
         Width           =   2235
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Return"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   7920
         TabIndex        =   6
         Top             =   1320
         Width           =   405
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1320
         Left            =   7755
         Shape           =   4  'Rounded Rectangle
         Top             =   765
         Width           =   3930
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   0
         Picture         =   "frmSalesReturn.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12330
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5745
         TabIndex        =   12
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   1380
         Width           =   330
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   1395
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Left            =   60
         Top             =   1290
         Width           =   6900
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblwords 
      Caption         =   "Label4"
      Height          =   615
      Left            =   480
      TabIndex        =   25
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   7170
   End
   Begin MSForms.TextBox TextAmount 
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
      Width           =   1905
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3360;661"
      BorderColor     =   -2147483647
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   8040
      TabIndex        =   13
      Top             =   7200
      Width           =   1530
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000001&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F9F0EB&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   6885
      Width           =   3690
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   7815
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3945
   End
End
Attribute VB_Name = "frmSalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustID As String
Dim SuppName As String
Dim SuppID As String
Dim PurchID As String
Dim ReturnID As String
Dim TXTLEN As Integer
Dim STRT As Integer
Dim PRILEN As Integer
Dim SalesID As String

Private Sub cmbCust_Click()
    Call CMB2("Select Distinct [date] from Salesret where Company='" & cmbCust & "'", cmbdate)
    Call CMB2("Select Distinct SalesRegistryID from Salesret where Company='" & cmbCust & "'", cmbReg)

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select customerid from salesret where company='" & cmbCust & "'")
If Rs1.RecordCount <> 0 Then
CustID = Rs1!customerid
End If
Set Rs1 = Nothing
Set DCON = Nothing

End Sub
Private Sub cmbdate_Click()
    Call CMB2("Select Distinct SalesRegistryID from salesret where company='" & cmbCust & "' And [date]=#" & cmbdate & "#", cmbReg)
End Sub
Private Sub cmbReg_Click()
 SQL = "Select * from salesret where company='" & cmbCust & "' and salesRegistryId='" & cmbReg & "'"
 Call GetNewConnection2
 Set Rs1 = New Recordset
 Set Rs1 = DCON.Execute(SQL)
   lvMain.ListItems.clear
    
While Not Rs1.EOF
    Dim LVITEM As ListItem
    Dim ProdID As String
    With lvMain
    ProdID = Rs1!productid
    
    If Rs1!quantity > 0 Then
    
    Set LVITEM = .ListItems.Add(, , Rs1!productid)
        LVITEM.SubItems(2) = Rs1!quantity
        LVITEM.SubItems(3) = Rs1!rate
        LVITEM.SubItems(1) = Rs1!Name
        LVITEM.SubItems(4) = CLng(Rs1!quantity) * CLng(Rs1!rate)
    End If
    
    End With
    Rs1.MoveNext
Wend
calculate
Set Rs1 = Nothing
Set DCON = Nothing
End Sub
'Private Sub cmdAdd_Click()
'Set LSTITEM = ListView1.FindItem(RS1!productid, lvwText, , lvwPartial)
       'If LSTITEM Is Nothing Then
'Dim LSTITEM As ListItem
'Dim CNT As Boolean
'Dim DD As Integer
'Dim findList As ListItem
'If text5.text <> "" Then
'    CNT = False
'    Call GetNewConnection2
'        Set RS1 = New Recordset
'        Set RS1 = DCON.Execute("Select * from Product where ProductID like'" & Text4 & "%' OR Name like'" & Text4 & "%'")
'If Not RS1.EOF Then
'
'
'         txtRate = RS1!unitcostprice
' With lvMain
'        TextAmount.text = ""
'
'        If .ListItems.Count <> 0 Then
'
'            For DD = 1 To .ListItems.Count
'
'                If InStr(1, .ListItems(DD).text, RS1!productid) = 1 Then
'                    If InStr(1, .ListItems(DD).SubItems(1), RS1!Name) = 1 Then
'
'                        If EDT = True Then
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = text5.text
'                        Else
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
'
'                        End If
'
'                    CNT = True
'
'                    End If
'                End If
'
'            Next
'
'         End If
'
'        If CNT = False Then
'
'         .ListItems.Add , , RS1!productid
'            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = text5.text
'
'
'             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
'
'        End If
'
'
'
'
''        If .ListItems.Count <= 0 Then
''
''            .ListItems.Add 1, , RS1!ProductID
''            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
''            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
''            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
''            .ListItems(.ListItems.Count).SubItems(4) = "1"
''            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
''
''
''        End If
'
'
'            For DD = 1 To .ListItems.Count
'                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
'            Next
'
'        ' lblunit.Caption = RS1!UnitsInStock
'
'
'       Set RS1 = DCON.Execute("Select * from Product")
'      '  Set DataGrid1.DataSource = RS1
'
'        Set RS1 = Nothing
'        Set DCON = Nothing
'
'  End With
'
'Else
'    MsgBox "Product Not Found", vbInformation, "Product"
'
'
'End If
'
'Text4.text = ""
'text5.text = ""
'txtQTY.text = ""
'txtRate.text = ""
'EDT = False
'Text4.SetFocus
'
'Else
'
'
'
'txtQTY.SetFocus
'
'
'End If


'End Sub

'Private Sub cmdAdd_GotFocus()


'Call GetNewConnection2
'
'Set RS1 = New Recordset
'SQL = "Select TOP 5 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'
'Set RS1 = DCON.Execute(SQL)
'
'
'
'
'    SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!ProductID & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'    'MsgBox SQL
'    DCON.Execute SQL
'
'
'Set RS1 = Nothing
'Set DCON = Nothing
'
'Text5.text = Val(txtqty.text) * Val(txtrate.text)
'End Sub

Private Sub cmdClear_Click()
lvMain.ListItems.clear

End Sub

Private Sub cmdAdd_Click()
Dim lst As ListItem
If lvMain.ListItems.Count <= 0 Then Exit Sub
Set lst = lvMain.FindItem(Text4, lvwText, , lvwPartial)
    If lst Is Nothing Then
        Exit Sub
    Else
        lst.EnsureVisible
        If Not (Val(lst.SubItems(2)) >= txtQTY Or Val(txtQTY) <= 0) Then Exit Sub
               lst.SubItems(2) = txtQTY
               lst.SubItems(4) = lst.SubItems(2) * lst.SubItems(3)
        lst.Selected = True
        lvMain.SetFocus
        calculate
    End If
End Sub

Private Sub cmdOk_Click()

Call SalesReturn

End Sub
Private Sub SalesReturn()
Dim CNT1 As Integer

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
   ' If CRED = True Then
   
   
        'SalesRegistryHeader
        'SalesRegistryDetail
        
'        SQL = "Insert into SalesReturnHeader values('" & SalesID & "','" _
'                                                  & 'SELECT CUSTOMERID FROM CUSTOMER WHERE name='CASH' & "','" _
'                                                  & CStr(DTPICKER4.Value) & "')"
                        
        
      SQL = "Insert into SalesReturnHeader values('" & SalesID & "','" _
      & CustID & "',#" & Date & "#)"
       ' CRED = False

        DCON.Execute SQL
     
 
    
        
        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into SalesReturnDetail values('" & SalesID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & ",'" & "1" & "')"
        
        DCON.Execute SQL
        
      
    Set Rs1 = New Recordset

        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "update Product set UnitsInStock=" & Val(Val(Rs1!UnitsInStock) + Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"

       DCON.Execute SQL
       
       'SQL = "Select * from SalesRegistryDetail where SalesRegistryID='" & cmbReg & "' And ProductID='" & .ListItems(CNT1).text & "'"
       SQL = "Select * from salesret where SalesRegistryID='" & cmbReg & "' And ProductID='" & .ListItems(CNT1).text & "'"
       
       Set Rs1 = DCON.Execute(SQL)
       
       SQL = "Update SalesRegistryDetail set Quantity=" & Val(Val(Rs1!quantity) - Val(.ListItems(CNT1).SubItems(2))) & " where SalesRegistryID='" & cmbReg & "' And ProductID='" & .ListItems(CNT1).text & "'"
        
     
       
       DCON.Execute SQL
       

'        SQL = "Delete * from SalesRegistryDetail where PRODUCTID='" & .ListItems(CNT1).text & "' AND SalesRegistryID='" & "1" & "'"
'
'       DCON.Execute SQL
'
'       Set RS1 = DCON.Execute("Select * from SalesRegistryDetail where SalesRegistryID='" & Combo5.text & "'")
'
'       If RS1.RecordCount = 0 Then
'       SQL = "Delete * from SalesRegistryHeader where SalesRegistryID='" & Combo5.text & "'"
'       DCON.Execute SQL
'       End If
       
        
        Next
       
        
'Call cmdSalesReturn_Click

Set Rs1 = Nothing
Set DCON = Nothing

End If

End With
End Sub
Private Sub cmdSalesOrder_Click()
Dim CDb As CDbase
Dim CIns As New CInsert
'Dim PurchID As String


''Dim CNT1 As Integer

Call GetNewConnection(CIns)
Set CDb = CIns


'PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO", TXTNUM)
'TXTNUM.text = PurchID

Set CIns = Nothing

'PurchOrder = True
'PurchReg = False
'PurchRet = False


    lvMain.ListItems.clear
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""

End Sub

'Private Sub cmdSalesReg_Click()
'Dim CDb As CDbase
'Dim CIns As New CInsert
''Dim PurchID As String
'
'
'''Dim CNT1 As Integer
'
'Call GetNewConnection(CIns)
'Set CDb = CIns
'
'
''PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PR", TXTNUM)
''TXTNUM.text = PurchID
'
'Set CIns = Nothing
'
'PurchReg = True
'PurchOrder = False
'PurchRet = False
'
'
'    Picture2.Visible = False
'    Picture3.Visible = False
'End Sub
'
'Private Sub cmdSalesReturn_Click()
'Dim CDb As CDbase
'Dim CIns As New CInsert
''Dim PurchID As String
'
'
'''Dim CNT1 As Integer
'
'Call GetNewConnection(CIns)
'Set CDb = CIns
'
'
'ReturnID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PRet")
'
'
'Set CIns = Nothing
'
'
'''Call CMB1("PurchaseRegistryHeader", "PurchaseRegistryID", Combo5, , True)
''
''    Picture2.Visible = False
''    Picture3.Visible = True
''    lvMain.ListItems.clear
''Text4.text = ""
''text5.text = ""
''txtQTY.text = ""
''txtRate.text = ""
'
'End Sub


'Private Sub Combo5_Click()
'On Error Resume Next
'
'Call GetNewConnection2
'
'Set Rs1 = New Recordset
'Set Rs1 = DCON.Execute("Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & Combo5.text & "'")
'
'While Not Rs1.EOF
'
'    Dim LVITEM As ListItem
'    Dim ProdID As String
'
'    With lvMain
'    .ListItems.clear
'
'    ProdID = Rs1!productid
'    Set LVITEM = .ListItems.Add(, , Rs1!productid)
'        LVITEM.SubItems(2) = Rs1!quantity
'        LVITEM.SubItems(4) = Rs1!discount
'        LVITEM.SubItems(3) = Rs1!rate
'
'    Set RS2 = New Recordset
'    Set RS2 = DCON.Execute("SElect * from product where productID='" & ProdID & "'")
'        LVITEM.SubItems(1) = RS2!Name
'
'
'    End With
'
'    Rs1.MoveNext
'Wend
'
'
'
'Set Rs1 = Nothing
'Set RS2 = Nothing
'Set DCON = Nothing
'End Sub

Private Sub Command1_Click()
Dim CDb As CDbase
Dim CIns As New CInsert
'Dim SalesID As String


Text4.SetFocus

Dim CNT1 As Integer

Call GetNewConnection(CIns)
Set CDb = CIns


SalesID = CIns.AUTONUM(CDb.OpenDb, "SalesReturnHeader", "SalesReturnID", "SRet")


Set CIns = Nothing

End Sub

Private Sub Form_Load()


    
    cmbCust.ListIndex = -1
    Call CMB2("Select distinct company from salesret", cmbCust)
    
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
       'frmMAIN.WindowState = 0
End Sub

Private Sub Label37_Click()

End Sub

Private Sub LvHeads()
    lvMain.ColumnHeaders(1).Width = lvMain.Width * 0.1
    lvMain.ColumnHeaders(2).Width = lvMain.Width * 0.2
    lvMain.ColumnHeaders(3).Width = lvMain.Width * 0.2

End Sub


Private Sub PurchOrderHeader()

Dim CDb As CDbase
Dim CIns As New CInsert
Dim PurchID As String
Dim CustID As String
Dim EmpId As String
Dim dtone As String
Dim dttwo As String
Dim CNT1 As Integer


Call GetNewConnection(CIns)
Set CDb = CIns


PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO")
'CustID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "SupplierID", "Supp") ' optional
'EmpId = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO")

CDb.TableName = "PurchaseOrderHeader"

'TXTNUM.text = PurchID

CIns.FieldVal PurchID, CText
CIns.FieldVal SuppID, CText
CIns.FieldVal CStr(cmbdate.text), CText


CIns.Insert

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   


        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into PurchaseOrderDetail values('" & PurchID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & ",'" & CStr(cmbdate.text) & "'," & .ListItems(CNT1).SubItems(3) & ")"
        DCON.Execute SQL
    
    Set Rs1 = New Recordset
  
        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"
       
       Set Rs1 = DCON.Execute(SQL)
       
      SQL = "update Product set UnitsInOrder=" & Val(Val(Rs1!UnitsInOrder) + Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"
                    
     DCON.Execute SQL
    
        
        Next

End If


Set DCON = Nothing
End With


End Sub



Private Function GetProduct(ProdID As String) As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select name from product where productid='" & ProdID & "'")

If Not Rs1.EOF Then

GetProduct = Rs1!Name

Else
    MsgBox "PRODUCT NOT FOUND"
    Exit Function
    
End If


Set Rs1 = Nothing
Set DCON = Nothing



End Function



Private Sub calculate()
Dim lst As ListItem
Dim i As Integer
Dim total As Long
If lvMain.ListItems.Count < 0 Then Exit Sub
For i = 1 To lvMain.ListItems.Count
   total = total + CLng(lvMain.ListItems(i).SubItems(2)) * CLng(lvMain.ListItems(i).SubItems(3))
Next i
    TextAmount = total
    total = 0
End Sub
Private Sub lvMain_DblClick()
'calculate
End Sub

Private Sub lvMain_GotFocus()
    calculate
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
      txtQTY.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
      txtRate.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(3)
      text5.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(5)
End If
End Sub

Private Sub lvMain_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DD As Integer

With lvMain
If .ListItems.Count <> 0 Then
If KeyCode = vbKeyDelete Then
TextAmount.text = ""
       
  .ListItems.Remove (.SelectedItem.Index)
  
Text4.text = ""

  'lblcat.Caption = ""
EDT = False
txtRate.text = ""
text5.text = ""
txtQTY.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""

            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
            Next
 End If
End If
End With

End Sub

Private Sub Text4_Change()

TXTLEN = Len(Text4.text)
STRT = 0
txtQTY.text = ""

End Sub

Private Sub TextAmount_Change()
lblwords.Caption = NumToWord(TextAmount.text)
End Sub



Private Sub txtqty_Change()
    text5.text = Val(txtQTY.text) * Val(txtRate.text)
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
    Call Decimals(KeyAscii, txtQTY, 2)
End Sub

Private Sub txtrate_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)

End Sub

Private Sub txtrate_KeyPress(KeyAscii As Integer)

Call Decimals(KeyAscii, txtRate, 2)

If KeyAscii = 13 Then

    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from Product where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'
            Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                
'                SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
               SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!productid & "'"
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!unitcostprice <> txtrate.text Then
'
'                         MsgBox "Cannot update UnitCostPrice" & vbTab, vbInformation, "UnitCostPrice"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!unitcostprice
'                         txtrate.SetFocus
'                      End If
                   
                    
            End If
         

Set Rs1 = Nothing
Set DCON = Nothing



End If
End Sub

Private Sub txtrate_LostFocus()
  
    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from Product where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
'
        Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                
'                SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtrate.text) & " where (PRODUCTID='" & RS1!productid & "' AND UnitCostPrice <" & Val(txtrate.text) & ")"
               SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!productid & "'"
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!unitcostprice <> txtrate.text Then
'
'                         MsgBox "Cannot update UnitCostPrice" & vbTab, vbInformation, "UnitCostPrice"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!unitcostprice
'                         txtrate.SetFocus
'                      End If
                   
                    
            End If
         

Set Rs1 = Nothing
Set DCON = Nothing


End Sub


