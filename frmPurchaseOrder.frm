VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchaseOrder 
   Caption         =   "Purchase"
   ClientHeight    =   8595
   ClientLeft      =   360
   ClientTop       =   2235
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   1200
      Top             =   5520
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Height          =   2310
      Left            =   0
      ScaleHeight     =   2250
      ScaleWidth      =   11820
      TabIndex        =   26
      Top             =   6285
      Width           =   11880
      Begin VB.CommandButton cmdOk 
         Caption         =   "Save"
         Height          =   855
         Left            =   600
         TabIndex        =   31
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   855
         Left            =   3360
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   6840
         TabIndex        =   27
         Top             =   120
         Width           =   4935
         Begin VB.TextBox TextAmount 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2445
            TabIndex        =   28
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   630
            TabIndex        =   29
            Top             =   315
            Width           =   1635
         End
      End
      Begin VB.Label lblwords 
         Caption         =   "Label7"
         Height          =   615
         Left            =   0
         TabIndex        =   32
         Top             =   1560
         Width           =   6375
      End
   End
   Begin VB.PictureBox picBody 
      Align           =   1  'Align Top
      Enabled         =   0   'False
      Height          =   3480
      Left            =   0
      ScaleHeight     =   3420
      ScaleWidth      =   11820
      TabIndex        =   18
      Top             =   1815
      Width           =   11880
      Begin MSComctlLib.ListView lvMain 
         Height          =   1695
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2990
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvLook 
         Height          =   1695
         Left            =   10440
         TabIndex        =   20
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2990
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sales ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEECE8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   3
      Top             =   480
      Width           =   3975
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483639
         CalendarTitleBackColor=   16777215
         Format          =   60948481
         CurrentDate     =   38530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Date"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00EEECE8&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.CommandButton cmdSalesOrder 
         Caption         =   "New"
         Height          =   375
         Left            =   4920
         TabIndex        =   33
         Top             =   120
         Width           =   1695
      End
      Begin VB.Frame frameTop 
         BackColor       =   &H00EEECE8&
         Height          =   1695
         Left            =   8865
         TabIndex        =   21
         Top             =   30
         Width           =   3255
         Begin VB.TextBox TXTNUM 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   720
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   60948481
            CurrentDate     =   38530
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   60
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   1275
         Width           =   375
      End
      Begin VB.TextBox txtqty 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4080
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6840
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtrate 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4920
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EEECE8&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   60
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Order ID"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1605
         End
      End
      Begin VB.ComboBox cmbCust 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   4080
         TabIndex        =   16
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/Unit"
         Height          =   195
         Left            =   4920
         TabIndex        =   15
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   6840
         TabIndex        =   14
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PurchReg As Boolean
Dim PurchOrder As Boolean
Dim PurchRet As Boolean
Dim SuppName As String
Dim SuppID As String
Dim PurchID As String
Dim ReturnID As String


Dim TXTLEN As Integer
Dim STRT As Integer
Dim PRILEN As Integer

Private Sub cmbCust_Click()

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Supplier where BusinessName='" & cmbCust & "'")

If Not Rs1.EOF Then
    SuppID = Rs1!suppliersid
End If


If cmbCust.text <> "" Then
    picBody.Enabled = True
Else
    picBody.Enabled = False
End If


End Sub

Private Sub cmdAdd_Click()

Dim LSTITEM As ListItem



Dim CNT As Boolean
Dim DD As Integer
Dim findList As ListItem




If text5.text <> "" Then
   
    CNT = False
  

    Call GetNewConnection2
        Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from Product where ProductID like'" & Text4 & "%' OR Name like'" & Text4 & "%'")

If Not Rs1.EOF Then
    
     
'Set LSTITEM = ListView1.FindItem(RS1!productid, lvwText, , lvwPartial)
'       If LSTITEM Is Nothing Then
            
       
       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
       txtRate = Rs1!UnitCostPrice
            
  With lvMain
        TextAmount.text = ""
        
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
               
                If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
                    If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
              
                        If EDT = True Then
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(txtqty.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(5) = text5.text
                        Else
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtqty.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
                    
                        End If
                             
                    CNT = True
                    
                    End If
                End If
                   
            Next
       
         End If
            
        If CNT = False Then

         .ListItems.Add , , Rs1!ProductID
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
            .ListItems(.ListItems.Count).SubItems(4) = "1"
            .ListItems(.ListItems.Count).SubItems(5) = text5.text
             
             
             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
        
        End If
           
            
             
        
'        If .ListItems.Count <= 0 Then
'
'            .ListItems.Add 1, , RS1!ProductID
'            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
'
'
'        End If


            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
            Next
            
        ' lblunit.Caption = RS1!UnitsInStock
       
        
       Set Rs1 = DCON.Execute("Select * from Product")
      '  Set DataGrid1.DataSource = RS1
         
        Set Rs1 = Nothing
        Set DCON = Nothing
   
  End With

Else
    MsgBox "Product Not Found", vbInformation, "Product"
    
    
End If
    
Text4.text = ""
text5.text = ""
txtqty.text = ""
txtRate.text = ""
EDT = False
Text4.SetFocus

Else



txtqty.SetFocus


End If


End Sub

Private Sub cmdAdd_GotFocus()
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
End Sub


Private Sub cmdClear_Click()
lvMain.ListItems.clear

End Sub

Private Sub cmdOk_Click()
If lvMain.ListItems.Count > 0 Then
Call PurchOrderHeader
Else
    MsgBox "THERE IS NO PRODUCT TO RECORD", vbInformation
End If



End Sub

Private Sub cmdSalesOrder_Click()


  Dim CDb As CDbase
Dim CIns As New CInsert
'Dim PurchID As String


''Dim CNT1 As Integer

Call GetNewConnection(CIns)
Set CDb = CIns


 CIns.AUTONUM CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO", TXTNUM
 


Set CIns = Nothing
    

cmbCust.ListIndex = -1

    
    lvMain.ListItems.clear
Text4.text = ""
text5.text = ""
txtqty.text = ""
txtRate.text = ""

End Sub




Private Sub Combo5_Click()
On Error Resume Next

Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & Combo5.text & "'")

While Not Rs1.EOF
    
    Dim LVITEM As ListItem
    Dim ProdID As String
    
    With lvMain
    .ListItems.clear
    
    ProdID = Rs1!ProductID
    Set LVITEM = .ListItems.Add(, , Rs1!ProductID)
        LVITEM.SubItems(2) = Rs1!Quantity
        LVITEM.SubItems(4) = Rs1!discount
        LVITEM.SubItems(3) = Rs1!Rate
        
    Set RS2 = New Recordset
    Set RS2 = DCON.Execute("SElect * from product where productID='" & ProdID & "'")
        LVITEM.SubItems(1) = RS2!Name


    End With
    
    Rs1.MoveNext
Wend



Set Rs1 = Nothing
Set RS2 = Nothing
Set DCON = Nothing
End Sub

Private Sub DTPicker1_Validate(KEEPFOCUS As Boolean)
If DTPicker1.Value < DTPicker4.Value Then
    MsgBox "The Date is not Valid", vbInformation
    KEEPFOCUS = True
    
End If

End Sub

Private Sub Form_Load()

Timer2.Enabled = True
Timer2.Interval = 100
   ' cmbCust.AddItem "Cash"
    cmbCust.ListIndex = -1
    Call CMB1("Supplier", "BusinessName", cmbCust)
  
    
End Sub

Private Sub Form_Resize()
'picBody.Height = Me.ScaleHeight - picTop.Height - picBottom
    lvMain.Width = Me.ScaleWidth * 0.8
    lvLook.Left = lvMain.Width
    lvLook.Width = Me.ScaleWidth * 0.2
    lvLook.Height = picBody.ScaleHeight
    lvMain.Height = picBody.ScaleHeight
frameTop.Left = Me.ScaleWidth - frameTop.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)

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


PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO", TXTNUM)
'CustID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "SupplierID", "Supp") ' optional
'EmpId = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PO")

CDb.TableName = "PurchaseOrderHeader"

TXTNUM.text = PurchID

CIns.FieldVal PurchID, CText
CIns.FieldVal SuppID, CText
CIns.FieldVal CStr(DTPicker4.Value), CText


CIns.Insert

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   


        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into PurchaseOrderDetail values('" & PurchID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & ",'" & CStr(DTPicker1.Value) & "'," & .ListItems(CNT1).SubItems(3) & ",FALSE" & ")"
        DCON.Execute SQL
    
    Set Rs1 = New Recordset
  
        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"
       
       Set Rs1 = DCON.Execute(SQL)
       
      SQL = "update Product set UnitsInOrder=" & Val(Val(Rs1!UnitsInOrder) + Val(.ListItems(CNT1).SubItems(2))) _
                    & ",DISCONTINUED='1' WHERE ProductID='" & .ListItems(CNT1).text & "'"
                    
     DCON.Execute SQL
    
        
        Next

End If


Set DCON = Nothing
End With


End Sub
Private Sub ReOrder()
Call GetNewConnection2

With lvLook

SQL = "Select * from Product where UnitsInStock <= ReOrderLevel"

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)

While Not Rs1.EOF

 .ListItems.Add , , Rs1!Name
 


Rs1.MoveNext
Wend






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




Private Sub lvLook_DblClick()
Dim LVindex As Integer

Dim LSTITEM As ListItem



Dim CNT As Boolean
Dim DD As Integer

''' query in quantity is not yet included


   
    CNT = False
  

    Call GetNewConnection2
        Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from Product where ProductID='" & lvLook.SelectedItem.text & "'")

If Not Rs1.EOF Then
    
     
    
       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
       txtRate = Rs1!UnitCostPrice
       txtqty.text = Rs1!ReorderQuantity
            
  With lvMain
        TextAmount.text = ""
        
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                  
              If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
                If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
                  
                        If EDT = True Then
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(txtqty.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(5) = text5.text
                        Else
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtqty.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
                    
                        End If
                             
                    CNT = True
                    
                End If
                    
               End If
               
            Next
       
         End If
            
        If CNT = False Then

         .ListItems.Add , , Rs1!ProductID
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
            .ListItems(.ListItems.Count).SubItems(4) = "1"
            .ListItems(.ListItems.Count).SubItems(5) = text5.text
             
             
             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
        
        End If
           
            
             
        
'        If .ListItems.Count <= 0 Then
'
'            .ListItems.Add 1, , RS1!ProductID
'            .ListItems(.ListItems.Count).SubItems(1) = RS1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
'
'
'        End If


            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
            Next
            
        ' lblunit.Caption = RS1!UnitsInStock
       
        
       Set Rs1 = DCON.Execute("Select * from Product")
      '  Set DataGrid1.DataSource = RS1
         
        Set Rs1 = Nothing
        Set DCON = Nothing
   
  End With

Else
    MsgBox "Product Not Found", vbInformation, "Product"
    
    
End If
End Sub

Private Sub lvLook_ItemClick(ByVal Item As MSComctlLib.ListItem)
Timer2.Enabled = False
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from product where productid='" & lvLook.SelectedItem.text & "'")

    If Not Rs1.EOF Then
        txtRate.text = Rs1!UnitCostPrice
    End If
    

If lvLook.ListItems.Count > 0 Then
    Text4.text = lvLook.SelectedItem.text
    
End If

End Sub

Private Sub lvLook_LostFocus()
Timer2.Enabled = True
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
If PurchRet = True Then
    Timer2.Enabled = False
    

Else

End If
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
      
       txtqty.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
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
txtqty.text = ""
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
If Len(Text4.text) = 0 Then
    Timer2.Enabled = True
End If

'If EDT = False Then
Timer2.Interval = 100


TXTLEN = Len(Text4.text)
STRT = 0

'End If
'EDT = True
txtqty.text = ""

End Sub

Private Sub TextAmount_Change()
lblwords.Caption = NumToWord(TextAmount.text)
End Sub



Private Sub TextTend_KeyPress(KeyAscii As Integer)
'Call Decimals(KeyAscii, TextTend, 2)

End Sub

Private Sub TextChange_Change()

End Sub

Private Sub Timer2_Timer()

'Static c As Integer


STRT = STRT + 1



If STRT = 3 Then
Timer2.Interval = 0
  




Dim FVAL As String
Dim DD As Integer
Dim LISTITM As ListItem

Call GetNewConnection2

Set Rs1 = New Recordset

'SQL = "Select TOP 10 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
SQL = "Select TOP 10 *,(UnitsInStock + UnitsInOrder) as Total from PRODUCT where (PRODUCTID='" & Text4 & "' OR NAME like'" & Text4 & "%')"
'SQL = "Select Top 20 * from lowINstock order by Total"
Set Rs1 = DCON.Execute(SQL)
 Set RS2 = New Recordset
        Set RS2 = DCON.Execute(SQL)
        lvLook.ListItems.clear
        While Not RS2.EOF
        
            Set LISTITM = lvLook.ListItems.Add(, , RS2!ProductID)
            
                LISTITM.SubItems(1) = RS2!Name
                
                If RS2!total <= 0 Then
                    LISTITM.ForeColor = vbRed
                End If
        
            RS2.MoveNext
        Wend

If Text4.text <> "" Then

    If Not Rs1.EOF Then
'        TXT_CODE.SelStart = PRILEN
'        TXT_CODE.text = RS1!Name
'        TXT_CODE.SelLength = Len(TXT_CODE.text)
'
      
        FVAL = Rs1!ProductID
        
      
       txtRate.text = Rs1!UnitCostPrice
       
        
       text5.text = Val(txtqty.text) * Val(txtRate.text)
        
        'lblselling.Caption = RS1!UnitSellingPrice
        'lblunit.Caption = RS1!UnitsInStock

  With lvMain
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                  
                If InStr(1, .ListItems(DD).SubItems(1), Rs1!ProductID) = 1 Then
                  
                 
                            .ListItems(DD).Selected = True
                         '   lblunit.Caption = Val(lblunit.Caption) - Val(.ListItems(DD).SubItems(3))
                    
                    
                End If
            
               
            Next
         End If
    End With

  

    Else
      txtRate.text = ""
        text5.text = ""
'        TXT_QTY.text = ""
'        lblselling.Caption = ""
'        lblunit.Caption = ""
'        lblcat.Caption = ""
'        PRILEN = 0

    End If

    
    Set Rs1 = Nothing
    Set RS2 = Nothing
    Set DCON = Nothing


ElseIf Text4.text = "" Then
   txtRate.text = ""
        text5.text = ""
        txtqty.text = ""
'lblcat.Caption = ""
'EDT = False
'TXT_RATE.text = ""
'TXT_AMT.text = ""
'TXT_QTY.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""
'TXT_CODE.SetFocus
End If

End If



End Sub

Private Sub txtqty_Change()
text5.text = Val(txtqty.text) * Val(txtRate.text)

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call Decimals(KeyAscii, txtqty, 2)

End Sub

Private Sub txtqty_Validate(Cancel As Boolean)
    txtqty.text = Round(Val(txtqty), 0)
End Sub

Private Sub txtRate_Change()
    text5.text = Val(txtqty.text) * Val(txtRate.text)

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

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
               SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!ProductID & "'"
              
                
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
               SQL = "UPDATE PRODUCT set UnitCostPrice=" & Val(txtRate.text) & " where PRODUCTID='" & Rs1!ProductID & "'"
              
                
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
