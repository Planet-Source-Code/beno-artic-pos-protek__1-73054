VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchase1 
   Caption         =   "Form1"
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
   Begin VB.CommandButton cmdGetProd 
      Caption         =   "Get Product"
      Height          =   255
      Left            =   10320
      TabIndex        =   42
      Top             =   5400
      Width           =   1335
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Height          =   2310
      Left            =   0
      ScaleHeight     =   2250
      ScaleWidth      =   11820
      TabIndex        =   28
      Top             =   6285
      Width           =   11880
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   855
         Left            =   600
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   855
         Left            =   1920
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   855
         Left            =   3360
         TabIndex        =   39
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   6840
         TabIndex        =   32
         Top             =   120
         Width           =   4935
         Begin VB.TextBox Text2 
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
            TabIndex        =   35
            Text            =   "Text2"
            Top             =   210
            Width           =   2175
         End
         Begin VB.TextBox Text7 
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
            Left            =   2460
            TabIndex        =   34
            Text            =   "Text2"
            Top             =   780
            Width           =   2175
         End
         Begin VB.TextBox Text8 
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
            Left            =   2460
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   1335
            Width           =   2175
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Tendered :"
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
            Left            =   150
            TabIndex        =   38
            Top             =   840
            Width           =   2100
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
            TabIndex        =   37
            Top             =   315
            Width           =   1635
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change :"
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
            Left            =   1245
            TabIndex        =   36
            Top             =   1365
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdSalesReg 
         Caption         =   "Purchase Registry F1"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalesReturn 
         Caption         =   "Purchase Return F2"
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalesOrder 
         Caption         =   "Purchase Order F3"
         Height          =   375
         Left            =   3840
         TabIndex        =   29
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.PictureBox picBody 
      Align           =   1  'Align Top
      Height          =   3480
      Left            =   0
      ScaleHeight     =   3420
      ScaleWidth      =   11820
      TabIndex        =   19
      Top             =   1815
      Width           =   11880
      Begin MSComctlLib.ListView lvMain 
         Height          =   1695
         Left            =   0
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
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
         OLEDragMode     =   1
         OLEDropMode     =   1
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
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1530
         TabIndex        =   4
         Top             =   60
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
         Format          =   57606145
         CurrentDate     =   38530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Required"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Frame frameTop 
         BackColor       =   &H00EEECE8&
         Height          =   1695
         Left            =   8865
         TabIndex        =   22
         Top             =   30
         Width           =   3255
         Begin VB.TextBox TXTNUM 
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   720
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   960
            TabIndex        =   24
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   57606145
            CurrentDate     =   38530
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "___ID"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Balance:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1470
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         Height          =   375
         Left            =   8160
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   5640
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EEECE8&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3615
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Order ID"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1290
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
         Left            =   1680
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   4080
         TabIndex        =   17
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate/Unit"
         Height          =   195
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   6840
         TabIndex        =   15
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPurchase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PurchReg As Boolean
Dim PurchOrder As Boolean
Dim PurchRet As Boolean


Private Sub cmdAdd_Click()
Dim LSTITEM As ListItem
With lvMain
Set LSTITEM = .ListItems.Add(, , GetProduct(Text4))
End With
End Sub

Private Sub cmdOk_Click()
If PurchOrder = True Then
Call PurchOrderHeader

ElseIf PurchRet = True Then


ElseIf PurchReg = True Then


End If

End Sub

Private Sub cmdSalesOrder_Click()
PurchOrder = True
PurchReg = False
PurchRet = False

    Picture2.Visible = True
    Picture3.Visible = False
End Sub

Private Sub cmdSalesReg_Click()

PurchReg = True
PurchOrder = False
PurchRet = False


    Picture2.Visible = False
    Picture3.Visible = False
End Sub

Private Sub cmdSalesReturn_Click()
PurchRet = True
PurchOrder = False
PurchReg = False

    Picture2.Visible = False
    Picture3.Visible = True
End Sub

Private Sub Form_Activate()
Call ReOrder

End Sub

Private Sub Form_Load()
PurchReg = False
PurchRet = False
PurchOrder = False

    cmbCust.AddItem "Cash"
    cmbCust.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
picBody.Height = Me.ScaleHeight - picTop.Height - picBottom
    lvMain.Width = Me.ScaleWidth * 0.8
    lvLook.Left = lvMain.Width
    lvLook.Width = Me.ScaleWidth * 0.2
    lvLook.Height = picBody.ScaleHeight
    lvMain.Height = picBody.ScaleHeight
frameTop.Left = Me.ScaleWidth - frameTop.Width - 100
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

Private Sub TextBox25_Change()

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
CustID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "SupplierID", "Supp") ' optional
'EmpId = CIns.AUTONUM(CDb.OpenDb, "SalesOrderHeader", "SalesOrderID", "SO")

CDb.TableName = "PurchaseOrderHeader"


CIns.FieldVal PurchID, CText
CIns.FieldVal CustID, CText
CIns.FieldVal CStr(DTPicker4.Value), CText


CIns.Insert

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   


        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into PurchaseOrderDetail values('" & PurchID & "','" _
        & CNT1 & "'," & "3,'" & CStr(DTPicker1.Value) & "')"
        DCON.Execute SQL
    
   ' Set RS1 = New Recordset
  '
      '  SQL = "Select * from Product where productid='" & .ListItems(CNT1).SubItems(1) & "'"
       
     '  Set RS1 = dcon.Execute(SQL)
       
    '   SQL = "update Product set item_no_stock=" & Val(Val(RS1!UnitsInStock) - Val(.ListItems(CNT1).SubItems(3))) _
                    & " WHERE ITEM_ID='" & .ListItems(CNT1).SubItems(1) & "'"
                    
    '   dcon.Execute SQL
    
        
        Next

End If

End With


End Sub


Private Function GetProduct(ProdID As String) As String


Call GetNewConnection2

Set RS1 = New Recordset
Set RS1 = DCON.Execute("Select name from product where productid='" & ProdID & "'")

If Not RS1.EOF Then

GetProduct = RS1!Name

Else
    MsgBox "PRODUCT NOT FOUND"
    Exit Function
    
End If


Set RS1 = Nothing
Set DCON = Nothing



End Function

Private Sub ReOrder()
Call GetNewConnection2

With lvLook

SQL = "Select * from Product where UnitsInStock <= ReOrderLevel"

Set RS1 = New Recordset
Set RS1 = DCON.Execute(SQL)

While Not RS1.EOF

 .ListItems.Add , , RS1!Name
 


RS1.MoveNext
Wend






End With

End Sub


