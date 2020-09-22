VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tempSalesReg 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProsVent Inventory Manager 2005"
   ClientHeight    =   8295
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
   ScaleHeight     =   8295
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   4815
      TabIndex        =   29
      Top             =   6960
      Width           =   4815
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   570
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   31
         Top             =   285
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not In Stock"
         Height          =   195
         Left            =   315
         TabIndex        =   36
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level Reached"
         Height          =   255
         Left            =   300
         TabIndex        =   35
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status OK"
         Height          =   195
         Left            =   330
         TabIndex        =   34
         Top             =   570
         Width           =   855
      End
      Begin VB.Label lblwords 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800080&
         Height          =   735
         Left            =   2325
         TabIndex        =   33
         Top             =   45
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox TextChange 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox TextAmount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      TabIndex        =   26
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox TextTend 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9600
      TabIndex        =   25
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Left            =   555
      Top             =   2640
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      Enabled         =   0   'False
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
      TabIndex        =   9
      Top             =   0
      Width           =   11910
      Begin VB.TextBox TXTNUM 
         Enabled         =   0   'False
         Height          =   375
         Left            =   9480
         TabIndex        =   21
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   1640
         Width           =   2415
      End
      Begin VB.TextBox txtQTY 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1640
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3945
         TabIndex        =   2
         Top             =   1640
         Width           =   735
      End
      Begin VB.TextBox text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   1640
         Width           =   1095
      End
      Begin VB.ComboBox cmbCust 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   405
         Left            =   7125
         TabIndex        =   4
         Top             =   1635
         Width           =   540
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38530
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Registry"
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
         TabIndex        =   12
         Top             =   120
         Width           =   2040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   4080
         TabIndex        =   20
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Registry ID:"
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
         Left            =   7515
         TabIndex        =   19
         Top             =   240
         Width           =   1920
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
         TabIndex        =   17
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Left            =   7920
         TabIndex        =   10
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   7920
         TabIndex        =   11
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
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5280
         TabIndex        =   15
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   3000
         TabIndex        =   14
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   240
         TabIndex        =   13
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
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   0
         Picture         =   "2_TEMP3.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12330
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4455
      Left            =   2760
      TabIndex        =   8
      Top             =   2280
      Width           =   9135
      _ExtentX        =   16113
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
      Enabled         =   0   'False
      NumItems        =   5
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
   End
   Begin MSComctlLib.ListView lvLook 
      Height          =   4455
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16380139
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tendered"
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
      Left            =   8400
      TabIndex        =   24
      Top             =   7440
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
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
      Left            =   8520
      TabIndex        =   23
      Top             =   7800
      Width           =   825
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   7650
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
      TabIndex        =   16
      Top             =   6960
      Width           =   1530
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000001&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F9F0EB&
      FillStyle       =   0  'Solid
      Height          =   1230
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
      Height          =   1320
      Left            =   7815
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3945
   End
   Begin VB.Menu mnuLook 
      Caption         =   "Look"
      Visible         =   0   'False
      Begin VB.Menu mnuUnit 
         Caption         =   "Unit In Stock"
      End
   End
End
Attribute VB_Name = "tempSalesReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleReg As Boolean
Dim SaleOrder As Boolean
Dim SaleRet As Boolean
Dim CustName As String
Dim CustID As String
Dim SalesID As String

Dim TXTLEN As Integer
Dim STRT As Integer
Dim PRILEN As Integer
Dim bLook As Boolean


Private Sub cmbCust_Click()


Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Customer where company='" & cmbCust.text & "'")

If Rs1.RecordCount <> 0 Then
    CustID = Rs1!customeriD
End If
If cmbCust.text <> "" Then
    Text4.Enabled = True
Else
    Text4.Enabled = False
End If


End Sub

Private Sub cmdAdd_Click()

Dim LSTITEM As ListItem



Dim CNT As Boolean
Dim DD As Integer

''' query in quantity is not yet included

txtQTY.text = Val(txtQTY.text)

If text5.text <> "" Then
   
    CNT = False
  

    Call GetNewConnection2
        Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from Product where ProductID like'" & Text4 & "%' OR Name like'" & Text4 & "%'")

If Not Rs1.EOF Then
    
     
    
       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
       txtRate = Rs1!UnitSellingPrice
            
  With lvMain
        TextAmount.text = ""
        
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                
              If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
                If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
               ' If StrComp(.ListItems(DD).text, RS1!ProductID) = 1 Then
                
                        If EDT = True Then
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(4) = text5.text
                        Else
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtQTY.text)
                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
                            .ListItems(DD).SubItems(4) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
                    
                        End If
                             
                    CNT = True
                    
                End If
              End If
                   
            Next
       
         End If
            
        If CNT = False Then

         .ListItems.Add , , Rs1!ProductID
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
           
            .ListItems(.ListItems.Count).SubItems(4) = text5.text
             
             
             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
        
        End If
           
            
             
        
        If .ListItems.Count <= 0 Then

            .ListItems.Add 1, , Rs1!ProductID
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
           
            .ListItems(.ListItems.Count).SubItems(4) = text5.text


        End If


            For DD = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(DD).SubItems(4)) + Val(TextAmount.text)
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
txtQTY.text = ""
txtRate.text = ""
EDT = False
'Text4.SetFocus

Else



txtQTY.SetFocus


End If


End Sub

Private Sub cmdAdd_GotFocus()
If Text4.text <> "" Then

    Call GetNewConnection2

    Set Rs1 = New Recordset
        SQL = "Select TOP 5 * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"

    Set Rs1 = DCON.Execute(SQL)

        SQL = "Select * from Product where (PRODUCTID='" & Rs1!ProductID & "' AND UnitCostPrice <" & Val(txtRate.text) & ")"
        
        Set Rs1 = DCON.Execute(SQL)
        
            If Rs1.RecordCount <> 0 Then

                SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(txtRate.text) & " where (PRODUCTID='" & Rs1!ProductID & "' AND UnitCostPrice <" & Val(txtRate.text) & ")"
   
                DCON.Execute SQL
               
                

            Else
               
                  SQL = "Select TOP 5 * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME='" & Text4 & "'"
                 Set Rs1 = DCON.Execute(SQL)
        
                
                If Rs1!UnitSellingPrice <> txtRate.text Then
                    MsgBox "Cannot Update UnitSellingPrice" & vbTab, vbInformation
                     txtRate.SetFocus
                End If
                
                If Rs1.RecordCount <> 0 Then
                    txtRate.text = Rs1!UnitSellingPrice
                End If
         
            End If
       
Set Rs1 = Nothing
Set DCON = Nothing

End If
End Sub

Private Sub cmdClear_Click()

EDT = False
lvMain.ListItems.clear

picTop.Enabled = False
cmbCust.Enabled = False
lvLook.Enabled = False
lvMain.Enabled = False
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""

'Text4.SetFocus
cmdOk.Enabled = False
cmdClear.Enabled = False
Command1.Enabled = True
TXTNUM.text = ""
cmbCust.ListIndex = -1


End Sub

Private Sub cmdOk_Click()

If lvMain.ListItems.Count > 0 Then
cmdClear.Enabled = False
cmdOk.Enabled = False
Command1.Enabled = True
picTop.Enabled = False
lvLook.Enabled = False
lvMain.Enabled = False
cmbCust.Enabled = False
Call SalesReg
    MsgBox "Record has been Saved", vbInformation
    
Else
    MsgBox "There is no Product to record", vbInformation
End If


End Sub









Private Sub Command1_Click()
Dim CDb As CDbase
Dim CIns As New CInsert
'Dim SalesID As String


Dim CNT1 As Integer

Call GetNewConnection(CIns)
Set CDb = CIns


SalesID = CIns.AUTONUM(CDb.OpenDb, "SalesRegistryHeader", "SalesRegistryID", "SR", TXTNUM)


Set CIns = Nothing

cmbCust.Enabled = True
cmdClear.Enabled = True
cmdOk.Enabled = True
Command1.Enabled = False
picTop.Enabled = True
lvLook.Enabled = True
lvMain.Enabled = True

lvMain.ListItems.clear
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""

End Sub

Private Sub Form_Load()
Timer2.Enabled = True
Timer2.Interval = 100


SaleReg = False
SaleRet = False
SaleOrder = False
bLook = False

   ' cmbCust.AddItem "Cash"
  '  cmbCust.ListIndex = 0
    Call CMB1("Customer", "Company", cmbCust)
    

    
End Sub

Private Sub Form_Unload(Cancel As Integer)

       'frmMAIN.WindowState = 0
End Sub


Private Sub LvHeads()
    lvMain.ColumnHeaders(1).Width = lvMain.Width * 0.1
    lvMain.ColumnHeaders(2).Width = lvMain.Width * 0.2
    lvMain.ColumnHeaders(3).Width = lvMain.Width * 0.2

End Sub







Private Sub SalesReg()

Dim CNT1 As Integer

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
   ' If CRED = True Then
   If cmbCust.text <> "Cash" Then
   
        'SalesRegistryHeader
        'SalesRegistryDetail
        
        SQL = "Insert into SalesRegistryHeader values('" & SalesID & "','" _
                                                  & CustID & "'," _
                                                  & 1 & ",'" & CStr(DTPicker4.Value) & "')"
                        
       ' CRED = False

      
        DCON.Execute SQL
     
    Else
            SQL = "Insert into SalesRegistryHeader values('" & SalesID & "','" _
                                                  & "Cash" & "'," _
                                                  & 1 & ",'" & CStr(DTPicker4.Value) & "')"
                        
                    
    
        DCON.Execute SQL
      
        
    End If

        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into SalesRegistryDetail values('" & SalesID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & "," & .ListItems(CNT1).SubItems(3) & ")"
        DCON.Execute SQL

    Set Rs1 = New Recordset

        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "update Product set UnitsInStock=" & Val(Val(Rs1!UnitsInStock) - Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"

       DCON.Execute SQL

        
        Next
       
        
Set Rs1 = Nothing
Set DCON = Nothing

End If

End With

End Sub



Private Sub lvLook_DblClick()
'Dim LVindex As Integer
'
'Dim LSTITEM As ListItem
'
'
'
'Dim CNT As Boolean
'Dim DD As Integer
'
'''' query in quantity is not yet included
'
'
'    CNT = False
'
'
'    Call GetNewConnection2
'        Set Rs1 = New Recordset
'        Set Rs1 = DCON.Execute("Select * from Product where ProductID='" & lvLook.SelectedItem.text & "'")
'
'If Not Rs1.EOF Then
'
'     If Rs1!UnitsInStock > 0 Then
'
'       'LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
'       txtRate = Rs1!UnitSellingPrice
'
'  With lvMain
'        TextAmount.text = ""
'
'        If .ListItems.Count <> 0 Then
'
'            For DD = 1 To .ListItems.Count
'
'              If InStr(1, .ListItems(DD).text, Rs1!ProductID) = 1 Then
'                If InStr(1, .ListItems(DD).SubItems(1), Rs1!Name) = 1 Then
'
'                        If EDT = True Then
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(4) = text5.text
'                        Else
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + 1
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(4) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
'
'                        End If
'
'
'
'                    CNT = True
'
'                End If
'
'
'
'               End If
'
'            Next
'
'         End If
'
'        If CNT = False Then
'
'
'         .ListItems.Add , , Rs1!ProductID
'            .ListItems(.ListItems.Count).SubItems(1) = Rs1!Name
'            .ListItems(.ListItems.Count).SubItems(2) = "1"
'            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
'            .ListItems(.ListItems.Count).SubItems(4) = text5.text
'            txtQTY.text = "1"
'
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
'                  TextAmount.text = Val(.ListItems(DD).SubItems(4)) + Val(TextAmount.text)
'            Next
'
'
'
'        Set Rs1 = Nothing
'        Set DCON = Nothing
'
'  End With
'  Else
'          MsgBox "This Product is out of Stock", vbInformation
'
'End If
'
'Else
'    MsgBox "Product Not Found", vbInformation, "Product"
'
'
'End If
'

End Sub

Private Sub lvLook_ItemClick(ByVal Item As MSComctlLib.ListItem)
Timer2.Enabled = False
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from product where productid='" & lvLook.SelectedItem.text & "'")

If Rs1!UnitsInStock > 0 Then
    
    If Not Rs1.EOF Then
        txtRate.text = Rs1!UnitSellingPrice

    End If


If lvLook.ListItems.Count > 0 Then
    Text4.text = lvLook.SelectedItem.text
    
End If

End If
End Sub

Private Sub lvLook_LostFocus()
Timer2.Enabled = True

End Sub

Private Sub lvLook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuLook
End If
End Sub

Private Sub lvMain_Click()
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
      
       txtQTY.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
    txtRate.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(3)
     text5.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4)
End If
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
      
       txtQTY.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
    txtRate.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(3)
     text5.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4)
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



Private Sub mnuUnit_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Product where ProductID='" & lvLook.SelectedItem.text & "'")

If Rs1.RecordCount <> 0 Then
    MsgBox "Available Stock: " & Rs1!UnitsInStock
End If
Set Rs1 = Nothing
Set DCON = Nothing
End Sub

Private Sub Text4_Change()
'If EDT = False Then
Timer2.Interval = 100


TXTLEN = Len(Text4.text)
STRT = 0

'End If
'EDT = True
txtQTY.text = ""

If Len(Text4.text) <> 0 Then
    txtQTY.Enabled = True
Else
    txtQTY.Enabled = False
    
End If

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Timer2.Interval = 0
End If

End Sub

Private Sub TextAmount_Change()
lblwords.Caption = NumToWord(TextAmount.text)
End Sub



Private Sub TextTend_Change()
  TextChange.text = Val(TextTend.text) - Val(TextAmount.text)
 
End Sub

Private Sub TextTend_KeyPress(KeyAscii As Integer)
Call Decimals(KeyAscii, TextTend, 2)

End Sub

'    Private Sub TextTend_KeyPress(KeyAscii As Integer)
'    'Call Decimals(KeyAscii, TextTend, 2)
'
'    End Sub

Private Sub Timer2_Timer()

'Static c As Integer


STRT = STRT + 1

If STRT = 3 Then
Timer2.Interval = 0
  




Dim FVAL As String
Dim DD As Integer
Dim LISTITM As ListItem

'DataGrid1.Visible = True

Call GetNewConnection2

Set Rs1 = New Recordset

'SQL = "Select TOP 10 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
SQL = "Select TOP 10 * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME like'" & Text4 & "%'"

Set Rs1 = DCON.Execute(SQL)

'Set DataGrid1.DataSource = RS1
     Set RS2 = New Recordset
        Set RS2 = DCON.Execute(SQL)
        lvLook.ListItems.clear
        While Not RS2.EOF
        Set LISTITM = lvLook.ListItems.Add(, , RS2!ProductID)
        LISTITM.SubItems(1) = RS2!Name
        
        If RS2!UnitsInStock <= 0 Then
            LISTITM.ForeColor = vbRed
            LISTITM.ListSubItems(1).ForeColor = vbRed
        Else
        
        If RS2!UnitsInStock <= RS2!ReorderLevel Then
            LISTITM.ForeColor = vbBlue
            LISTITM.ListSubItems(1).ForeColor = vbBlue
        End If
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
        txtRate.text = Rs1!UnitSellingPrice
        text5.text = Val(txtQTY.text) * Val(txtRate.text)
        
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

   
     

    Set RS2 = Nothing
    Set Rs1 = Nothing
    Set DCON = Nothing


ElseIf Text4.text = "" Then
   txtRate.text = ""
        text5.text = ""
        txtQTY.text = ""
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
text5.text = Val(txtQTY.text) * Val(txtRate.text)

If Len(txtQTY.text) <> 0 Then
    txtRate.Enabled = True
Else
    txtRate.Enabled = False
    
End If

If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
End Sub

Private Sub txtQTY_GotFocus()
If Text4.text <> "" Then
Call GetNewConnection2
Set Rs1 = New Recordset

SQL = "Select TOP 10 * from PRODUCT where PRODUCTID='" & Text4 & "' OR NAME like'" & Text4 & "%'"

Set Rs1 = DCON.Execute(SQL)
    
    If Rs1.RecordCount <> 0 Then
        If Rs1!UnitsInStock <= 0 Then
            MsgBox "This Product is out of Stock", vbInformation
            Text4.SetFocus
            Text4.SelStart = 0
            Text4.SelLength = Len(Text4.text)
        End If
     End If
     
Set Rs1 = Nothing
Set DCON = Nothing

End If

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtQTY)
If KeyAscii = 13 Then
    
    Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * From Product where ProductID='" & Text4 & "' OR Name='" & "'")
     If Rs1.RecordCount <> 0 Then
        If Rs1!UnitsInStock < txtQTY.text Then
            MsgBox "The Quantity needed is greater than Stock", vbInformation
            txtQTY.SetFocus
            txtQTY.SelStart = 0
            txtQTY.SelLength = Len(txtQTY.text)
            
        End If
     End If
     
  Set Rs1 = Nothing
  Set DCON = Nothing
  
End If
End Sub

Private Sub txtRate_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)
If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If

End Sub

Private Sub txtRate_GotFocus()
If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If

If Val(txtQTY.text) <> 0 Then
  Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * From Product where ProductID='" & Text4 & "' OR Name='" & "'")
     If Rs1.RecordCount <> 0 Then
     
     
        If Rs1!UnitsInStock < txtQTY.text Then
            MsgBox "The Quantity needed is greater than Stock", vbInformation
            txtQTY.SetFocus
            txtQTY.SelStart = 0
            txtQTY.SelLength = Len(txtQTY.text)
            
       
        End If
     End If
     
  Set Rs1 = Nothing
  Set DCON = Nothing
Else
    MsgBox "Please give specific quantity needed", vbInformation
    txtQTY.SetFocus
    
End If

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

Call Decimals(KeyAscii, txtRate, 2)

If KeyAscii = 13 Then

    Call GetNewConnection2

    Set Rs1 = New Recordset
        SQL = "Select TOP 5 * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"

    Set Rs1 = DCON.Execute(SQL)

        SQL = "Select * from Product where (PRODUCTID='" & Rs1!ProductID & "' AND UnitCostPrice <" & Val(txtRate.text) & ")"
        
        Set Rs1 = DCON.Execute(SQL)
        
            If Rs1.RecordCount <> 0 Then

                SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(txtRate.text) & " where (PRODUCTID='" & Rs1!ProductID & "' AND UnitCostPrice <" & Val(txtRate.text) & ")"
   
                DCON.Execute SQL
                 Call cmdAdd_Click
                txtRate.SetFocus

            Else
                
                SQL = "Select * from PRODUCT where PRODUCTID like '" & Text4 & "%' OR NAME like'" & Text4 & "%'"
                Set Rs1 = DCON.Execute(SQL)
                
                
                If Rs1!UnitSellingPrice <> txtRate.text Then
                    MsgBox "Cannot Update UnitSellingPrice" & vbTab, vbInformation
                End If
                
                If Rs1.RecordCount <> 0 Then
                    txtRate.text = Rs1!UnitSellingPrice
                End If
         
            End If
       
Set Rs1 = Nothing
Set DCON = Nothing


End If
End Sub

