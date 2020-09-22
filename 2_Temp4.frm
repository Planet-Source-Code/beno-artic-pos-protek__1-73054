VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form tempPRET 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProsVent Inventory Manager 2005"
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
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   5055
      TabIndex        =   26
      Top             =   6960
      Width           =   5055
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         Top             =   570
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         Top             =   285
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   15
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not In Stock"
         Height          =   195
         Left            =   315
         TabIndex        =   33
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder Level Reached"
         Height          =   255
         Left            =   300
         TabIndex        =   32
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status OK"
         Height          =   195
         Left            =   330
         TabIndex        =   31
         Top             =   570
         Width           =   855
      End
      Begin VB.Label lblwords 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800080&
         Height          =   735
         Left            =   2325
         TabIndex        =   30
         Top             =   30
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   11850
      TabIndex        =   13
      Top             =   0
      Width           =   11910
      Begin VB.ComboBox cmbdate 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbreg 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TXTNUM 
         Enabled         =   0   'False
         Height          =   375
         Left            =   9480
         TabIndex        =   24
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtQTY 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2910
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4065
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5625
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbCust 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   405
         Left            =   7080
         TabIndex        =   10
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order ID"
         Height          =   195
         Left            =   7800
         TabIndex        =   25
         Top             =   1740
         Width           =   1605
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   8640
         TabIndex        =   23
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   4320
         TabIndex        =   22
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order  ID"
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
         Left            =   7395
         TabIndex        =   21
         Top             =   240
         Width           =   2040
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
         TabIndex        =   20
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   960
         Width           =   705
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Return"
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
         TabIndex        =   15
         Top             =   120
         Width           =   2325
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
         Picture         =   "2_Temp4.frx":0000
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
         TabIndex        =   18
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   195
         TabIndex        =   16
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
      TabIndex        =   11
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      TabIndex        =   12
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
      TabIndex        =   19
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
Attribute VB_Name = "tempPRET"
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


Dim TxtLen As Integer
Dim STRT As Integer
Dim PRILEN As Integer

Private Sub PurchaseReturn()

Dim CNT1 As Integer

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
   ' If CRED = True Then
  '' If cmbCust.text <> "Cash" Then
   
        'PurchaseRegistryHeader
        'PurchaseRegistryDetail
        
        SQL = "Insert into PurchaseReturnHeader values('" & ReturnID & "','" _
                                                  & SuppID & "','" & cmbdate.text & "')"
                        
       ' CRED = False

      
        DCON.Execute SQL
     
   '' Else
     '  '     SQL = "Insert into PurchaseRegistryHeader values('" & PurchID & "','" _
                                                  & "Cash" & "'," _
                                                  & 1 & ",'" & CStr(DTPicker4.Value) & "')"
                        
                    
    
     '  ' DCON.Execute SQL
      
        
  ' ' End If

        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into PurchaseReturnDetail values('" & ReturnID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & ",'" & cmbreg.text & "')"
        DCON.Execute SQL

    Set Rs1 = New Recordset

        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "update Product set UnitsInStock=" & Val(Val(Rs1!UnitsInStock) - Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"

       DCON.Execute SQL
       
         SQL = "Select * from Purret where PurchaseRegistryID='" & cmbreg & "' And ProductID='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "Update PurchaseRegistryDetail set Quantity=" & Val(Val(Rs1!Quantity) - Val(.ListItems(CNT1).SubItems(2))) & ",Rate=" & .ListItems(CNT1).SubItems(3) & " where PurchaseRegistryID='" & cmbreg & "' And ProductID='" & .ListItems(CNT1).text & "'"



       DCON.Execute SQL
'        SQL = "update PurchaseRegistryDetail set Quantity=" & Val(.ListItems(CNT1).SubItems(2)) _
'                    & ",Rate=" & Val(.ListItems(3).text) & " WHERE ProductID='" & .ListItems(CNT1).text & "'"
'
'       DCON.Execute SQL


        SQL = "Delete * from PurchaseRegistryDetail where (PRODUCTID='" & .ListItems(CNT1).text & "' AND PurchaseRegistryID='" & cmbreg.text & "') AND Quantity<=0"
        
       DCON.Execute SQL
       
       Set Rs1 = DCON.Execute("Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & cmbreg.text & "'")
       
       If Rs1.RecordCount = 0 Then
       SQL = "Delete * from PurchaseRegistryHeader where PurchaseRegistryID='" & cmbreg.text & "'"
       DCON.Execute SQL
       End If
       

        
        Next
       
        
Set Rs1 = Nothing
Set DCON = Nothing

End If

End With
End Sub
Private Sub UpdateReturn()
Dim CNT1 As Integer
Call GetNewConnection2

SQL = "Delete * from PurchaseReturnDetail where PurchaseReturnID='" & TXTNUM.text & "'"

DCON.Execute SQL

With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
             
                    
    
  

        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into PurchaseReturnDetail values('" & ReturnID & "','" _
        & .ListItems(CNT1).text & "'," & .ListItems(CNT1).SubItems(2) & ",'" & cmbreg.text & "')"
        
        DCON.Execute SQL

    Set Rs1 = New Recordset

        SQL = "Select * from Product where productid='" & .ListItems(CNT1).text & "'"

       Set Rs1 = DCON.Execute(SQL)

       SQL = "update Product set UnitsInStock=" & Val(Val(Rs1!UnitsInStock) - Val(.ListItems(CNT1).SubItems(2))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).text & "'"

       DCON.Execute SQL
       
       

        
        Next
        
'        SQL = "Delete * from PurchaseReturnDetail where PRODUCTID='" & .ListItems(CNT1).text & "' AND SalesRegistryID='" & Combo5.text & "'"
'
'       DCON.Execute SQL
'
'       Set Rs1 = DCON.Execute("Select * from SalesRegistryDetail where SalesRegistryID='" & Combo5.text & "'")
'
'       If Rs1.RecordCount = 0 Then
'       SQL = "Delete * from SalesRegistryHeader where SalesRegistryID='" & Combo5.text & "'"
'       DCON.Execute SQL
'       End If
       
        
Set Rs1 = Nothing
Set DCON = Nothing



'
'SQL = "Delete * from PurchaseReturnDetail where PurchaseReturnID='" & TXTNUM.text & "'"
'
'DCON.Execute SQL
'SQL = "Delete * from PurchaseReturnHeader where PurchaseReturnID='" & TXTNUM.text & "'"
'
'DCON.Execute SQL

TXTNUM.text = ""

End If

End With
End Sub
Private Sub ReturnModify()
Dim PID As String
 SQL = "Select Distinct * from ModifyPurchase where PurchaseReturnID='" & MODIFYID & "'"
 
 Call GetNewConnection2
 Set Rs1 = New Recordset
 Set Rs1 = DCON.Execute(SQL)
 

   lvMain.ListItems.clear
    
While Not Rs1.EOF
    Dim LVITEM As ListItem
    Dim ProdID As String
    With lvMain
    ProdID = Rs1!ProductID
    
    If Rs1!Quantity > 0 Then
    
    Set LVITEM = .ListItems.Add(, , Rs1!ProductID)
        LVITEM.SubItems(2) = Rs1!Quantity
        LVITEM.SubItems(3) = Rs1!Rate
        LVITEM.SubItems(1) = Rs1!Name
        LVITEM.SubItems(4) = CLng(Rs1!Quantity) * CLng(Rs1!Rate)
    End If
    
    End With
    Rs1.MoveNext
Wend
calculate
Set Rs1 = Nothing
Set DCON = Nothing
TXTNUM.text = MODIFYID

cmdClear.Enabled = True
cmdOk.Enabled = True
Command1.Enabled = False
End Sub

Private Sub cmbCust_Click()
    Call CMB2("Select Distinct [date] from PurRet where BusinessName='" & cmbCust & "'", cmbdate)
    
    
    
    Call CMB2("Select Distinct PurchaseRegistryID from PurRet where BusinessName='" & cmbCust & "'", cmbreg)

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select suppliersid from PurRet where BusinessName='" & cmbCust & "'")
If Rs1.RecordCount <> 0 Then
SuppID = Rs1!suppliersid
End If
Set Rs1 = Nothing
Set DCON = Nothing

If cmbCust.text <> "" Then
    cmbdate.Enabled = True
Else
    cmbdate.Enabled = False
End If


End Sub

Private Sub cmbdate_Click()
    Call CMB2("Select Distinct PurchaseRegistryID from PurRet where BusinessName='" & cmbCust & "' And [date]=#" & cmbdate & "#", cmbreg)

If cmbdate.text <> "" Then
    cmbreg.Enabled = True
Else
    cmbreg.Enabled = False
End If

End Sub

Private Sub cmbreg_Click()

 SQL = "Select * from PurRet where (businessname='" & cmbCust & "' and PurchaseRegistryId='" & cmbreg & "') AND Quantity>0"
 Call GetNewConnection2
 Set Rs1 = New Recordset
 Set Rs1 = DCON.Execute(SQL)
   lvMain.ListItems.clear
    
While Not Rs1.EOF
    Dim LVITEM As ListItem
    Dim ProdID As String
    With lvMain
    ProdID = Rs1!ProductID
    
    If Rs1!Quantity > 0 Then
    
    Set LVITEM = .ListItems.Add(, , Rs1!ProductID)
        LVITEM.SubItems(2) = Rs1!Quantity
        LVITEM.SubItems(3) = Rs1!Rate
        LVITEM.SubItems(1) = Rs1!Name
        LVITEM.SubItems(4) = CLng(Rs1!Quantity) * CLng(Rs1!Rate)
    End If
    
    End With
    Rs1.MoveNext
Wend
calculate
Set Rs1 = Nothing
Set DCON = Nothing

If cmbreg.text <> "" Then
    txtQTY.Enabled = True
Else
    txtQTY.Enabled = False
End If

End Sub
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

Private Sub cmdAdd_Click()


Dim LSTITEM As ListItem



Dim CNT As Boolean
Dim dd As Integer
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
          
            For dd = 1 To .ListItems.Count
               
                If InStr(1, .ListItems(dd).text, Rs1!ProductID) = 1 Then
                    If InStr(1, .ListItems(dd).SubItems(1), Rs1!Name) = 1 Then
              
                        If EDT = True Then
                            .ListItems(dd).Selected = True
                            .ListItems(dd).SubItems(2) = Val(txtQTY.text)
                            .ListItems(dd).SubItems(3) = Val(txtRate.text)
                            .ListItems(dd).SubItems(4) = text5.text
                        Else
                            .ListItems(dd).Selected = True
                            .ListItems(dd).SubItems(2) = Val(.ListItems(dd).SubItems(2)) + Val(txtQTY.text)
                            .ListItems(dd).SubItems(3) = Val(txtRate.text)
                            .ListItems(dd).SubItems(4) = Val(.ListItems(dd).SubItems(2)) * Val(.ListItems(dd).SubItems(3))
                    
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
          '  .ListItems(.ListItems.Count).SubItems(4) = "1"
            .ListItems(.ListItems.Count).SubItems(4) = text5.text
             
             
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


            For dd = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(dd).SubItems(5)) + Val(TextAmount.text)
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


Else



txtQTY.SetFocus


End If

End Sub

Private Sub cmdAdd_GotFocus()

If Val(txtQTY.text) <> 0 Then
  Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * From PurchaseRegistryDetail where PurchaseRegistryID='" & cmbreg & "' and ProductID='" & Text4 & "'")
     If Rs1.RecordCount <> 0 Then
     
     
        If Val(Rs1!Quantity) < Val(txtQTY.text) Then
            MsgBox "The Quantity that you want to return " & vbCrLf _
            & "is greater than Quantity you Buy", vbInformation
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

Private Sub cmdClear_Click()
EDT = False
lvMain.ListItems.clear

picTop.Enabled = False


lvMain.Enabled = False
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""



cmdOk.Enabled = False
cmdClear.Enabled = False
Command1.Enabled = True
TXTNUM.text = ""
cmbCust.ListIndex = -1

cmbreg.Enabled = False
cmbCust.Enabled = False
cmbdate.Enabled = False


End Sub

Private Sub cmdOk_Click()
If lvMain.ListItems.Count > 0 Then
cmdClear.Enabled = False
cmdOk.Enabled = False
Command1.Enabled = True
picTop.Enabled = False

lvMain.Enabled = False

cmbreg.Enabled = False
cmbCust.Enabled = False
cmbdate.Enabled = False

'If ADDING = True Then

Call PurchaseReturn
'Else
'Call UpdateReturn
'End If

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


ReturnID = CIns.AUTONUM(CDb.OpenDb, "PurchaseOrderHeader", "PurchaseOrderID", "PRet", TXTNUM)


Set CIns = Nothing



''Call CMB1("PurchaseRegistryHeader", "PurchaseRegistryID", Combo5, , True)

cmdClear.Enabled = True
cmdOk.Enabled = True
Command1.Enabled = False
picTop.Enabled = True

lvMain.Enabled = True

lvMain.ListItems.clear

'cmbreg.Enabled = True
cmbCust.Enabled = True
'cmbdate.Enabled = True

Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""


End Sub



Private Sub Form_Load()

'If ADDING = False Then
'    Call ReturnModify
'      Call CMB2("Select distinct BusinessName from PurRet", cmbCust)
'
'Else
'  cmbCust.ListIndex = -1
    Call CMB2("Select distinct BusinessName from PurRet", cmbCust)
'End If
    
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
Dim dd As Integer

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

            For dd = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(dd).SubItems(5)) + Val(TextAmount.text)
            Next
 End If
End If
End With
End Sub

Private Sub txtqty_Change()

    text5.text = Val(txtQTY.text) * Val(txtRate.text)


If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
Call OFFCHar(KeyAscii, txtQTY)
End Sub

Private Sub txtQTY_LostFocus()
If Len(txtQTY.text) = 0 Or Val(txtQTY.text) = 0 Then
    cmdAdd.Enabled = False
Else
    
    cmdAdd.Enabled = True
End If
End Sub

