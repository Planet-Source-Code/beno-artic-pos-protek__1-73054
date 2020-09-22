VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLookup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProsVent Inventory Manager 2005"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3780
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   720
      Width           =   315
   End
   Begin MSForms.ComboBox cmbCategory 
      Height          =   330
      Left            =   945
      TabIndex        =   8
      Top             =   2580
      Width           =   2655
      VariousPropertyBits=   748701723
      DisplayStyle    =   3
      Size            =   "4683;582"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2670
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Stokc Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   885
   End
   Begin MSForms.ComboBox cmbStockLevel 
      Height          =   330
      Left            =   945
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4683;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblStatus 
      Height          =   495
      Left            =   165
      TabIndex        =   3
      Top             =   8085
      Width           =   2055
      VariousPropertyBits=   8388627
      Size            =   "3625;873"
      FontName        =   "Verdana"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Searh For Label"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F9F0EB&
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1905
   End
   Begin VB.Image imgTop 
      Height          =   600
      Left            =   0
      Picture         =   "frmLookup.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3825
   End
   Begin MSForms.TextBox txtProdName 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   690
      Width           =   2415
      VariousPropertyBits=   746604571
      MaxLength       =   25
      BorderStyle     =   1
      Size            =   "4260;661"
      BorderColor     =   -2147483647
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   750
      Width           =   450
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000003&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   570
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   3585
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub Check2_Click()
'If LocalBool = False Then
'    Call dbCmbBind("Select Distinct Category From V_Prod_Stat Where Category IS NOT NULL order by Category", cmbCategory, "<ALL>")
'    Call dbCmbBind("Select Distinct Location From V_Prod_Stat Where Location IS NOT Null", cmblocation, "<ALL>")
'    Call dbCmbBind("Select Distinct Prod_Stat From V_Prod_Stat", cmbStockLevel, "<ALL>")
'    LocalBool = True
'End If
'    Frame1.Visible = Check2.Value
'    If Check2.Value Then
'        ListView1.Top = ListView1.Top + Frame1.Height
'        ListView1.Height = ListView1.Height - Frame1.Height
'    Else
'        ListView1.Top = Frame1.Top
'        ListView1.Height = ListView1.Height + Frame1.Height
'        cmbCategory.ListIndex = 0
'        cmblocation.ListIndex = 0
'        cmbStockLevel.ListIndex = 0
'        End If
'End Sub

'Private Function Search(ParamArray arguments() As Variant) As String
'On Error GoTo adder:
'Dim i As Integer
'Dim counts As Variant
'Dim strSQL As String
'    For i = LBound(arguments) To UBound(arguments)
'        If arguments(i) = "<ALL>" Then
'            arguments(i) = ""
'        End If
'   Next i
'
'    If IsNumeric(cmbTopCount.text) = True Then
'        counts = CInt(cmbTopCount.text)
'    Else
'        counts = "100 PERCENT"
'    End If
'Dim txtSearchMain As String
'txtSearchMain = Replace(txtProdName, "'", Empty, 1, Len(txtProdName), vbTextCompare)
'strSQL = "SELECT TOP " & counts & " ProductID As [Product ID], " _
'        & " Name As Name,TotalStock " _
'        & " From V_Prod_Stat  WHERE " _
'        & "((Name LIKE '" & txtSearchMain & "%')  " _
'        & " OR (ProductID LIKE '" & txtSearchMain & "%'))  " _
'        & " AND (Location LIKE '" & arguments(1) & "%' Or Location IS Null)" _
'        & " AND (Category LIKE '" & arguments(0) & "%')  " _
'        & " AND (Prod_Stat Like '" & arguments(2) & "%') and Discontinued=" & Check1.Value
'
'    Dim tempRs As New ADODB.Recordset
'    Connect
'    ListView1.ListItems.clear
'    Call tempRs.Open(strSQL, DCON, adOpenForwardOnly, adLockReadOnly)
'    Dim lst As ListItem
'
'    While Not tempRs.EOF
'        Set lst = ListView1.ListItems.Add(, , tempRs.Collect(0))
'                      lst.SubItems(1) = tempRs.Collect(1)
'
'                    If tempRs.Collect(2) <= 0 Or IsNull(tempRs.Collect(2)) = True Then
'                        lst.ForeColor = vbRed
'                        lst.ListSubItems(1).ForeColor = vbRed
'                    End If
'        tempRs.MoveNext
'    Wend
'    Set tempRs = Nothing
'    lblStatus.Caption = ListView1.ListItems.Count
'Exit Function
'adder:
'    Exit Function
'
'End Function

Private Sub Command1_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from Product where Producid='" & txtProdName & "' And name like '" & txtProdName & "%'")

If Rs1.RecordCount <> 0 Then
    
    
End If

End Sub
