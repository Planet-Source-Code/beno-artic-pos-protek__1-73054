VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form brow_nt 
   BackColor       =   &H00FFC0C0&
   Caption         =   "brow_nt"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form7"
   ScaleHeight     =   6210
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Dodaj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   8
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton cmdFilter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      Picture         =   "brow_nt.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   4
      Text            =   "Vnesi iskalni niz"
      Top             =   5160
      Width           =   3075
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Preklici"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   3
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Uredi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      TabIndex        =   2
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Height          =   270
      Left            =   0
      MaskColor       =   &H8000000F&
      Picture         =   "brow_nt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   270
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      Picture         =   "brow_nt.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DATAGRID1 
      DragIcon        =   "brow_nt.frx":0812
      Height          =   3840
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      ForeColorSel    =   16777215
      BackColorUnpopulated=   16777152
      GridColor       =   12632256
      GridColorFixed  =   16777215
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLines       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin VB.Shape shpCapBor 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   75
      Top             =   720
      Width           =   9705
   End
   Begin VB.Shape shpMB 
      BorderColor     =   &H00926747&
      Height          =   5145
      Left            =   0
      Top             =   600
      Width           =   10185
   End
   Begin VB.Line line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   630
      X2              =   6510
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iskalnik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   9945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   360
      X2              =   9960
      Y1              =   5040
      Y2              =   5040
   End
End
Attribute VB_Name = "brow_nt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim ssqq As String
Dim xdata As String
Dim poll As Integer
Dim po As String


Private Sub cmdCancel_Click()
'MsgBox ssqq
Unload Me

End Sub

Private Sub cmdFilter_Click()
 Dim sqs As String
    If txtFilter.text = "Vnesi iskalni niz" Then
    'Call GRIDBINDx(Me.DataGrid1, ssqq)
   If RS.State = 1 Then RS.Close
RS.Open ssqq, myConection, adOpenDynamic, adLockOptimistic

Set DATAGRID1.DataSource = RS


    Else
    sqs = ssqq & " where " & poll & " like('%" & txtFilter.text & "%')"
    '   Call GRIDBINDx(Me.DataGrid1, sqs)
    If RS.State = 1 Then RS.Close
RS.Open sqs, myConection, adOpenDynamic, adLockOptimistic

Set DATAGRID1.DataSource = RS

 

    End If
    LoadFlexGridColumnWidths DATAGRID1, ssqq
End Sub

Private Sub Command3_Click()
tip_dok = "NT"
ma_ured = 0
Dim xxx As Integer
xxx = Val(Right(RTrim(Getnazi("select max(id_dok) as dd from nabasif where tip_dok='NT' and id_dok like '" & po & "%'")), 3)) + 1
normati = "NT" & RTrim(po) & "-" & novast(LTrim(Str(xxx)), 3)
frmblag.Show vbModal
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 
    
End Sub




Private Sub Form_Unload(cancel As Integer)
'UnHook
End Sub

Private Sub LoadFlexGridColumnWidths(ByVal flx As MSHFlexGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Cols - 1
        ' Get the column width. Use its current
        ' width as the default value.
        flx.ColWidth(i) = GetSetting( _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.ColWidth(i))
    Next i
End Sub
Private Sub SaveFlexGridColumnWidths(ByVal flx As MSHFlexGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Cols - 1
        ' Save the column width.
        SaveSetting _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.ColWidth(i)
    Next i
End Sub
Private Sub cmdSelect_Click()
tip_dok = "NT"
ma_ured = 0
normati = "NT" & Me.DATAGRID1.text
frmblag.Show vbModal
End Sub

Private Sub Command1_Click()
  SaveFlexGridColumnWidths DATAGRID1, ssqq

End Sub

Private Sub DataGrid1_DblClick()
cmdSelect_Click
End Sub

Private Sub Form_Activate()
'If myConection.State = adStateOpen Then
'Else
'myConection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path + "\DATABASE\Thesis.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password="
'End If
'If Rs1.State = 1 Then Rs1.Close

'Rs1.Open ssqq, myConection, adOpenStatic, adLockOptimistic
'MsgBox (ssqq)
'Filll List87, ssqq
'Set DataGrid1.DataSource = Rs1
'
'Call GRIDBINDx(Me.DataGrid1, ssqq)
 
 'PathFileName = App.path + "\baza.mdb"
'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path + "\database\thesis.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password="
If RS.State = 1 Then RS.Close
RS.Open ssqq, myConection, adOpenDynamic, adLockOptimistic
Set DATAGRID1.DataSource = RS
 DATAGRID1.Redraw = True


 LoadFlexGridColumnWidths DATAGRID1, ssqq
 
 
 
End Sub
Public Function odpri(sSQL As String, ByRef polje As String)
'Me.Top = zgoraj
'me.Top=ctl.
'Me.Left = levo
ssqq = sSQL
po = polje
    
Me.Show vbModal

odpri = xdata
End Function

Private Sub LaVolpeButton1_Click()
tip_dok = "NT"
ma_ured = 0
normati = Me.DATAGRID1.text
frmblag.Show vbModal
End Sub

Private Sub Timer1_Timer()
'If FR = "" Then

'FR = "1"
'End If
End Sub

Private Sub txtFilter_GotFocus()
    If txtFilter.text = "Vnesi iskalni niz" Then
        txtFilter.text = ""
    End If
End Sub

Private Sub txtFilter_LostFocus()
    If Len(Trim(txtFilter.text)) < 1 Then
        txtFilter.text = "Vnesi iskalni niz"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpMB.Move 10, 0, Me.ScaleWidth, Me.ScaleHeight
    lblCaption.Move 10, lblCaption.Top, Me.ScaleWidth
    shpCapBor.Move 10, 0, Me.ScaleWidth
    
    DATAGRID1.Move 2, _
                    shpCapBor.Top + shpCapBor.Height + 1, _
                    Me.ScaleWidth - 4, _
                    Me.ScaleHeight - (shpCapBor.Top + shpCapBor.Height + 1 + cmdCancel.Height + 8) - 100
    line2.X1 = 10
    line2.X2 = Me.ScaleWidth
    line2.Y1 = listEntries.Top + listEntries.Height + 1
    line2.Y2 = line2.Y1
    
    txtFilter.Move 4, Me.ScaleHeight - txtFilter.Height - 3, Me.ScaleWidth / 2
    cmdFilter.Move txtFilter.Left + txtFilter.Width + 2, txtFilter.Top
     Command2.Move txtFilter.Left + txtFilter.Width + 2 + cmdFilter.Width, txtFilter.Top
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 3, Me.ScaleHeight - cmdCancel.Height - 3
    cmdSelect.Move Me.ScaleWidth - cmdCancel.Width - cmdSelect.Width - 3, Me.ScaleHeight - cmdSelect.Height - 3
      Command3.Move Me.ScaleWidth - cmdSelect.Width - cmdCancel.Width - Command3.Width - 3, Me.ScaleHeight - Command3.Height - 3
    Command1.Top = DATAGRID1.Top
    Command1.Left = DATAGRID1.Left
  ' MsgBox (cmdCancel.Top)
  '  Me.Width = lblCaption.Width
  '  Me.Height = lblCaption.Height + DataGrid1.Height + cmdCancel.Height
    err.clear
End Sub




