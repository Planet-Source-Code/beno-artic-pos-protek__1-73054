VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Izbor 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Iskalnik"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6900
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3840
      Picture         =   "Izbor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   270
      Left            =   120
      MaskColor       =   &H8000000F&
      Picture         =   "Izbor.frx":018E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   270
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Izberi"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   3240
      Width           =   885
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
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   885
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Vnesi iskalni niz"
      Top             =   3240
      Width           =   3075
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
      Left            =   3360
      Picture         =   "Izbor.frx":0288
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2040
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Izbor.frx":0812
      Height          =   1935
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1060
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1060
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   0
      X2              =   6840
      Y1              =   3120
      Y2              =   3120
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
      Left            =   0
      TabIndex        =   4
      Top             =   100
      Width           =   6825
   End
   Begin VB.Line line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   270
      X2              =   6150
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Shape shpMB 
      BorderColor     =   &H00926747&
      Height          =   3825
      Left            =   0
      Top             =   0
      Width           =   6945
   End
   Begin VB.Shape shpCapBor 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   70
      Top             =   120
      Width           =   6705
   End
End
Attribute VB_Name = "Izbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ssqq As String
Dim xdata As String
Dim poll As String

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdFilter_Click()
 Dim sqs As String
    If txtFilter.text = "Vnesi iskalni niz" Then
    'Call GRIDBINDx(Me.DataGrid1, ssqq)
    Adodc1.RecordSource = ssqq

 Adodc1.Refresh

    Else
    sqs = ssqq & " where " & poll & " like('%" & txtFilter.text & "%')"
    '   Call GRIDBINDx(Me.DataGrid1, sqs)
    Adodc1.RecordSource = sqs

 Adodc1.Refresh

    End If
     LoadFlexGridColumnWidths DataGrid1, ssqq
End Sub

Private Sub Command2_Click()
If DataGrid1.BackColor = 255 Then
DataGrid1.AllowUpdate = False
DataGrid1.SetFocus
DataGrid1.BackColor = &H80000005

Else
DataGrid1.AllowUpdate = True
DataGrid1.SetFocus
DataGrid1.BackColor = 255
End If
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
        If DataGrid1.Columns(0).text <> "" Then
        If DataGrid1.Columns(0).text = "CASH" Then
            frmMAIN.mnuEdit.Enabled = False
        Else
            DataGrid1.SetFocus
            frmMAIN.mnuEdit.Enabled = True
            frmMAIN.mnuModify.Enabled = True
            PopupMenu frmMAIN.mnuEdit
        End If
        End If
    End If
    
End Sub




Private Sub Form_Unload(cancel As Integer)
'UnHook
End Sub
 Public Sub MouseWheel(ByVal fwKeys As Long, ByVal zDelta As Long, ByVal Xpos As Long, _
    ByVal Ypos As Long)

   'put a label on your for to check changing values
   ' Label1.Caption = "Keys=" & fwKeys & " Delta=" & zDelta & " xPos=" & Xpos & " yPos=" & Ypos
   If zDelta > 0 Then
   SendKeys "{up}"
   Else
   SendKeys "{down}"
   End If
'then you can change toprow of flex grid accordingly
End Sub
Private Sub LoadFlexGridColumnWidths(ByVal flx As DataGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Columns.Count - 1
        ' Get the column width. Use its current
        ' width as the default value.
        flx.Columns(i).Width = GetSetting( _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.Columns(i).Width)
    Next i
End Sub
Private Sub SaveFlexGridColumnWidths(ByVal flx As DataGrid, kira As String)
Dim i As Integer

    For i = 0 To flx.Columns.Count - 1
        ' Save the column width.
        SaveSetting _
            kira, _
            "ColumnWidths", "Col" & Format$(i), _
            flx.Columns(i).Width
    Next i
End Sub

Private Sub cmdSelect_Click()
'Parent.Parent.txtDisplay.text
'mBoundDatax = DataGrid1.Columns(0).text
 xdata = DataGrid1.Columns(poll).text
  '  mDisplayData = DataGrid1.Columns(0).text
    'Frm.Ctl.txtDisplay.text = DataGrid1.Columns(0).text
Unload Me
End Sub

Private Sub Command1_Click()
  SaveFlexGridColumnWidths DataGrid1, ssqq

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
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path + "\database\thesis.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password="


Adodc1.RecordSource = ssqq

 Adodc1.Refresh

 
 LoadFlexGridColumnWidths DataGrid1, ssqq
 
 
 
End Sub
Public Function odpri(levo As Integer, zgoraj As Integer, ssql As String, ByRef sBoundData As String, ByRef sDisplayData As String, ByRef ctl As UserControl1, ByRef frm As Form, ByRef polje As String) As String
'Me.Top = zgoraj
'me.Top=ctl.
'Me.Left = levo
ssqq = ssql
poll = polje
Me.Show vbModal

odpri = xdata
End Function

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
    
    DataGrid1.Move 2, _
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
    Command1.Top = DataGrid1.Top
    Command1.Left = DataGrid1.Left
  ' MsgBox (cmdCancel.Top)
  '  Me.Width = lblCaption.Width
  '  Me.Height = lblCaption.Height + DataGrid1.Height + cmdCancel.Height
    err.clear
End Sub


