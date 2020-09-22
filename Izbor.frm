VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Izbor 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Iskalnik"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9720
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   510
   End
   Begin LVbuttons.LaVolpeButton uredi 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Uredi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Izbor.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton dodaj 
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Dodaj"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Izbor.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   4800
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
      Picture         =   "Izbor.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   270
      Left            =   120
      MaskColor       =   &H8000000F&
      Picture         =   "Izbor.frx":01C6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
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
      Left            =   7800
      TabIndex        =   4
      Top             =   4560
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
      Left            =   9120
      TabIndex        =   3
      Top             =   4560
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
      Left            =   360
      TabIndex        =   2
      Text            =   "Vnesi iskalni niz"
      Top             =   4560
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
      Left            =   3600
      Picture         =   "Izbor.frx":02C0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DATAGRID1 
      DragIcon        =   "Izbor.frx":084A
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   360
      X2              =   9960
      Y1              =   4440
      Y2              =   4440
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
      TabIndex        =   5
      Top             =   120
      Width           =   9465
   End
   Begin VB.Line line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   630
      X2              =   6510
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Shape shpMB 
      BorderColor     =   &H00926747&
      Height          =   5145
      Left            =   0
      Top             =   0
      Width           =   10185
   End
   Begin VB.Shape shpCapBor 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   75
      Top             =   120
      Width           =   9705
   End
End
Attribute VB_Name = "Izbor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim ssqq As String

Private Const MARGIN_SIZE = 60      ' in Twips
' variables for data binding


' variables for enabling column sort
Private m_iSortCol As Integer
Private m_iSortType As Integer


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

Private Sub Combo1_LostFocus()
On Error GoTo bvc:
If Getnazi("select id_dok from dokm where atribut='ODPO' and tekst='" & ssqq & "'") <> "" Then
myConection.Execute ("delete from dokm where atribut='ODPO' and tekst='" & ssqq & "'")
End If
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('ODPO','" & Me.Combo1.text & "','" & ssqq & "')")
bvc:
End Sub

Private Sub DATAGRID1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call cmdSelect_Click
End If
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
 
    
End Sub




Private Sub dodaj_Click()
If Me.Combo1.text <> "" Then
Dim ress As Boolean
Dim odko As String
Dim pazi_id
pazi_id = MODIFYID
MODIFYID = ""
odko = Trim(Getnazi("select tekst from dokm where atribut='FORM' and id_dok='" & Me.Combo1.text & "'")) & ".show vbmodal"
ress = FExecuteCode(odko)
  reff
 MODIFYID = pazi_id
Else
MsgBox "Combo z formami je prazen vnesi ime forme!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
'Parent.Parent.txtDisplay.text
'mBoundDatax = DataGrid1.Columns(0).text
For i = 0 To DATAGRID1.Cols - 1
'MsgBox UCase(PO)
      If UCase(Trim(DATAGRID1.TextMatrix(0, i))) = UCase(po) Then
      DATAGRID1.Col = i
      poll = DATAGRID1.Col
      End If
     Next i
 xdata = DATAGRID1.TextMatrix(DATAGRID1.Row, poll)
 xizb = xdata
  '  mDisplayData = DataGrid1.Columns(0).text
    'Frm.Ctl.txtDisplay.text = DataGrid1.Columns(0).text
Unload Me
End Sub

Private Sub Command1_Click()
  SaveFlexGridColumnWidths DATAGRID1, ssqq

End Sub

Private Sub DataGrid1_DblClick()
 Dim i As Integer

    ' sort only when a fixed row is clicked
    If DATAGRID1.MouseRow < DATAGRID1.FixedRows Then

    i = m_iSortCol                  ' save old column
    m_iSortCol = DATAGRID1.Col   ' set new column

    ' increment sort type
    If i <> m_iSortCol Then
        ' if clicking on a new column, start with ascending sort
        m_iSortType = 1
    Else
        ' if clicking on the same column, toggle between ascending and descending sort
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort
    Else
    Call cmdSelect_Click
    End If





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
On Error GoTo bee:
Call CMB1("dokm", "id_dok", Combo1, "where atribut='FORM'")
Me.Combo1.text = ""
If ssqq Like "*from partner*" Then
Me.Combo1.text = "partner"
End If
If ssqq Like "*from zaposleni*" Then
Me.Combo1.text = "zaposleni"
End If
If ssqq Like "*from tip_art*" Then
Me.Combo1.text = "tip_art"
End If
If ssqq Like "*from em*" Then
Me.Combo1.text = "merske"
End If
If ssqq Like "*from grupa*" Then
Me.Combo1.text = "grupe"
End If
If ssqq Like "*from skla*" Then
Me.Combo1.text = "skladisce"
End If
If Me.Combo1.text = "" Then
Me.Combo1.text = Getnazi("select id_dok from dokm where tekst='" & ssqq & "' and atribut='ODFO'")
End If

If Left(ssqq, 4) = "sele" Then
If RS.State = 1 Then RS.Close
RS.Open ssqq, myConection, adOpenDynamic, adLockOptimistic
Set DATAGRID1.DataSource = RS
 DATAGRID1.Redraw = True


 LoadFlexGridColumnWidths DATAGRID1, ssqq
 Exit Sub
 Else
 Unload Me
 End If
 
bee:
 Unload Me
End Sub
Public Function odpri(levo As Integer, zgoraj As Integer, sSQL As String, ByRef sBoundData As String, ByRef sDisplayData As String, ByRef ctl As UserControl1, ByRef frm As Form, ByRef polje As String) As String
'Me.Top = zgoraj
'me.Top=ctl.
'Me.Left = levo
ssqq = sSQL
po = polje
    
Me.Show vbModal

odpri = xdata
End Function

Private Sub Timer1_Timer()
If FR = "" Then

FR = "1"
End If
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
    Me.Combo1.Top = Me.lblCaption.Top
    shpCapBor.Move 10, 0, Me.ScaleWidth
    
    DATAGRID1.Move 2, _
                    shpCapBor.Top + shpCapBor.Height + 1, _
                    Me.ScaleWidth - 4, _
                    Me.ScaleHeight - (shpCapBor.Top + shpCapBor.Height + 1 + cmdCancel.Height + 8) - 100
    line2.X1 = 10
    line2.X2 = Me.ScaleWidth
    line2.Y1 = ListEntries.Top + ListEntries.Height + 1
    line2.Y2 = line2.Y1
    
    txtFilter.Move 4, Me.ScaleHeight - txtFilter.Height - 3, Me.ScaleWidth / 2
    cmdFilter.Move txtFilter.Left + txtFilter.Width + 2, txtFilter.Top
     Command2.Move txtFilter.Left + txtFilter.Width + 2 + cmdFilter.Width, txtFilter.Top
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 3, Me.ScaleHeight - cmdCancel.Height - 3
    cmdSelect.Move Me.ScaleWidth - cmdCancel.Width - cmdSelect.Width - 3, Me.ScaleHeight - cmdSelect.Height - 3
    dodaj.Move txtFilter.Left + txtFilter.Width + 2 + cmdFilter.Width + Command2.Width, txtFilter.Top
    UREDI.Move txtFilter.Left + txtFilter.Width + 2 + cmdFilter.Width + Command2.Width + dodaj.Width, txtFilter.Top
    Command1.Top = DATAGRID1.Top
    Command1.Left = DATAGRID1.Left
  ' MsgBox (cmdCancel.Top)
  '  Me.Width = lblCaption.Width
  '  Me.Height = lblCaption.Height + DataGrid1.Height + cmdCancel.Height
    err.clear
End Sub


Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------

    With DATAGRID1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType

        .FillStyle = flexFillRepeat
        .Col = 0
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
      '  .CellBackColor = &HFFFFFF
        ' grey every other row
        Dim iLoop As Integer
      
       
        For iLoop = .FixedRows To .Rows - 1
        Dim asx As String
        asx = DATAGRID1.TextMatrix(iLoop, 1)
       
         .Row = iLoop
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
           ' MsgBox asx
            'MsgBox (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'"))
        Next iLoop
        .FillStyle = flexFillSingle

        .Redraw = True
        
    End With

End Sub

Private Sub UREDI_Click()
If Me.Combo1.text <> "" Then
For i = 0 To DATAGRID1.Cols - 1
'MsgBox UCase(PO)
      If UCase(Trim(DATAGRID1.TextMatrix(0, i))) = UCase("SIFRA") Then
      DATAGRID1.Col = i
      poll = DATAGRID1.Col
      End If
     Next i
     Dim xdatta
 xdatta = DATAGRID1.TextMatrix(DATAGRID1.Row, poll)
 
'If ssqq Like "*from partner*" Then
Dim pazi_id
pazi_id = MODIFYID
MODIFYID = xdatta

  Dim ress As Boolean
Dim odko As String
odko = Trim(Getnazi("select tekst from dokm where atribut='FORM' and id_dok='" & Me.Combo1.text & "'")) & ".show vbmodal"
ress = FExecuteCode(odko)
   

    MODIFYID = pazi_id
    reff
Else
MsgBox "Combo z formami je prazen vnesi ime forme!"
End If
'End If
End Sub
Private Sub reff()
If Left(ssqq, 4) = "sele" Then
If RS.State = 1 Then RS.Close
RS.Open ssqq, myConection, adOpenDynamic, adLockOptimistic
Set DATAGRID1.DataSource = RS
 DATAGRID1.Redraw = True
End If

End Sub
