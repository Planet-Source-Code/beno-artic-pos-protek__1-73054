VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form potni 
   Caption         =   "Vnos potnih"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form7"
   ScaleHeight     =   7260
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin ProsVent.UserControl1 UserControl13 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      polje           =   "relacija"
      ssql            =   "select * from relacije"
   End
   Begin ProsVent.UserControl1 UserControl12 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      polje           =   "avto"
      ssql            =   "select * from avto"
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      polje           =   "ime"
      ssql            =   "select * from zaposleni"
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Row"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete Row"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Data"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFetch 
      Caption         =   "Fetch Data"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtNewData 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdUnDel 
      Caption         =   "UnDelete"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   6360
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TrialGrd"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=TrialGrd"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "GrdData"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgtrial 
      DragIcon        =   "potni.frx":0000
      Height          =   3840
      Left            =   -120
      TabIndex        =   6
      Top             =   2400
      Width           =   11760
      _ExtentX        =   20743
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
         Size            =   9.75
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
End
Attribute VB_Name = "potni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim cn As New ADODB.Connection

Dim ftchFlag As Boolean     ' 0 - default 1 - add,  2 - modi, +3 - del
Dim adRwFlag As Boolean
Dim edRwFalg As Boolean
Dim svRwFlag As Boolean

Dim edCol As Integer
Dim curCol As Integer
Dim curRow As Integer
Dim msgFlag As Boolean

Dim clk As Boolean

Dim st As String

Private Sub cmdAdd_Click()
    Dim lstRow As Integer
If RS.State = 1 Then RS.Close
   RS.Open "select * from trenutna", myConection, adOpenDynamic, adLockOptimistic
   Dim xro As Integer
   xro = 1
   Do While xro = fgtrial.Rows - 1
   RS.AddNew
       For i = fgtrial.FixedCols To fgtrial.Cols - 1
          RS.Fields(Trim(fgtrial.TextMatrix(0, i))) = fgtrial.TextMatrix(xro, i)
      Next i
    RS.Update
      xro = xro + 1
    Loop
    txtNewData.Visible = False
        txtNewData.text = ""
        If fgtrial.TextMatrix(fgtrial.Rows - 1, 1) = "" Then
        Else
        fgtrial.Rows = fgtrial.Rows + 1
        End If
    lstRow = fgtrial.Rows - 1
    fgtrial.Row = lstRow
    fgtrial.Col = 1
     fgtrial.TextMatrix(lstRow, 0) = lstRow
     fgtrial.TextMatrix(lstRow, 1) = ""
     fgtrial.SetFocus
End Sub


Private Sub cmdSave_Click()
    If svRwFlag = True Then
        Dim id As String
        Dim fld As String
        Dim dt As String
        Dim ftFg As Integer
                
        'Open "E:\VishwaPrg\Rohini\all_vb_prog\TrialGrid\trialSqul.sql" For Output As FreeFile
        For i = 1 To fgtrial.Rows - 1
            fgtrial.Row = i
            Dim rw As Integer
            Cols = fgtrial.Cols - 1
            
            fgtrial.Col = 0
            id = fgtrial.text
            dt = ""
            
            For k = 1 To Cols - 1
                fgtrial.Col = k
                dt = dt & "," & fgtrial.text
            Next k
            
            fgtrial.Col = k
            ftFg = Val(fgtrial.text)
        
            Select Case ftFg
                Case 0
                    MsgBox "Fetched" & " - " & Mid(dt, 2, Len(dt))
                Case 1
                    MsgBox "Added" & " - " & Mid(dt, 2, Len(dt))
                Case 2
                    MsgBox "Modified" & " - " & Mid(dt, 2, Len(dt))
                Case Else
                    MsgBox "Dele" & " - " & dt
            End Select
        Next i
        'Close
        svRwFlag = False
    Else
        MsgBox "Data Allready saved"
    End If
End Sub

Private Sub cmdUnDel_Click()
    For i = 1 To fgtrial.Cols - 1
        fgtrial.Col = i
        fgtrial.CellForeColor = vbBlack
    Next i
    
    If Val(fgtrial.text) > 2 Then
        fgtrial.text = Val(fgtrial.text) - 3
    Else
        MsgBox "The Record Is Not Deleted"
    End If
    
    fgtrial.Col = curCol
End Sub

Private Sub cmdDel_Click()
    For i = 1 To fgtrial.Cols - 1
        fgtrial.Col = i
        fgtrial.CellForeColor = vbRed
    Next i
     If Val(fgtrial.text) < 3 Then
        fgtrial.text = Val(fgtrial.text) + 3
    Else
        MsgBox "The Record Is Already Deleted"
    End If
    
    fgtrial.Col = curCol
End Sub

Private Sub cmdFetch_Click()
    If RS.State = 0 And ftchFlag = False Then
        ftchFlag = True
      '  rs.Source = "Select * from utna"
        RS.Open "select pozicija,sifra,naziv,cena,kol,znes,x,y from trenutna", myConection, adOpenDynamic, adLockOptimistic
        'Set fgTrial.DataSource = rs.Source
        
        If RS.EOF = False Then
            RS.MoveFirst
            For i = 1 To RS.RecordCount - 1
                If i <> fgtrial.Rows - 1 Then
                    fgtrial.Rows = fgtrial.Rows + 1
                End If
                fgtrial.Row = i
                For J = 0 To fgtrial.Cols - 2
                    fgtrial.Col = J
                    fgtrial.text = RS.Fields(J)
                Next J
                ' to set last col as fetch flag - 0
                fgtrial.Col = J
                fgtrial.text = 0
                RS.MoveNext
            Next i
        End If
    End If
End Sub

Private Sub fgTrial_Click()
    clk = False
    curRow = fgtrial.Row
    curCol = fgtrial.Col
    msgFlag = False
End Sub

Private Sub fgTrial_DblClick()
    clk = True
    
    
    If ftchFlag = True And adRwFlag = False Then
        edRwFalg = True
    End If
    
    curCol = 1
    curRow = fgtrial.Rows - 1
    fgTrial_KeyPress (0)
End Sub

Private Sub fgTrial_KeyPress(KeyAscii As Integer)
    Dim tmpCol As Integer
    'clk = True
If KeyAscii = 13 And fgtrial.Col <> coollsi Then
    If fgtrial.Col < fgtrial.Cols - 1 And fgtrial.Col <> coollko Then
      fgtrial.Col = fgtrial.Col + 1
      Else
       txtNewData.Visible = False
       
       
      cmdAdd_Click
     ' MsgBox fgtrial.Col
      'fgtrial.Col = coollsi
      End If
      
Else
    tmpCol = fgtrial.Col
    'fgtrial.Col = fgtrial.Cols - 1
        curRow = fgtrial.Row
       ' If adRwFlag = True Then
       '     curCol = 1
       ' Else
            curCol = fgtrial.Col
        'End If
       ' MsgBox curRow
        'MsgBox Chr(KeyAscii')
        
     '   fgtrial.Col = curCol
        
        txtNewData.text = Chr(KeyAscii)
        txtNewData.Move fgtrial.CellLeft + fgtrial.Left, fgtrial.CellTop + _
                        fgtrial.Top, fgtrial.CellWidth, fgtrial.CellHeight
        
        'to set col no to previous value
        txtNewData.Visible = True
        txtNewData.SetFocus
    'Else
    '    MsgBox "Double Click To Make Edit Mode Active"
    'End If
    End If
End Sub
Private Sub fgTrial_GotFocus()
If fgtrial.Rows = 2 And fgtrial.TextMatrix(1, 0) = "" Then
fgtrial.Row = 1
fgtrial.Col = 0
fgtrial.text = 1
fgtrial.Col = 1

End If
End Sub
Private Sub fgTrial_LeaveCell()


    If txtNewData.Visible = False Then
        Exit Sub
    Else
        'If IsNumeric(txtNewData.text) Then
        '    fgtrial.text = Str(txtNewData.text)
        'Else
            fgtrial.text = txtNewData.text
        'End If
        txtNewData.Visible = False
        txtNewData.text = ""
    End If
    If fgtrial.Col = coollsi Then
    Else
    If fgtrial.Col < fgtrial.Cols - 1 Then
      fgtrial.Col = fgtrial.Col + 1
      Else
      cmdAdd_Click
      fgtrial.Col = coollsi
      End If
     End If
End Sub


Private Sub Form_Load()
Call GetNewConnection2

'Set Rs1 = New Recordset
SQL = "select  pozicija,sifra,naziv,cena,kol,znes,x,y from trenutna"
Rs1.Open SQL, myConection, adOpenDynamic, adLockOptimistic

'Set Rs1 = DCON.Execute(SQL)
'ssqq = SQL

    If Rs1.EOF Then
    Rs1.AddNew
    'Rs1.Fields("datum") = Date
    Rs1.Update
    End If
        Set fgtrial.DataSource = Rs1

DoColumnSort
Set Rs1 = Nothing
Set DCON = Nothing
End Sub

Private Sub Form_Unload(cancel As Integer)
    If RS.State = 1 Then
        RS.Close
    End If
'    cn.Close
End Sub
Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------

    With fgtrial
    
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
        .CellBackColor = &HFFFFFF
        ' grey every other row
        Dim iLoop As Integer
       'If CatalogueName = "Category" Then
       'Else
       ' For iLoop = .FixedRows To .Rows - 1
       ' Dim asx As String
       ' asx = fgTrial.TextMatrix(iLoop, 1)
       '  .Row = iLoop
       '     .Col = .FixedCols
       '     .ColSel = .Cols() - .FixedCols - 1
           ' MsgBox asx
            'MsgBox (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'"))
        'If (Getnazi("select poknj from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & Trim(asx) & "'")) = "K" Then
       
           
      '      .CellBackColor = &HC0C0FF
      '      Else
       '     .CellBackColor = &HC0FFC0
            
       'End If
       ' Next iLoop
        '.FillStyle = flexFillSingle

        'End If
        .Redraw = True
        
    End With
   '  Dim iLoop As Integer
       
        For i = fgtrial.FixedCols To fgtrial.Cols - 1
        Dim asx As String
        
        asx = Trim(fgtrial.TextMatrix(0, i))
        'MsgBox asx
        If asx = "sifra" Then
        coollsi = i
        'Exit For
        End If
        If asx = "naziv" Then
        coollna = i
        
        'Exit For
        End If
        If asx = "cena" Then
        coollce = i
        'Exit For
        End If
        If asx = "kol" Then
        coollko = i
        'Exit For
        End If
         
         'fgtrial.Col = ""
      '   fgtrial.Col = iLoop
       Next i
fgtrial.ColWidth(coollna) = 3000

End Sub

Private Sub txtNewData_GotFocus()
      fgtrial.Row = curRow
    fgtrial.Col = curCol
   adRwFlag = False
   txtNewData.SelStart = Len(txtNewData)
   txtNewData.SelLength = Len(txtNewData) + 1
    'txtNewData.text = fgtrial.text
End Sub

Private Sub txtNewData_LostFocus()

 If curCol < edCol Then
        fgtrial.Row = curRow
        fgtrial.Col = curCol + 1
        fgTrial_KeyPress (0)
    Else
       fgTrial_LeaveCell
    End If
   End Sub
Private Sub txtNewData_KeyPress(KeyAscii As Integer)
'MsgBox ("3")
If KeyAscii = 13 Then
 fgtrial.text = txtNewData.text
    If fgtrial.Col = coollsi Then
       If RS.State = 1 Then RS.Close
       Dim ax As String
       ax = ""
       ax = (Getnazi("select madasifr from mada where madasifr='" & txtNewData.text & "'"))
       If ax = "" Then
       Dim novas, vi, dol As String
       vi = ""
       dol = ""
       novas = "/" & Trim(txtNewData.text) & "/"
       ax = (Getnazi("select madasifr from mada where dobavit_id like '%" & novas & "%'"))
       End If
       If ax = "" Then
       idar = ""
       iskalni = fgtrial.text
      ' DoSQL = ""
       ax = DoSQL("mada", "madasifr", "madanazi", "madanaz1")
       'MsgBox ax
       End If
       txtNewData.text = Trim((ax))
       fgtrial.text = Trim(ax)
       sifrt = (ax)
        RS.Open "select MADANAZI,MADAMPCD,madapd,postava from MADA where MADASIFR='" & ax & "'", myConection, adOpenStatic, adLockOptimistic
          If Not RS.EOF Then
              fgtrial.TextMatrix(fgtrial.Row, coollna) = Trim(RS!MADANAZI) & " "
             fgtrial.TextMatrix(fgtrial.Row, coollce) = Round(RS!MADAMPCD / (1 + (RS!madapd / 100)), 2)
          End If
          Call txtNewData_LostFocus
          
          fgtrial.Col = coollko
        'End If
    
      ElseIf fgtrial.Col = coollko Then
      'fgtrial.Rows = fgtrial.Rows + 1
      
      cmdAdd_Click
    fgtrial.Col = coollsi
      
      
      Else
      If fgtrial.Col = coollce Then
      txtNewData.text = strval(txtNewData.text)
 '
      End If
      End If


 
fgtrial.SetFocus
End If
End Sub



