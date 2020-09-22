VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDataPicker 
   BackColor       =   &H00EDEBE9&
   BorderStyle     =   0  'None
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView listEntries 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Left            =   3150
      Picture         =   "frmDataPicker.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3090
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
      Left            =   60
      TabIndex        =   2
      Text            =   "Enter search text here. - [ Ctrl + F ]"
      Top             =   3090
      Width           =   3075
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   3120
      Width           =   765
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
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
      Left            =   4770
      TabIndex        =   0
      Top             =   3120
      Width           =   885
   End
   Begin VB.PictureBox ilList 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   5100
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   1740
      Width           =   1200
   End
   Begin VB.Shape shpMB 
      BorderColor     =   &H00926747&
      Height          =   3465
      Left            =   0
      Top             =   120
      Width           =   6705
   End
   Begin VB.Line line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   16
      X2              =   408
      Y1              =   202
      Y2              =   202
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Entry"
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
      Left            =   -30
      TabIndex        =   4
      Top             =   60
      Width           =   6705
   End
   Begin VB.Shape shpCapBor 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   6705
   End
End
Attribute VB_Name = "frmDataPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim R As RECT


Dim Alignable As Boolean


Dim MyDPCtl As b8DataPicker

Dim ReadyToDisplay As Boolean

Dim mBoundData As String
Dim mDisplayData As String

Dim mShowPicker As Boolean

Public Function ShowPicker(ByRef frm As Form, ByRef ctl As b8DataPicker, ByRef sBoundData As String, ByRef sDisplayData As String) As Boolean
    
    'default
    ShowPicker = False
    ReadyToDisplay = False
    
    Set MyDPCtl = ctl
    
    'set flag
    ReadyToDisplay = True
    
    'resize form
    Me.Width = MyDPCtl.DropWinWidth
    Me.Height = MyDPCtl.DropWinHeight

    'align
    Dim R As RECT
    GetWindowRect MyDPCtl.hWnd, R

    Dim NewLeft As Long
    Dim NewTop As Long

    If (R.Left * Screen.TwipsPerPixelX + Me.Width) > Screen.Width Then
        NewLeft = (R.Right * Screen.TwipsPerPixelX) - Me.Width
    Else
        NewLeft = R.Left * Screen.TwipsPerPixelX
    End If
        
    If (R.Bottom * Screen.TwipsPerPixelY + Me.Height) > Screen.Height Then
        NewTop = (R.Top * Screen.TwipsPerPixelY) - Me.Height
        If NewTop < 0 Then NewTop = 0
    Else
        NewTop = R.Bottom * Screen.TwipsPerPixelY
    End If
        
    If NewLeft < 0 Then
        NewLeft = 0
    End If
    If NewTop < 0 Then
        NewTop = 0
    End If
    
    Me.Left = NewLeft
    Me.Top = NewTop
    
    'set caption
    lblCaption.Caption = MyDPCtl.DropCaption

    'show form
    'temp
    Me.Show vbModal ', Frm
    
    'return
    If mShowPicker = True Then
        ShowPicker = mShowPicker
        sBoundData = mBoundData
        sDisplayData = mDisplayData
    Else
        ShowPicker = False
    End If
End Function


Private Sub FillBlank(ByVal lRowCount As Long)
    
    Dim li As Long
    
    'listEntries.Redraw = False
   ' listEntries.ListItems.Clear
    
    For li = 0 To lRowCount - 1
    
        listEntries.ListItems.Add "Loading..."
        'format
       ' listEntries.CellForeColor(li, MyDPCtl.DisplayFieldIndex) = &HC00000
       ' listEntries.ItemImage(li) = 1
    
    Next
    
   ' listEntries.Redraw = True
    listEntries.Refresh

End Sub







Private Sub cmdCancel_Click()
    mShowPicker = False
    Unload Me
End Sub


Private Sub cmdFilter_Click()

    If Not ReadyToDisplay Then
        Exit Sub
    End If
    
    If txtFilter.text = "Enter search text here. - [ Ctrl + F ]" Then
    
        MyDPCtl.SQLFilterString = ""
    Else
    
        MyDPCtl.SQLFilterString = txtFilter.text
    End If
    
    MyDPCtl.LoadData
    'fill blank
    FillBlank MyDPCtl.GetCurRecCount
    
End Sub

Private Sub cmdSelect_Click()

    If Not ReadyToDisplay Then
        Exit Sub
    End If
    
    If listEntries.RowCount < 1 Then
        Exit Sub
    End If
    
    
    mBoundData = listEntries.CellText(listEntries.Row, CLng(MyDPCtl.BoundFieldIndex))
    mDisplayData = listEntries.CellText(listEntries.Row, CLng(MyDPCtl.DisplayFieldIndex))
    
    mShowPicker = True
    
    Unload Me
End Sub

Private Sub Form_Activate()
    
    'listEntries.RowHeightMin = 21
    'listEntries.ImageList = ilList
    
    DoEvents
    Me.AutoRedraw = False


    Call MyDPCtl.LoadData
    'Call MyDPCtl.LoadColumnHeaders
    
    'fill blank
    FillBlank MyDPCtl.GetCurRecCount
    
    ReadyToDisplay = True

    Me.AutoRedraw = True
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Me.ActiveControl.Name = listEntries.Name Then
        If Me.ActiveControl.Name = listEntries.Name Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 70 And Shift = 2 Then
        'Ctrl + 'F'
        txtFilter.SetFocus
    ElseIf KeyCode = 40 Or KeyCode = 38 Then
        If Me.ActiveControl.Name <> listEntries.Name Then
            listEntries.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpMB.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    lblCaption.Move 0, lblCaption.Top, Me.ScaleWidth
    shpCapBor.Move 0, 0, Me.ScaleWidth
    
    listEntries.Move 2, _
                    shpCapBor.Top + shpCapBor.Height + 1, _
                    Me.ScaleWidth - 4, _
                    Me.ScaleHeight - (shpCapBor.Top + shpCapBor.Height + 1 + cmdCancel.Height + 8)
    line2.X1 = 0
    line2.X2 = Me.ScaleWidth
    line2.Y1 = listEntries.Top + listEntries.Height + 1
    line2.Y2 = line2.Y1
    
    txtFilter.Move 4, Me.ScaleHeight - txtFilter.Height - 3, Me.ScaleWidth - 170
    cmdFilter.Move txtFilter.Left + txtFilter.Width + 2, txtFilter.Top
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 3, Me.ScaleHeight - cmdCancel.Height - 3
    cmdSelect.Move Me.ScaleWidth - cmdCancel.Width - cmdSelect.Width - 3, Me.ScaleHeight - cmdSelect.Height - 3
    
    err.clear
End Sub

Private Sub Form_Unload(cancel As Integer)
    ReadyToDisplay = False
    
    'release recordset
    On Error Resume Next
    'Call MyDPCtl.DropRS.Close
    err.clear
End Sub




Private Sub listEntries_BeforeDrawText(Row As Long, Col As Long, sNewValue As String)
    MyDPCtl.GetCellTextToDisplay Row, Col, sNewValue
End Sub

'Private Sub listEntries_ColumnClick(Col As Long)

    'If Not ReadyToDisplay Then
    '    Exit Sub
    'End If
    
   ' If MyDPCtl.SQLOrderBy = MyDPCtl.DropRS.Fields(Col).Properties.Item(0) Then
   '     MyDPCtl.SQLOrderBy = MyDPCtl.DropRS.Fields(Col).Properties.Item(0) & " DESC"
   ' Else
   '     MyDPCtl.SQLOrderBy = MyDPCtl.DropRS.Fields(Col).Properties.Item(0)
   ' End If
   '
   ' MyDPCtl.SQLFilterString = txtFilter.Text
    
   ' MyDPCtl.LoadData
   ' FillBlank MyDPCtl.GetCurRecCount
    
'End Sub

Private Sub listEntries_DblClick()
    If ReadyToDisplay Then
        Call cmdSelect_Click
    End If
End Sub







Private Sub txtFilter_Change()

    If Not ReadyToDisplay Then
        Exit Sub
    End If

    'delay 0.4 second
    'code by: VIncent J. Jamero
    '------------------------------------------------
    Static DelayStart As Single
    Static notFirst As Boolean
    DelayStart = GetTickCount + 400
    If notFirst = True Then Exit Sub
    notFirst = True
    While GetTickCount < DelayStart
        DoEvents
    Wend
    notFirst = False
    '------------------------------------------------
    'the next line will be if executed if user pause typing in 0.3 second

    Call cmdFilter_Click
End Sub

Private Sub txtFilter_GotFocus()
    If txtFilter.text = "Enter search text here. - [ Ctrl + F ]" Then
        txtFilter.text = ""
    End If
End Sub

Private Sub txtFilter_LostFocus()
    If Len(Trim(txtFilter.text)) < 1 Then
        txtFilter.text = "Enter search text here. - [ Ctrl + F ]"
    End If
End Sub
