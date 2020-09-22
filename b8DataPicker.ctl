VERSION 5.00
Begin VB.UserControl b8DataPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   Begin VB.CommandButton cmdClear 
      DisabledPicture =   "b8DataPicker.ctx":0000
      Height          =   345
      Left            =   3150
      Picture         =   "b8DataPicker.ctx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Width           =   345
   End
   Begin VB.CommandButton cmdPicker 
      DisabledPicture =   "b8DataPicker.ctx":0B14
      Height          =   345
      Left            =   2790
      Picture         =   "b8DataPicker.ctx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   345
   End
   Begin VB.TextBox txtDisplay 
      Height          =   345
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   2715
   End
End
Attribute VB_Name = "b8DataPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'constant
Private Const mDefAutoConnect As Boolean = True
'members

Private mSQLFields As String
Private mSQLTable As String
Private mSQLWhere As String
Private mSQLWhereFields As String
Private mSQLGroupBy As String
Private mSQLOrderBy As String
Private mSQLFilterString As String
Private mSQLWhereSeparator As String
Private mBoundFieldIndex As Integer
Private mDisplayFieldIndex As Integer
Private mAutoConnect As Boolean

Private mRecCount As Long
Private mBoundData As String

Private mForeColor As OLE_COLOR

Private Type udtColumn
    EditCtrl As Object
    dCustomWidth As Single
    nAlignment As Integer
    nSortOrder As lgSortTypeEnum
    nType As Integer
    lWidth As Long
    lX As Long
    MoveControl As Integer
    bVisible As Boolean
    sCaption As String
    sFormat As String
    sTag As String
End Type


'public vars
Public DropDBCon As New ADODB.Connection
Public DropRS As New ADODB.Recordset
Public DropGrid As LynxGrid3
Private mCols() As udtColumn

'events
Public Event BeforeDropDown()
Public Event Change()
'Default Property Values:
Const m_def_DropCaption = "Select Entry"
Const m_def_DropWinWidth = 6735
Const m_def_DropWinHeight = 3510
'Property Variables:
Dim m_DropCaption As String
Dim m_DropWinWidth As Integer
Dim m_DropWinHeight As Integer


Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function


Private Sub cmdClear_Click()
    Call ClearCurData
End Sub

Private Sub cmdPicker_Click()
    
    Dim sDT As String
    Dim sBT As String
    Dim OldBT As String
    
    RaiseEvent BeforeDropDown
    
    'clear custom search
    mSQLFilterString = ""
    
    'Call LoadData
    'Call LoadColumnHeaders
    
    If frmDataPicker.ShowPicker(UserControl.Parent, Me, sBT, sDT) = True Then
        
        OldBT = mBoundData
        
        txtDisplay.Text = sDT
        mBoundData = sBT
        
        If mBoundData <> OldBT Then
            RaiseEvent Change
        End If
    End If

End Sub

Public Sub LoadColumnHeaders()
    
    Dim li As Long
    
    For li = 0 To UBound(mCols)
        frmDataPicker.listEntries.AddColumn mCols(li).sCaption, CSng(mCols(li).lWidth), CLng(mCols(li).nAlignment), CLng(mCols(li).nType), CStr(mCols(li).sFormat)
    Next
    
End Sub

Private Sub UserControl_Initialize()
        ReDim mCols(0)
        Load frmDataPicker
        Set DropGrid = frmDataPicker.listEntries

End Sub

Private Sub UserControl_Resize()
    
    If GetWidth < 58 Then
        UserControl.Width = 58 * 15
    End If
    If GetHeight < 21 Then
        UserControl.Height = 21 * 15
    End If
    
    txtDisplay.Move 0, 1, GetWidth - cmdClear.Width - cmdPicker.Width - 4, GetHeight - 1
    cmdPicker.Move GetWidth - cmdPicker.Width - cmdClear.Width - 2, 0, cmdPicker.Width, GetHeight - 1
    cmdClear.Move GetWidth - cmdClear.Width, 0, cmdClear.Width, GetHeight - 1

End Sub


Public Property Get SQLFields() As String
    SQLFields = mSQLFields
End Property
Public Property Let SQLFields(ByVal NewValue As String)
    mSQLFields = NewValue
End Property

Public Property Get SQLTable() As String
    SQLTable = mSQLTable
End Property
Public Property Let SQLTable(ByVal NewValue As String)
    mSQLTable = NewValue
End Property

Public Property Get SQLWhereFields() As String
    SQLWhereFields = mSQLWhereFields
End Property
Public Property Let SQLWhereFields(ByVal NewValue As String)
    mSQLWhereFields = NewValue
End Property


Public Property Get SQLGroupBy() As String
    SQLGroupBy = mSQLGroupBy
End Property
Public Property Let SQLGroupBy(ByVal NewValue As String)
    mSQLGroupBy = NewValue
    PropertyChanged "SQLGroupBy"
End Property

Public Property Get SQLOrderBy() As String
    SQLOrderBy = mSQLOrderBy
End Property
Public Property Let SQLOrderBy(ByVal NewValue As String)
    mSQLOrderBy = NewValue
    PropertyChanged "SQLOrderBy"
End Property

Public Property Get SQLWhereSeparator() As String
    SQLWhereSeparator = mSQLWhereSeparator
End Property
Public Property Let SQLWhereSeparator(ByVal NewValue As String)
    mSQLWhereSeparator = NewValue
    PropertyChanged "SQLWhereSeparator"
End Property

Public Property Get SQLFilterString() As String
    SQLFilterString = mSQLFilterString
End Property
Public Property Let SQLFilterString(ByVal NewValue As String)
    mSQLFilterString = NewValue
    PropertyChanged "SQLFilterString"
End Property

Public Property Get SQLWhere() As String
    SQLWhere = mSQLWhere
End Property
Public Property Let SQLWhere(ByVal NewValue As String)
    mSQLWhere = NewValue
    PropertyChanged "SQLWhere"
End Property

Public Property Get BoundFieldIndex() As Integer
    BoundFieldIndex = mBoundFieldIndex
End Property
Public Property Let BoundFieldIndex(ByVal NewValue As Integer)
    mBoundFieldIndex = NewValue
    PropertyChanged "BoundFieldIndex"
End Property

Public Property Get DisplayFieldIndex() As Integer
    DisplayFieldIndex = mDisplayFieldIndex
End Property
Public Property Let DisplayFieldIndex(ByVal NewValue As Integer)
    mDisplayFieldIndex = NewValue
    PropertyChanged "DisplayFieldIndex"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Font() As StdFont
    Set Font = txtDisplay.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
    Set txtDisplay.Font = NewValue
    PropertyChanged "Font"
End Property

Public Property Get BoundData() As String
    BoundData = mBoundData
End Property
Public Property Let BoundData(ByVal NewValue As String)
    mBoundData = NewValue
End Property

Public Property Get DisplayData() As String
    DisplayData = txtDisplay.Text
End Property
Public Property Let DisplayData(ByVal NewValue As String)
    txtDisplay.Text = NewValue
End Property




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    mSQLFields = PropBag.ReadProperty("SQLFields", "")
    mSQLTable = PropBag.ReadProperty("SQLTable", "")
    mSQLWhere = PropBag.ReadProperty("SQLWhere", "")
    mSQLWhereFields = PropBag.ReadProperty("SQLWhereFields", "")
    mSQLGroupBy = PropBag.ReadProperty("SQLGroupBy", "")
    mSQLOrderBy = PropBag.ReadProperty("SQLOrderBy", "")

    mSQLWhereSeparator = PropBag.ReadProperty("SQLWhereSeparator", ",")
    
    
    mBoundFieldIndex = PropBag.ReadProperty("BoundFieldIndex", 0)
    mDisplayFieldIndex = PropBag.ReadProperty("DisplayFieldIndex", 0)

    Set txtDisplay.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cmdClear.Enabled = PropBag.ReadProperty("ClearEnabled", True)
    cmdPicker.Enabled = PropBag.ReadProperty("DropEnabled", True)
    Set Picture = PropBag.ReadProperty("ClearIcon", Nothing)
    Set Picture = PropBag.ReadProperty("DropIcon", Nothing)
    txtDisplay.Locked = PropBag.ReadProperty("TextLocked", True)
    m_DropWinWidth = PropBag.ReadProperty("DropWinWidth", m_def_DropWinWidth)
    m_DropWinHeight = PropBag.ReadProperty("DropWinHeight", m_def_DropWinHeight)
    txtDisplay.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtDisplay.Locked = PropBag.ReadProperty("Locked", True)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_DropCaption = PropBag.ReadProperty("DropCaption", m_def_DropCaption)
End Sub



Private Sub UserControl_Terminate()
    On Error Resume Next
    Unload frmDataPicker
    Err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SQLFields", mSQLFields, "")
    Call PropBag.WriteProperty("SQLTable", mSQLTable, "")
    Call PropBag.WriteProperty("SQLWhere", mSQLWhere, "")
    Call PropBag.WriteProperty("SQLWhereFields", mSQLWhereFields, "")
    Call PropBag.WriteProperty("SQLGroupBy", mSQLGroupBy, "")
    Call PropBag.WriteProperty("SQLOrderBy", mSQLOrderBy, "")
    Call PropBag.WriteProperty("SQLWhereSeparator", mSQLWhereSeparator, ",")
    
    Call PropBag.WriteProperty("BoundFieldIndex", mBoundFieldIndex, 0)
    Call PropBag.WriteProperty("DisplayFieldIndex", mDisplayFieldIndex, 0)
    Call PropBag.WriteProperty("Font", txtDisplay.Font, Ambient.Font)
    Call PropBag.WriteProperty("ClearEnabled", cmdClear.Enabled, True)
    Call PropBag.WriteProperty("DropEnabled", cmdPicker.Enabled, True)
    Call PropBag.WriteProperty("ClearIcon", Picture, Nothing)
    Call PropBag.WriteProperty("DropIcon", Picture, Nothing)
    Call PropBag.WriteProperty("TextLocked", txtDisplay.Locked, True)
    Call PropBag.WriteProperty("DropWinWidth", m_DropWinWidth, m_def_DropWinWidth)
    Call PropBag.WriteProperty("DropWinHeight", m_DropWinHeight, m_def_DropWinHeight)
    Call PropBag.WriteProperty("BackColor", txtDisplay.BackColor, &H80000005)
    Call PropBag.WriteProperty("Locked", txtDisplay.Locked, True)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("DropCaption", m_DropCaption, m_def_DropCaption)
End Sub


Public Function AddColumn(Optional Caption As String, Optional Width As Single, Optional Alignment As lgAlignmentEnum = lgAlignLeftCenter, Optional DataType As lgDataTypeEnum = lgString, Optional Format As String) As Long

    Dim lNewCol As Long
    
    If mCols(0).nAlignment <> 0 Then
        lNewCol = UBound(mCols) + 1
        ReDim Preserve mCols(lNewCol)
    End If
 
    With mCols(lNewCol)
        .sCaption = Caption
        .dCustomWidth = Width
        .lWidth = ScaleX(.dCustomWidth, vbPixels, vbPixels)
        .nAlignment = Alignment
        .nSortOrder = lgSTAscending
        .nType = DataType
        .sFormat = Format
        
        .bVisible = True
    End With
    AddColumn = lNewCol
    
End Function


Public Function LoadData() As Boolean

    'default
    LoadData = False
    mRecCount = 0
    
    DropGrid.Redraw = False
    DropGrid.Clear

    'generate and validate
    If Len(Trim(GenSQLCon)) < 1 Then
        GoTo RAE
    End If
        
    'connect
    If ConRS(DropDBCon, DropRS, GenSQLCon) = False Then
        GoTo RAE
    End If
    
    If AnyRecExist(DropRS) = False Then
        LoadData = True
    End If
    
    'fill
    mRecCount = GetRecCount(DropRS)
    
    If mRecCount < 1 Then
        GoTo RAE
    End If
    
    'return success
    LoadData = True
    
RAE:
    DropGrid.Redraw = True
    DropGrid.Refresh
End Function

Private Function GenSQLCon() As String
    
    Dim sNewWhere As String
    
    sNewWhere = GenSQLWhere
    
    If Len(Trim(sNewWhere)) > 1 Then
        sNewWhere = " WHERE " & sNewWhere
    Else
        sNewWhere = ""
    End If

    GenSQLCon = "SELECT " & mSQLFields & " " & _
                " FROM " & mSQLTable & " " & _
                sNewWhere & " " & _
                mSQLGroupBy & " " & _
                GenSQLOrderBy
                
End Function

Private Function GenSQLOrderBy() As String
    
    'default
    GenSQLOrderBy = ""
    
    If Len(Trim(mSQLOrderBy)) > 0 Then
        GenSQLOrderBy = "ORDER BY " & mSQLOrderBy
    End If
    
End Function

Private Function GenSQLWhere() As String
    
    Dim sNewWhere As String
    Dim i As Integer
    
    If Len(Trim(mSQLFilterString)) > 1 Then
        
        sNewWhere = Replace(mSQLWhereFields, mSQLWhereSeparator, " " & Chr(Asc("&")) & " ") & _
                " like '%" & Trim(mSQLFilterString) & "%'"
        
    End If
    
        
    If Len(Trim(mSQLWhere)) > 0 Then
        If Len(Trim(sNewWhere)) > 0 Then
            'add 'AND'
            sNewWhere = sNewWhere & " AND "
        End If
        
        sNewWhere = sNewWhere & "(" & mSQLWhere & ")"
    End If
    
    GenSQLWhere = Trim(sNewWhere)

End Function


Public Function GetCellTextToDisplay(ByVal lRow As Long, ByVal lCol As Long, ByRef sNewValue As String)

    Dim lDif As Long
    
    lDif = (DropRS.AbsolutePosition - 1) - lRow
    
    If lDif > 0 Then
        DropRS.MoveFirst
        DropRS.Move lRow
    ElseIf lDif < 0 Then
        DropRS.Move 0 - lDif
    End If
    
    sNewValue = ReadField(DropRS.Fields(lCol))
End Function

Public Function GetCurRecCount() As Long
    GetCurRecCount = mRecCount
End Function

Public Sub ClearCurData()
    txtDisplay.Text = ""
    mBoundData = ""
    RaiseEvent Change
End Sub

Public Sub FocusedDropButton()
    cmdPicker.SetFocus
End Sub

Public Sub FocusedClearButton()
    cmdClear.SetFocus
End Sub
Private Function ConRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String) As Boolean
    
    'default
    ConRS = False
    
    On Error GoTo Errh
    
    Set vRS = Nothing
    Set vRS = New ADODB.Recordset

    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConRS = True

Errh:

End Function


Public Function AnyRecExist(ByRef vRS As ADODB.Recordset) As Boolean
    
    If vRS.State = adStateClosed Then
        AnyRecExist = False
        Exit Function
    End If
        
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecExist = False
    Else
        On Error GoTo Errh
        vRS.MoveFirst
        AnyRecExist = True
    End If

    Exit Function
    '--------------------------
    
Errh:
    AnyRecExist = False
End Function


Private Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo Errh

    If Not IsNull(vField.Value) Then
        ReadField = vField.Value
    Else
        Select Case vField.Type
            Case adBigInt
                ReadField = 0
            Case adBinary
                ReadField = 0
            Case adBoolean
                ReadField = False
            'Case adByRef 'temp
            '    ReadField = 0
            Case adBSTR
                ReadField = ""
            Case adChar
                ReadField = ""
            Case adCurrency
                ReadField = 0
            Case adDate
                ReadField = CDate(0)
            Case adDBDate
                ReadField = CDate(0)
            Case adDBTime
                ReadField = FormatDateTime(CDate(0), vbLongTime)
            Case adDBTimeStamp
                ReadField = CDate(0)
            Case adDecimal
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case adEmpty 'temp
                ReadField = ""
            Case adError
                ReadField = 0
            Case adNumeric
                ReadField = 0
            Case adDouble
                ReadField = 0
            Case Else
                ReadField = ""
            End Select
    End If
    
    Exit Function
    
Errh:
    ReadField = ""
End Function


Private Function GetRecCount(ByRef vRS As ADODB.Recordset) As Long
    If AnyRecExist(vRS) Then
        vRS.Requery
        vRS.MoveLast
        GetRecCount = vRS.RecordCount
    Else
        GetRecCount = 0
    End If
End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdClear,cmdClear,-1,Enabled
Public Property Get ClearEnabled() As Boolean
Attribute ClearEnabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    ClearEnabled = cmdClear.Enabled
End Property

Public Property Let ClearEnabled(ByVal New_ClearEnabled As Boolean)
    cmdClear.Enabled() = New_ClearEnabled
    PropertyChanged "ClearEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdPicker,cmdPicker,-1,Enabled
Public Property Get DropEnabled() As Boolean
Attribute DropEnabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    DropEnabled = cmdPicker.Enabled
End Property

Public Property Let DropEnabled(ByVal New_DropEnabled As Boolean)
    cmdPicker.Enabled() = New_DropEnabled
    PropertyChanged "DropEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdClear,cmdClear,-1,Picture
Public Property Get ClearIcon() As Picture
Attribute ClearIcon.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox control, if Style is set to 1."
    Set ClearIcon = cmdClear.Picture
End Property

Public Property Set ClearIcon(ByVal New_ClearIcon As Picture)
    Set cmdClear.Picture = New_ClearIcon
    PropertyChanged "ClearIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdPicker,cmdPicker,-1,Picture
Public Property Get DropIcon() As Picture
Attribute DropIcon.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox control, if Style is set to 1."
    Set DropIcon = cmdPicker.Picture
End Property

Public Property Set DropIcon(ByVal New_DropIcon As Picture)
    Set cmdPicker.Picture = New_DropIcon
    PropertyChanged "DropIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDisplay,txtDisplay,-1,Locked
Public Property Get TextLocked() As Boolean
Attribute TextLocked.VB_Description = "Determines whether a control can be edited."
    TextLocked = txtDisplay.Locked
End Property

Public Property Let TextLocked(ByVal New_TextLocked As Boolean)
    txtDisplay.Locked() = New_TextLocked
    PropertyChanged "TextLocked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,5000
Public Property Get DropWinWidth() As Integer
    DropWinWidth = m_DropWinWidth
End Property

Public Property Let DropWinWidth(ByVal New_DropWinWidth As Integer)
    m_DropWinWidth = New_DropWinWidth
    PropertyChanged "DropWinWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,500
Public Property Get DropWinHeight() As Integer
    DropWinHeight = m_DropWinHeight
End Property

Public Property Let DropWinHeight(ByVal New_DropWinHeight As Integer)
    m_DropWinHeight = New_DropWinHeight
    PropertyChanged "DropWinHeight"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DropWinWidth = m_def_DropWinWidth
    m_DropWinHeight = m_def_DropWinHeight
    m_DropCaption = m_def_DropCaption
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDisplay,txtDisplay,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtDisplay.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtDisplay.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDisplay,txtDisplay,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtDisplay.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtDisplay.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtDisplay.Enabled = New_Enabled
    cmdPicker.Enabled = New_Enabled
    cmdClear.Enabled = New_Enabled
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Select Entry
Public Property Get DropCaption() As String
    DropCaption = m_DropCaption
End Property

Public Property Let DropCaption(ByVal New_DropCaption As String)
    m_DropCaption = New_DropCaption
    PropertyChanged "DropCaption"
End Property

