VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CDbase

Private vSQL As String
Private DBCon As CDbase
Private CRS As ADODB.Recordset

Public Property Get SQL() As String
SQL = vSQL
End Property

Public Property Let SQL(nSQL As String)
vSQL = nSQL
End Property

'Public Sub Bind(Cntrl As Object, Optional Fields As String)
' If TypeOf Cntrl Is DataGrid Then
'    Call BindGrid(Cntrl, Fields)
' ElseIf TypeOf Cntrl Is TextBox Then
'    Call BindText(Fields)
' ElseIf TypeOf Cntrl Is ComboBox Then
'    Call BindCombo(Fields)
' ElseIf TypeOf Cntrl Is ListView Then
'    Call BindListview(Fields)
' ElseIf TypeOf Cntrl Is ListBox Then
'    Call BindListBox(Fields)
' Else
'    Call BindElse(Fields)
' End If
'
'End Sub
Public Sub BindGrid(Gname As DataGrid, Optional Fields As String)
  '  Call RS
     '  Set Gname = RS
       ' Set CDbase_RS = CDbase_RS
'

Dim RSS As ADODB.Recordset


'Set RSS = New ADODB.Recordset
'Set RSS = CDbase_OpenDb.Execute(SQL)


Set Gname.DataSource = RS


'Set RSS = Nothing

End Sub
Public Sub BindText(TxtLabel As Object, Optional Fields As String)
     TxtLabel = RS.Fields(Fields)
End Sub
Private Function BindCombo(Optional Fields As String)
    Call RS
End Function
Private Function BindListview(Optional Fields As String)
    Call RS
End Function
Private Function BindListBox(Optional Fields As String)
    Call RS
End Function
Private Function BindElse(Optional Fields As String)
    Call RS
End Function

Private Property Let CDbase_ClauseStatement(RHS As String)
DBCon.ClauseStatement = RHS
End Property

Private Property Get CDbase_ClauseStatement() As String
CDbase_ClauseStatement = DBCon.ClauseStatement
End Property

Private Sub CDbase_CloseDB()
DBCon.CloseDB
End Sub

Private Property Let CDbase_DBPath(RHS As String)
DBCon.DBPath = RHS
End Property

Private Property Get CDbase_DBPath() As String
CDbase_DBPath = DBCon.DBPath
End Property

Private Property Let CDbase_FieldName(RHS As String)
DBCon.FieldName = RHS
End Property

Private Property Get CDbase_FieldName() As String
CDbase_FieldName = DBCon.FieldName
End Property

Private Function CDbase_OpenDb(Optional StrProvider As String) As ADODB.Connection
On Error GoTo OPENERR

DBCon.OpenDb

Exit Function

OPENERR:
    MsgBox Err.Description, vbInformation

End Function

Private Property Let CDbase_TableName(RHS As String)
DBCon.TableName = RHS
End Property

Private Property Get CDbase_TableName() As String
CDbase_TableName = DBCon.TableName
End Property

Private Sub Class_Initialize()
Set DBCon = New CDbase
End Sub

Private Sub Class_Terminate()
Set DBCon = Nothing
End Sub

Public Function RS() As ADODB.Recordset

Set RS = New ADODB.Recordset

'Set RS = CDbase_OpenDb.Execute(SQL)
Set RS = DBCon.OpenDb.Execute(SQL)


End Function
Private Property Set CDbase_RS(ByVal RHS As ADODB.Recordset)
Set DBCon.RS = RHS
End Property

Private Property Get CDbase_RS() As ADODB.Recordset
Set CDbase_RS = New ADODB.Recordset
End Property
Public Sub CloseRS()
Set CRS = Nothing
End Sub
Public Sub NextRecord()
On Error Resume Next
RS.MoveNext
End Sub

Public Sub PreviousRecord()
On Error Resume Next
RS.MovePrevious
End Sub
Public Sub LastRecord()
On Error Resume Next
RS.MoveLast
End Sub

Public Sub FirstRecord()
On Error Resume Next
RS.MoveFirst
End Sub

