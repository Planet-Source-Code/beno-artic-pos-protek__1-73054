VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public rs As New ADODB.Recordset
Private Table_Name As String
Private Field_Name As String
Private Clause As String
Private DPath

Public Property Get DBPath() As String
DBPath = DPath
End Property

Public Property Let DBPath(NewPath As String)
 DPath = NewPath
End Property

Public Function OpenDb(Optional StrProvider As String) As ADODB.Connection
On Error GoTo OPENERR
Dim CNSTR As String
Set OpenDb = New ADODB.Connection
OpenDb.CursorLocation = adUseClient
If StrProvider = "" Then
'StrProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";Persist Security Info=False;Jet OLEDB:Database Password=prosvent"
CNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
OpenDb.Open CNSTR
Else
On Error GoTo PROVERR
OpenDb.Open StrProvider & DBPath
PROVERR:
       OpenDb.Open StrProvider
End If
Exit Function
OPENERR:
    MsgBox Err.Description, vbInformation
End Function

Public Sub CloseDB()
''Set OpenDb = Nothing
End Sub
Public Property Get TableName() As String
    TableName = Table_Name
End Property

Public Property Let TableName(NewTable As String)
    Table_Name = NewTable
End Property

Public Property Get FieldName() As String
    FieldName = Field_Name
End Property

Public Property Let FieldName(NewField As String)
    Field_Name = NewField
End Property

Public Property Get ClauseStatement() As String
    ClauseStatement = Clause
End Property

Public Property Let ClauseStatement(NewClause As String)
    Clause = NewClause
End Property

