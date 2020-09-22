VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CDbase
Private DBCon As CDbase
Private Field_Val As String
Private Farray() As String
Private Varray() As String
Private Farray2(50) As String
Private Fcnt As Integer
Private finalVal As String
Private TotalVal As Integer
Public Enum CDataType
    [CText]
    [CNum]
    [CBoolean]
End Enum

Private Sub CDbase_CloseDB()
DBCon.CloseDB
End Sub

Private Property Let CDbase_DBPath(RHS As String)
DBCon.DBPath = RHS
End Property

Private Property Get CDbase_DBPath() As String
CDbase_DBPath = DBCon.DBPath

End Property



Private Function CDbase_OpenDb(Optional StrProvider As String) As ADODB.Connection
On Error GoTo OPENERR

Dim CNSTR As String

Set CDbase_OpenDb = New ADODB.Connection

CDbase_OpenDb.CursorLocation = adUseClient

If StrProvider = "" Then

'StrProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";Persist Security Info=False;Jet OLEDB:Database Password=prosvent"

CNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CDbase_DBPath

CDbase_OpenDb.Open CNSTR

Else
On Error GoTo PROVERR

CDbase_OpenDb.Open StrProvider & CDbase_DBPath


PROVERR:
        CDbase_OpenDb.Open StrProvider

End If

Exit Function
OPENERR:
    MsgBox err.Description, vbInformation

End Function

Private Sub Class_Initialize()

Set DBCon = New CDbase

End Sub

Private Sub Class_Terminate()
Set DBCon = Nothing
End Sub
Private Property Set CDbase_RS(ByVal RHS As ADODB.Recordset)
Set DBCon.RS = New ADODB.Recordset
End Property

Private Property Get CDbase_RS() As ADODB.Recordset
Set CDbase_RS = DBCon.RS
End Property

Private Property Get CDbase_ClauseStatement() As String
CDbase_ClauseStatement = DBCon.ClauseStatement
End Property

Private Property Let CDbase_ClauseStatement(RHS As String)
DBCon.ClauseStatement = RHS
End Property

Private Property Get CDbase_FieldName() As String
CDbase_FieldName = DBCon.FieldName
End Property

Private Property Let CDbase_FieldName(RHS As String)
DBCon.FieldName = RHS
End Property

Private Property Let CDbase_TableName(RHS As String)
DBCon.TableName = RHS
End Property

Private Property Get CDbase_TableName() As String
CDbase_TableName = DBCon.TableName
End Property

Public Function FieldVal(FieldValue As String, FieldDataType As CDataType)

Fcnt = Fcnt + 1
'ReDim Farray2()

        
    If FieldDataType = CText Then
      Farray2(Fcnt) = "'" & FieldValue & "'"
    Else
       Farray2(Fcnt) = FieldValue
   
    End If
   
       

End Function

'Private Sub Insert2()
'On Error GoTo InsErr2
'If DBCon.ClauseStatement = "" Then
'    If DBCon.FieldName <> "" Then
'        DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & " (" & DBCon.FieldName & _
'         ") values(" & FieldVal & ")"
'    Else
'         DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & " values(" & FieldVal & ")"
'    End If
'
'Else
'     If DBCon.FieldName <> "" Then
'        DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & " (" & DBCon.FieldName & _
'         ") values(" & FieldVal & ") " & DBCon.ClauseStatement
'    Else
'         DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & " values(" & FieldVal & ") " & DBCon.ClauseStatement
'    End If
'
'End If
'
'Exit Sub
'
'InsErr2:
'    MsgBox Err.Number & ": " & Err.description
'
'    Resume Next
'
'End Sub
Public Sub Insert()
On Error GoTo InsErr

Dim Fcount As Integer


    If DBCon.FieldName <> "" Then
       For Fcount = 1 To Fcnt
            If Fcount < Fcnt Then
                finalVal = finalVal & Farray2(Fcount) & ","
            Else
                finalVal = finalVal & Farray2(Fcount)
            End If
        Next
        
        DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & "(" & DBCon.FieldName & _
         ") values(" & finalVal & ")"
    
    Else
        For Fcount = 1 To Fcnt
           If Fcount < Fcnt Then
                finalVal = finalVal & Farray2(Fcount) & ","
            Else
                finalVal = finalVal & Farray2(Fcount)
            End If
        Next
      
            DBCon.OpenDb.Execute "Insert into " & DBCon.TableName & " values(" & finalVal & ")"
       
    End If
    
 
    
'Fcnt = 0






Exit Sub

InsErr:
        MsgBox err.Description & vbCrLf & "Please check if specified the Table Name.", vbInformation
        Resume Next
End Sub



Public Function AUTONUM(ByVal Active_Connection As ADODB.Connection, ByVal TABLE As String, ByVal PKEY As String, ByVal PREFIX As String, Optional ByRef DISPLAY As Object) As String
Dim AUTORS As New ADODB.Recordset
Dim ID1 As String
TABLE = UCase(TABLE)
PKEY = UCase(PKEY)

Set AUTORS = New Recordset
Set AUTORS = Active_Connection.Execute("Select * from " & TABLE)
If AUTORS.RecordCount >= 1 Then
    AUTORS.MoveLast
    AUTONUM = AUTORS.Fields("sifra") + 1
Else
AUTONUM = "0"
AUTONUM = AUTONUM + 1

End If

 
     
     If Not DISPLAY Is Nothing Then
        DISPLAY = AUTONUM
     End If
     
     
    Set AUTORS = Nothing
    Set Active_Connection = Nothing
    

End Function








