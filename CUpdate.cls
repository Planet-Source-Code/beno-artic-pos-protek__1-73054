VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements CDbase

Private DBCon As CDbase

Private CUpd As New CUpdate
Dim Farray() As String




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
DBCon.OpenDb
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

Public Function FieldVal(ParamArray Fields() As Variant)
Dim Fcount As Integer

 ReDim Farray(UBound(Fields))

    For Fcount = LBound(Fields) To UBound(Fields)
        
            Farray(Fcount) = Fields(Fcount)
     
    Next
    
End Function

Public Sub Update(ParamArray FieldName() As Variant)
On Error GoTo UpErr
Dim Fcount As Integer
      
    If DBCon.ClauseStatement = "" Then
        For Fcount = LBound(FieldName) To UBound(FieldName)
            DBCon.OpenDb.Execute "Update " & DBCon.TableName & " set " & _
            FieldName(Fcount) & "='" & Farray(Fcount) & "'"
        Next
    Else
        For Fcount = LBound(FieldName) To UBound(FieldName)
            DBCon.OpenDb.Execute "Update " & DBCon.TableName & " set " & _
            FieldName(Fcount) & "='" & Farray(Fcount) & "' " & DBCon.ClauseStatement
        Next
    End If
    

Exit Sub

UpErr:
        If DBCon.ClauseStatement = "" Then
        For Fcount = LBound(FieldName) To UBound(FieldName)
            DBCon.OpenDb.Execute "Update " & DBCon.TableName & " set " & _
            FieldName(Fcount) & "=" & Farray(Fcount)
        Next
    Else
        For Fcount = LBound(FieldName) To UBound(FieldName)
            DBCon.OpenDb.Execute "Update " & DBCon.TableName & " set " & _
            FieldName(Fcount) & "='" & Farray(Fcount) & "' " & DBCon.ClauseStatement
        Next
    End If
    
    Resume Next

End Sub

