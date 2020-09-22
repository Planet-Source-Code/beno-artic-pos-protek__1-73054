Attribute VB_Name = "modRSPack"
Option Explicit


Public Type tPack

    PackID As Long
    PackTitle As String
    
End Type


Public Function AddPack(ByVal sPackTitle As String) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim vPack As tPack
    
    
    'default
    AddPack = False
    
    sSQL = "SELECT * FROM tblPack WHERE PackTitle='" & sPackTitle & "'"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
      '  WriteErrorLog "modRSPack", "AddPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        AddPack = True
        GoTo RAE
    End If
    
    'set new pack
    vPack.PackTitle = sPackTitle
    'generate New Package ID
    vPack.PackID = GetNewPackID
        
    'add new record
    vRS.AddNew
    
    If WritePack(vRS, vPack) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    AddPack = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditPack(vPack As tPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    EditPack = False
    
    sSQL = "SELECT * FROM tblPack WHERE PackID=" & vPack.PackID
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSPack", "EditPack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        'WriteErrorLog "modRSPack", "EditPack", "PackID does not exist. PackID= " & vPack.PackID
        GoTo RAE
    End If
    
    'edit
    If WritePack(vRS, vPack) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditPack = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeletePack(ByVal iPackID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    On Error GoTo RAE
    'default
    DeletePack = False
    
    sSQL = "DELETE * FROM tblPack WHERE PackID=" & iPackID
    
    Dim sErrD As String
    Dim iErrN As Long
    If ConnectRS(myConection, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
          '  WriteErrorLog "modRSPack", "DeletePack", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If
     
    DeletePack = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetPackByTitle(sPackTitle As String, vPack As tPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPackByTitle = False
    
    sSQL = "SELECT * FROM tblPack WHERE PackTitle='" & sPackTitle & "'"

    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSPack", "GetPackByTitle", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPack(vRS, vPack) = False Then
        GoTo RAE
    End If
    
    GetPackByTitle = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function GetPackByID(ByVal iPackID As Long, ByRef vPack As tPack) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPackByID = False
    
    sSQL = "SELECT * FROM tblPack WHERE PackID=" & iPackID

    If ConnectRS(myConection, vRS, sSQL) = False Then
      '  WriteErrorLog "modRSPack", "GetPackByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadPack(vRS, vPack) = False Then
        GoTo RAE
    End If
    
    GetPackByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetPackTitleByID(ByVal iPackID As Long) As String
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetPackTitleByID = ""
    
    sSQL = "SELECT tblPack.PackTitle FROM tblPack WHERE PackID=" & iPackID

    If ConnectRS(myConection, vRS, sSQL) = False Then
       ' WriteErrorLog "modRSPack", "GetPackTitleByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    GetPackTitleByID = ReadField(vRS.Fields("PackTitle"))
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyPackExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyPackExist = False
    
    sSQL = "SELECT * FROM tblPack"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSPack", "AnyPackExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyPackExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewPackID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewPackID = -1
    
    sSQL = "SELECT Max(tblPack.PackID)+1 AS MaxOfPackID" & _
            " From tblPack"

    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSPack", "GetNewPackID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewPackID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewPackID = ReadField(vRS.Fields("MaxOfPackID"))
    
    If GetNewPackID < 1 Then
        GetNewPackID = 1
    End If
    
RAE:
    Set vRS = Nothing
    err.clear
End Function


Public Sub FillPackToCMB(ByRef cmb As ComboBox)

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    

    sSQL = "SELECT tblPack.PackTitle" & _
            " From tblPack" & _
            " ORDER BY tblPack.PackTitle"


    cmb.clear
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSAddress", "FillPackToCMB", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If

    vRS.MoveFirst
    While vRS.EOF = False
        cmb.AddItem ReadField(vRS.Fields("PackTitle"))
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    
End Sub



Public Function ReadPack(ByRef vRS As ADODB.Recordset, ByRef vPack As tPack) As Boolean
    
    'default
    ReadPack = False
    
    On Error GoTo RAE
    
    With vPack
        
        .PackID = ReadField(vRS.Fields("PackID"))
        .PackTitle = ReadField(vRS.Fields("PackTitle"))

    End With
    
    ReadPack = True
    Exit Function
    
RAE:
    
End Function

Public Function WritePack(ByRef vRS As ADODB.Recordset, ByRef vPack As tPack) As Boolean
    
    'default
    WritePack = False
    
    'On Error GoTo RAE

    With vPack
    
        vRS.Fields("PackID") = .PackID
        vRS.Fields("PackTitle") = .PackTitle
    
    End With

    WritePack = True
    Exit Function
    
RAE:
    MsgBox err.Description
End Function

