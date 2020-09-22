Attribute VB_Name = "modRSProd"
Option Explicit


Public Type tProd

    ProdID As Long
    ProdCode As String
    ProdDescription As String

    FK_PackID As Double
    FK_CatID As Long
        
    BegInvStock As Double
    
    SupPrice As Double
    SRPrice As Double
    
    Active As Boolean
     RC As Double
    RM As Double
    RCU As String
    RMU As Long
    
End Type


Public Function AddProd(vProd As tProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AddProd = False
    
    sSQL = "SELECT * FROM mada WHERE madanazi='" & vProd.ProdDescription & "' OR madasifr='" & vProd.ProdID & "'"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSProd", "AddProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = True Then
        'WriteErrorLog "modRSProd", "AddProd", "Adding Failed. Reaseon: Duplication of ProdCode or ProdID"
        GoTo RAE
    End If
    
    
    'add new record
    vRS.AddNew
    
    If WriteProd(vRS, vProd) = False Then
        GoTo RAE
    End If
    
    vRS.Update
   
    
    AddProd = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function EditProd(vProd As tProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim tmpProduct As tProd
    
    'default
    EditProd = False
    
    sSQL = "SELECT * FROM mada WHERE madasifr='" & vProd.ProdID & "'"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
      '  WriteErrorLog "modRSProd", "EditProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If

    If GetProdByID(vProd.ProdID, tmpProduct) = False Then
       ' WriteErrorLog "modRSProd", "EditProd", "Failed on: 'GetProdByID(vProd.ProdID, tmpProduct) = False'"
        GoTo RAE
    End If

    'check for description duplication
    If LCase(Trim(vProd.ProdDescription)) <> LCase(Trim(tmpProduct.ProdDescription)) Then
        If modRSProd.GetProdByDescription(vProd.ProdDescription, tmpProduct) = True Then
            '' "modRSProd", "EditProd", "Duplicate Description | Failed on: 'LCase(Trim(vProd.ProdDescription)) <> LCase(Trim(tmpProduct.ProdDescription))'"
            GoTo RAE
        End If
    End If
    
    
    'edit
    If WriteProd(vRS, vProd) = False Then
        GoTo RAE
    End If
    
    vRS.Update

    EditProd = True
    
    'Update Inventory
   ' Call modRSStockInv.ClearStockInvByProd(vProd.ProdID, vProd.RM)
    
RAE:
    Set vRS = Nothing
End Function


Public Function DeleteProd(ByVal iProdID As Long) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    Dim sErrD As String
    Dim iErrN As Long
    
    On Error GoTo RAE
    'default
    DeleteProd = False
    
    sSQL = "DELETE * FROM mada WHERE madasifr='" & iProdID & "'"
    
    If ConnectRS(myConection, vRS, sSQL, False, iErrN, sErrD) = False Then
        If iErrN = -2147467259 Then
            'it includes releted data
            MsgBox "Unable to delete entry. It includes other related record." & vbNewLine & vbNewLine & _
                    "Details: " & sErrD, vbExclamation
        Else
         '   WriteErrorLog "modRSProd", "DeleteProd", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
            GoTo RAE
        End If
    End If

    DeleteProd = True
    
RAE:
    Set vRS = Nothing
End Function



Public Function GetProdByID(ByVal iProdID As Long, ByRef vProd As tProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProdByID = False
    
    sSQL = "SELECT * FROM mada WHERE madasifr='" & iProdID & "'"
'MsgBox (sSQL)
    If ConnectRS(myConection, vRS, sSQL) = False Then
        ' "modRSProd", "GetProdByID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProd(vRS, vProd) = False Then
        GoTo RAE
    End If
    
    GetProdByID = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetProdByDescription(ByVal sProdDescription As String, vProd As tProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProdByDescription = False
    
    sSQL = "SELECT * FROM mada WHERE madanazi='" & sProdDescription & "'"

    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSProd", "GetProdByDescription", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProd(vRS, vProd) = False Then
        GoTo RAE
    End If
    
    GetProdByDescription = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetProdByCode(ByVal sProdCode As String, vProd As tProd) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProdByCode = False
    
    sSQL = "SELECT * FROM mada WHERE madagrup='" & sProdCode & "'"

    If ConnectRS(myConection, vRS, sSQL) = False Then
      '  WriteErrorLog "modRSProd", "GetProdByCode", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    
    If ReadProd(vRS, vProd) = False Then
        GoTo RAE
    End If
    
    GetProdByCode = True
    
RAE:
    Set vRS = Nothing
End Function

Public Function AnyProdExist() As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    AnyProdExist = False
    
    sSQL = "SELECT * FROM mada"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
       ' WriteErrorLog "modRSProd", "AnyProdExist", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    
    AnyProdExist = True
    
RAE:
    Set vRS = Nothing
End Function


Public Function GetNewProdID() As Long
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetNewProdID = -1
    
    sSQL = "SELECT Max(mada.madasifr)+1 AS MaxOfProdID" & _
            " From mada"

    
    If ConnectRS(myConection, vRS, sSQL) = False Then
       ' WriteErrorLog "modRSProd", "GetNewProdID", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetNewProdID = 1
        GoTo RAE
    End If
    
    On Error Resume Next
    GetNewProdID = ReadField(vRS.Fields("MaxOfProdID"))
    
    If GetNewProdID < 1 Then
        GetNewProdID = 1
    End If
    
RAE:
    Set vRS = Nothing
    err.clear
End Function


Public Function GetProdBegInvStock(ByVal iProdID As Long) As Double
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    GetProdBegInvStock = -1
    
    sSQL = "SELECT mada.madazacs" & _
            " From mada" & _
            " WHERE madasifr='" & iProdID & "'"
            
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
        'WriteErrorLog "modRSProd", "GetProdBegInvStock", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GetProdBegInvStock = 0
        GoTo RAE
    End If

    GetProdBegInvStock = ReadField(vRS.Fields("madazacs"))
    
RAE:
    Set vRS = Nothing
End Function

Public Function SetProdBegInvStock(ByVal iProdID As Long, ByVal dNewBegInvStock As Double) As Boolean
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    'default
    SetProdBegInvStock = False
    
    sSQL = "SELECT *" & _
            " From mada" & _
            " WHERE madasifr='" & iProdID & "'"
    
    If ConnectRS(myConection, vRS, sSQL) = False Then
      '  WriteErrorLog "modRSProd", "SetProdBegInvStock", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    On Error GoTo RAE
    
    vRS.MoveFirst
    vRS.Fields("madazacs").Value = dNewBegInvStock
    vRS.Update
    
    SetProdBegInvStock = True
    
    'refresh inventory
    'Call modRSStockInv.ClearStockInvByProd(iProdID, CDate(0))
    
    
RAE:
    Set vRS = Nothing
End Function





Public Function ReadProd(ByRef vRS As ADODB.Recordset, ByRef vProd As tProd) As Boolean
    
    'default
    ReadProd = False
    
    On Error GoTo RAE
    
    With vProd
        
        .ProdID = ReadField(vRS.Fields("madasifr"))
        .ProdCode = ReadField(vRS.Fields("madaean"))
        .ProdDescription = ReadField(vRS.Fields("madanazi"))
        
        .FK_PackID = ReadField(vRS.Fields("madadoza"))
        .FK_CatID = ReadField(vRS.Fields("madagrup"))

        .SupPrice = ReadField(vRS.Fields("madanabc"))
        .SRPrice = ReadField(vRS.Fields("madampcd"))
        .BegInvStock = ReadField(vRS.Fields("madazacs"))
        
        .Active = ReadField(vRS.Fields("kontrola"))
        
        .RC = ReadField(vRS.Fields("madapd"))
        .RM = ReadField(vRS.Fields("madapdv"))
       .RCU = ReadField(vRS.Fields("madaenme"))
      .RMU = ReadField(vRS.Fields("madaminz"))
        
    End With
    
    ReadProd = True
    Exit Function
    
RAE:
    
End Function

Public Function WriteProd(ByRef vRS As ADODB.Recordset, ByRef vProd As tProd) As Boolean
    
    'default
    WriteProd = False
    
    On Error GoTo RAE

    With vProd
    
        vRS.Fields("madasifr") = .ProdID
        vRS.Fields("madaean") = .ProdCode
        vRS.Fields("madanazi") = .ProdDescription
        
        vRS.Fields("madadoza") = .FK_PackID
        vRS.Fields("madagrup") = .FK_CatID

        vRS.Fields("madanabc") = .SupPrice
        vRS.Fields("madampcd") = .SRPrice
        vRS.Fields("madazacs") = .BegInvStock
        
       vRS.Fields("kontrola") = .Active
        vRS.Fields("madapd") = .RC
        vRS.Fields("madapdv") = .RM
        vRS.Fields("madaenme") = .RCU
        vRS.Fields("madaminz") = .RMU

    End With

    WriteProd = True
    Exit Function
    
RAE:
    MsgBox err.Description
End Function
