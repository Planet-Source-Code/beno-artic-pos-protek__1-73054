Attribute VB_Name = "Mod_CON_DISP"
Option Explicit
Public Rs1 As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public CON As New CDbase
Public DCON As New ADODB.Connection
Public WebSQL As String
Public SQL As String
Public EDT As Boolean
Public CurUser As String
Public CURRDATE As String
Public CURRTIME As String
Public Destination As String
Public fsystem As FileSystemObject
Public Source As String
Public rptState As String
Public CatalogueName As String
Public ADDING As Boolean
Public MODIFYID As String
Public cst As Integer
Public Const DBFileName = "thesis.mdb"

Public Sub Mainx()

Dim frmLog         As frmLogin
Dim frmMd          As frmMAIN
Dim COUNT1      As Long
'If App.PrevInstance = True Then
'        MsgBox "The System Is In Use." & vbTab, vbInformation
'Else
'If App.PrevInstance = True Then Exit Sub
'    Set frmLog = New frmLogin
'        frmLog.Show vbModal
'If Not frmLog.OK Then
'        MsgBox "Unauthorized Validataion"
'End
'End If
  DBPathFileName = App.path & "\DATABASE\" & DBFileName
        PathFileName = App.path + "\DATABASE\Thesis.mdb"
If myConection.State = adStateOpen Then
Else
myConection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="

   frmsalesbill.Show

'   MsgBox "Error in Connecting Database please check Connection", vbCritical
End If
'frmSplash.Show
'Unload frmLog
'If nivo = 1 Then
'Load frmMAIN
'End If
'frmMAIN.Show
'frmSplash.Timer1.Enabled = True
'End If
End Sub

Public Sub Main()

Dim frmLog         As frmLogin
Dim frmMd          As frmMAIN
Dim COUNT1      As Long

If App.PrevInstance = True Then
       ' MsgBox "The System Is In Use." & vbTab, vbInformation
      
       'KillAppPrev "ProInv"
       ActivatePrevInstance
      
Else
If App.PrevInstance = True Then Exit Sub
If myConection.State = adStateOpen Then
Else
 DBPathFileName = App.path & "\DATABASE\" & DBFileName
        PathFileName = App.path + "\DATABASE\Thesis.mdb"

myConection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="

  ' frmsalesbill.Show

'   MsgBox "Error in Connecting Database please check Connection", vbCritical
End If
    
    Set frmLog = New frmLogin
        frmLog.Show vbModal
If Not frmLog.OK Then
'        MsgBox "Unauthorized Validataion"
End
End If
 'frmSplash.Show
Unload frmLog
If nivo = 1 Then
Load frmMAIN
End If
'frmMAIN.Show
frmSplash.Timer1.Enabled = True
End If
End Sub



Public Sub GetNewConnection2()
Dim sCNSTR As String

Set DCON = New Connection
DCON.CursorLocation = adUseClient

    sCNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.path + "\DATABASE\Thesis.mdb;"
    DCON.Open sCNSTR

'If DCON.State = adStateOpen Then
 '   Set GetNewConnection = DCON
'End If


End Sub
Public Sub GetNewConnection(ClassVar As Object)
'If TypeOf Classname Is CUpdate Then
   ' Set a = Classname
Set DCON = New ADODB.Connection
Set CON = ClassVar

CON.DBPath = App.path & "\Database\Thesis.mdb"
CON.OpenDb

End Sub

Public Function CMB1(ByVal TABLE As String, ByVal Field As String, cmb As ComboBox, Optional Clause As String, Optional ItemClear As Boolean)

If ItemClear = True Then

cmb.clear
End If

Call GetNewConnection2


Set Rs1 = New Recordset

If Clause = "" Then
Set Rs1 = DCON.Execute("Select * from " & TABLE & "")

Else
Set Rs1 = DCON.Execute("Select * from " & TABLE & " " & Clause)

End If

If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
       
        cmb.AddItem Rs1.Fields(Field)
       
        Rs1.MoveNext
    Wend

    
    
End If

Set Rs1 = Nothing
Set DCON = Nothing


End Function
Public Function CMB3(ByVal TABLE As String, ByVal Field As String, cmb As ComboBox, Optional Clause As String, Optional ItemClear As Boolean)

If ItemClear = True Then

cmb.clear
End If

Call GetNewConnection2


Set Rs1 = New Recordset

If Clause = "" Then
Set Rs1 = DCON.Execute("Select DISTINCT " & Field & " from " & TABLE & "")

Else
Set Rs1 = DCON.Execute("Select DISTINCT " & Field & "  from " & TABLE & " " & Clause)

End If

If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
       
        cmb.AddItem Rs1.Fields(Field)
       
        Rs1.MoveNext
    Wend

    
    
End If

Set Rs1 = Nothing
Set DCON = Nothing


End Function
Public Function CMB2(ByVal sqlArg As String, cmb As ComboBox)

cmb.clear
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(sqlArg)
If Rs1.RecordCount > 0 Then
    While Not Rs1.EOF
        cmb.AddItem Rs1.Collect(0)
       Rs1.MoveNext
    Wend
End If
Set Rs1 = Nothing
Set DCON = Nothing
End Function
Public Function Decimals(Key_Ascii As Integer, ByVal ControlName As Object, ByVal DecimalPlace As Integer)
On Error GoTo DECERR

Static DecPlace As Integer
If InStr(1, ControlName, ".") Then
    If Key_Ascii <> 13 And Key_Ascii <> 8 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    
    End If
    If DecPlace = 0 Then
    DecPlace = Val(Len(ControlName) + DecimalPlace)
    ControlName.MaxLength = DecPlace
   
    End If
Else
    DecPlace = 0
    If Key_Ascii <> 13 And Key_Ascii <> 8 And Key_Ascii <> 46 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    End If
End If
Exit Function

DECERR:
    MsgBox err.Description & vbTab, vbInformation
    
End Function
Public Function OFFCHar(Key_Ascii As Integer, ByVal ControlName As Object)
On Error GoTo DECERR


    If Key_Ascii <> 13 And Key_Ascii <> 8 Then
    If Key_Ascii < 48 Or Key_Ascii > 57 Then Key_Ascii = 0
    
    End If
  
 
Exit Function

DECERR:
    MsgBox err.Description & vbTab, vbInformation
    
End Function
Public Function offDefine(Key_Ascii As Integer, ByVal ControlName As Object, sFilter As String)

If InStr(sFilter, Chr(Key_Ascii)) = 0 Then
    Key_Ascii = 0
End If

End Function
Private Function WordTens(ByVal SNUM As Long) As String
Select Case SNUM
    Case 1
        WordTens = " Ena"
    Case 2
        WordTens = " Dva"
    Case 3
        WordTens = " Tri"
    Case 4
        WordTens = " Štiri"
    Case 5
        WordTens = " Pet"
    Case 6
        WordTens = " Šest"
    Case 7
        WordTens = " Sedem"
    Case 8
        WordTens = " Osem"
    Case 9
        WordTens = " Devet"
    Case 10
        WordTens = " Deset"
    Case 11
        WordTens = " enajst"
    Case 12
        WordTens = " dvanajst"
    Case 13
        WordTens = " Trinajst"
    Case 14
        WordTens = " Štirinajst"
    Case 15
        WordTens = " Petnajst"
    Case 16
        WordTens = " Šestnajst"
    Case 17
        WordTens = " Sedemnajst"
    Case 18
        WordTens = " Osemnajst"
    Case 19
        WordTens = " Devetnajst"
    Case 20
        WordTens = " Dvajset"
    Case 30
        WordTens = " Trideset"
    Case 40
        WordTens = " Štiridest"
    Case 50
        WordTens = " Petdeset"
    Case 60
        WordTens = " Šestdeset"
    Case 70
        WordTens = " Sedemdeset"
    Case 80
        WordTens = " Osemdeset"
    Case 90
        WordTens = " Devetdeset"
End Select
End Function


Public Function NumToWord(ByVal src_num As String) As String
Dim SNUM  As Double
SNUM = Val(src_num)
If SNUM > 999999999999999# Then
    NumToWord = "Error: To much number."
    Exit Function
End If
Dim WHOLE As String
Dim extra As String
Dim WORD  As String
Dim NWHOLE As Double

If InStr(1, str$(SNUM), ".", vbTextCompare) <> 0 Then
   WHOLE = Split(str$(SNUM), ".")(0)
    extra = Split(src_num, ".")(1)
Else
    WHOLE = SNUM
End If

If SNUM < 1 Then WORD = "Zero"

NWHOLE = Val(WHOLE)
'Check for One and Tens
If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
    WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
ElseIf Val(Right(NWHOLE, 2)) > 20 Then
    WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
    WORD = WORD & WordTens(Right(NWHOLE, 1))
End If
'Check for Hundred
If NWHOLE > 99 Then
   If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Sto" & WORD
End If
'Check for Thousand
If NWHOLE > 999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Tisoè" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Sto" & WORD
        End If
    End If
End If
'Check for Million
If NWHOLE > 999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 6))) & " Miljon" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 6), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 2, 1)) & " Million" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 99 Then
            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Sto" & WORD
        End If
    End If
End If
'Check for Billion
If NWHOLE > 999999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " Biljon" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " Billion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 99 Then
            If Left(Right(NWHOLE, 12), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 12), 1)) & " sto" & WORD
        End If
    End If
End If
'Check for Trillion
If NWHOLE > 999999999999# Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 12))) & " Triljon" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 12), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 2, 1)) & " Trillion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 99 Then
            If Left(Right(NWHOLE, 15), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 15), 1)) & " Sto" & WORD
        End If
    End If
End If
If extra = "" Then
    WORD = WORD & "   in   00/100"
Else
    If Val(extra) < 10 Then extra = "0" & extra
    WORD = WORD & "   in   " & extra & "/100"
End If
NumToWord = WORD

NWHOLE = 0
WORD = ""
extra = ""
WHOLE = ""
End Function

Public Function GRIDBIND(ByVal TABLE As String, ByVal Grid2 As DataGrid, Optional Clause As String)
Call GetNewConnection2
Set Rs1 = New Recordset
If Clause <> "" Then
Set Rs1 = DCON.Execute("Select * from " & TABLE & " " & Clause)
SQL = "Select * from " & TABLE & " " & Clause
Else
Set Rs1 = DCON.Execute("Select * from " & TABLE)
SQL = "Select * from " & TABLE
'Rs1.Open "Select * from " + Table, DCON, adOpenDynamic, adLockPessimistic
End If
  
Set Grid2.DataSource = Rs1



Set Rs1 = Nothing
Set DCON = Nothing


End Function

Public Function GRIDBINDx(ByVal Grid2 As DataGrid, Optional Clause As String)
Call GetNewConnection2
Set Rs1 = New Recordset

Set Rs1 = DCON.Execute(Clause)
SQL = Clause

Set Grid2.DataSource = Rs1



Set Rs1 = Nothing
Set DCON = Nothing


End Function
Public Sub GridRefresh()
Select Case CatalogueName
  
  Case "Customer"
'         Call GRIDBIND("Partner", frmControlMain.MSHFlexGrid1)

  Case "Supplier"
 '        Call GRIDBIND("partner", frmControlMain.MSHFlexGrid1, "WHERE SuppliersID<>'CASH'")
  Case "Category"
         Call GRIDBIND("mada", frmControlMain.MSHFlexGrid1)
  Case "Location"
'         Call GRIDBIND("grupa", frmControlMain.MSHFlexGrid1)
  Case "Purchase Order"
        Call GRIDBIND("PurchaseOrderHeader", frmControlMain.MSHFlexGrid1)
  Case "Purchase Return"
            Call GRIDBIND("PurchaseReturnHeader", frmControlMain.MSHFlexGrid1)
  Case "Purchase Registry"
      Call GRIDBIND("PurchaseRegistryHeader", frmControlMain.MSHFlexGrid1)
  Case "Sales Return"
            Call GRIDBIND("racusif", frmControlMain.MSHFlexGrid1)
  Case "Sales Registry"
             Call GRIDBIND("tdr", frmControlMain.MSHFlexGrid1)
    
End Select
End Sub
