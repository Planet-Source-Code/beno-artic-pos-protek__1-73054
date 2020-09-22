Attribute VB_Name = "modGeneral"
Option Explicit
Public kia, kib, kic As String
Public intFontSize As Integer
Public strFontFace As String
Public dblCellWidthTot As Double
Public intInv As Long
Public msearchResult
'registry
Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

   Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpdata As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpdata As Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpdata As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

'ends


Dim mboolAsc As Boolean
Dim mctlTxt As VB.TextBox
Public strConString(9) As New ADODB.Connection
Public boolJoin As Boolean
Public boolCancelled As Boolean
Public strError As String
Public boolConfirm
Public dbLocal As New ADODB.Connection
Public boolUnload As Boolean

Public strPrimaryTable As String
Public boolFromRun As Boolean, boolFromPrevious As Boolean
Public strReportTitle As String
Public strTemplateFileName As String
Public intDistictData As Integer
Public strAlias() As String
Public strDeleteCols() As String
Public strCalcField() As String
'wait
    Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'ends

'wait for given seconds
Public Function fWait(ByVal SecondsToWait As Double) 'Time In seconds
Dim EndTime As Double
EndTime = GetTickCount + SecondsToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
    DoEvents
Loop
End Function

'write to registry
Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
        
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key

    'open the specified key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_SET_VALUE, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub
Sub PrintFlexi(strReportTitle As String, msfFlexi As MSFlexGrid)
msfFlexi.Redraw = False
Dim intFontSize As Integer
Dim strFontFace As String
Dim dblCellWidthTot As Double
Dim intInv As Integer
Dim msearchResult

intFontSize = msfFlexi.Font.Size - 7
strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
On Error GoTo err:
Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report " & intInv & ".htm", True)
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "<body>"

Dim intCol As Integer
Dim intRow As Integer
msearchResult.WriteLine "<font size='2' face='Arial'><center><b>" & strReportTitle & "</center></b></font>"
msearchResult.WriteLine "<BR>"
msearchResult.WriteLine "<table border='1' width='100%' cellspacing='0' cellpadding='0'>"
intCol = 0
dblCellWidthTot = 0
msfFlexi.Row = 0

'calculate space
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    dblCellWidthTot = dblCellWidthTot + msfFlexi.CellWidth
    intCol = intCol + 1
Wend

'set col headings - ignore first contains A,B,C etc
intCol = 0
msfFlexi.Row = 1
msearchResult.WriteLine "<tr>"
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    msearchResult.WriteLine "<td align='center' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'><b> " & msfFlexi.Text & " </font></b></td>"
    intCol = intCol + 1
Wend
msearchResult.WriteLine "</tr>"

'add data - 2 since 1st col is added in above
intRow = 2
While intRow < msfFlexi.Rows
    msfFlexi.Row = intRow
    intCol = 0
    msearchResult.WriteLine "<tr>"
    While intCol < msfFlexi.Cols
        msfFlexi.Col = intCol
        msearchResult.WriteLine "<td align='left' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'>&nbsp;" & msfFlexi.Text & " </font></td>"
        intCol = intCol + 1
    Wend
    msearchResult.WriteLine "</tr>"
    intRow = intRow + 1
Wend


msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume
End Sub

'to write to registry
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
    Case REG_SZ
        sValue = vValue & Chr$(0)
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    
    Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function

'To read from registry
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
     Dim cch As Long
     Dim lrc As Long
     Dim lType As Long
     Dim lValue As Long
     Dim sValue As String

     On Error GoTo QueryValueExError

     ' Determine the size and type of data to be read
     lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
     If lrc <> ERROR_NONE Then Error 5

     Select Case lType
         ' For strings
         Case REG_SZ:
             sValue = String(cch, 0)

 lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
             If lrc = ERROR_NONE Then
                 vValue = Left$(sValue, cch - 1)
             Else
                 vValue = Empty
             End If
         ' For DWORDS
         Case REG_DWORD:
 lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
             If lrc = ERROR_NONE Then vValue = lValue
         Case Else
             'all other data types not supported
             lrc = -1
     End Select

QueryValueExExit:
     QueryValueEx = lrc
     Exit Function

QueryValueExError:
     Resume QueryValueExExit
 End Function
   
'To read from registry
Function QueryValue(sKeyName As String, sValueName As String) As String
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value
    
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function

Function testDBConnection(strDBType As String, strServer As String, strDB As String, strUID As String, strPWD As String) As Boolean
On Error GoTo err:
strError = ""
Dim dbTest As New ADODB.Connection
testDBConnection = True
If UCase(strDBType) <> "MS ACCESS" And UCase(strDBType) <> "MS SQL SERVER" And UCase(strDBType) <> "ORACLE" And UCase(strDBType) <> "MYSQL" And UCase(strDBType) <> "POSTGRESQL" Then
    MsgBox "The database type " & strDBType & " is not supported", vbExclamation
    testDBConnection = False
    Exit Function
Else
    If UCase(strDBType) = "MS ACCESS" Then
        dbTest.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB & ";Jet OLEDB:Database Password=" & strPWD)
    ElseIf UCase(strDBType) = "MS SQL SERVER" Then
        dbTest.Open ("Provider=SQLOLEDB;data Source=" & strServer & ";Initial Catalog=" & strDB & ";User Id=" & strUID & ";Password=" & strPWD & ";")
    ElseIf UCase(strDBType) = "ORACLE" Then
        dbTest.Open "DRIVER={Microsoft ODBC For Oracle};UID=" & strUID & ";PWD=" & strPWD & ";SERVER=" & strServer
    ElseIf UCase(strDBType) = "MYSQL" Then
        dbTest.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strServer & ";DATABASE=" & strDB & ";UID=" & strUID & ";PWD=" & strPWD
    ElseIf UCase(strDBType) = "POSTGRESQL" Then
        dbTest.Open "DRIVER={PostgreSQL Unicode};SERVER=" & strServer & ";DATABASE=" & strDB & ";UID=" & strUID & ";PWD=" & strPWD
    Else
        testDBConnection = False
    End If
End If
Exit Function

err:
strError = err.Description
testDBConnection = False
End Function

Sub unloadAllForms()
boolUnload = True
Dim frmForm As Form
For Each frmForm In Forms
    Unload frmForm
    Set frmForm = Nothing
Next
For Each frmForm In Forms
    Unload frmForm
    Set frmForm = Nothing
Next
Set dbLocal = Nothing
End Sub

Sub delDups(lstBox As ListBox)
Dim lngX As Long, lngY As Long
Dim strTemp As String
strTemp = "": lngX = 0: lngY = 0

While lngX <= lstBox.ListCount
    strTemp = lstBox.List(lngX)
    lngY = 0
    While lngY <= lstBox.ListCount
        If UCase(lstBox.List(lngY)) = UCase(lstBox.List(lngX)) And lngX <> lngY Then
            lstBox.RemoveItem lngY
            lngY = lngY - 1
        End If
        lngY = lngY + 1
    Wend
    lngX = lngX + 1
Wend
End Sub

Function chkListMatch(lstBox As ListBox, strVal As String) As Boolean
chkListMatch = False
Dim intX As Long
While intX <= lstBox.ListCount
    If UCase(lstBox.List(intX)) = UCase(strVal) Then
        chkListMatch = True
        Exit Function
    End If
    intX = intX + 1
Wend
End Function

Function chkComboMatch(cboBox As ComboBox, strVal As String) As Boolean
chkComboMatch = False
Dim intX As Long
While intX <= cboBox.ListCount
    If UCase(cboBox.List(intX)) = UCase(strVal) Then
        chkComboMatch = True
        Exit Function
    End If
    intX = intX + 1
Wend
End Function

Function dataType(intType As Long) As String
   If CInt(intType) = 3 Or CInt(intType) = 139 Then
        dataType = "Long"
    ElseIf CInt(intType) = 6 Then
        dataType = "Currency"
    ElseIf CInt(intType) = 7 Or CInt(intType) = 135 Then
        dataType = "Date"
    ElseIf CInt(intType) = 11 Then
        dataType = "YesNo"
    ElseIf CInt(intType) = 203 Then
        dataType = "Memo"
    Else
        dataType = "VarChar"
    End If
End Function

Sub SortFlexiNoArrows(MSFGrid As MSFlexGrid, boolLastRowBlank As Boolean, Optional intColNo As Integer)
With MSFGrid
.FormatString = Replace(.FormatString, " (+)", "")
.FormatString = Replace(.FormatString, " (-)", "")

'set the col no if passed as parameter
If intColNo > 0 And intColNo <= MSFGrid.Cols Then
    MSFGrid.Col = intColNo
End If

'remove blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows - 1
End If

'sort
If mboolAsc = False Then
    .Sort = 6
    mboolAsc = True
    .Row = 0
    .Text = .Text & " (-)"
Else
    .Sort = 7
    mboolAsc = False
    .Row = 0
    .Text = .Text & " (+)"
End If

'add blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows + 1
End If

Call AltFlexiColors(MSFGrid, 2, 1)
End With
End Sub

Sub SortFlexiArrows(MSFGrid As MSFlexGrid, boolLastRowBlank As Boolean, boolAltFlexiColors As Boolean, Optional sortColNo As Integer)
With MSFGrid
'generate text box randomly
On Error GoTo err
Set mctlTxt = MSFGrid.Parent.Controls.Add("VB.TextBox", "txt_txt_txt")
Set mctlTxt.Container = MSFGrid.Container
mctlTxt.Appearance = 0
mctlTxt.BorderStyle = 0
mctlTxt.Font = "Wingdings"
mctlTxt.Text = "â"
mctlTxt.BackColor = MSFGrid.BackColorFixed
mctlTxt.ZOrder (0)
mctlTxt.TabStop = False

mctlTxt.Height = 175
mctlTxt.Width = 175
mctlTxt.Locked = True
mctlTxt.Enabled = False
mctlTxt.Visible = False

'set text box top
mctlTxt.Top = MSFGrid.Top + 60
'set the text box left
mctlTxt.Left = MSFGrid.Left + MSFGrid.CellLeft + MSFGrid.CellWidth - 225

'set the col no if passed as parameter
If sortColNo > 0 And sortColNo <= MSFGrid.Cols Then
    MSFGrid.Col = sortColNo
End If

'remove blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows - 1
End If

'sort
If mboolAsc = False Then
    .Sort = 6
    .Row = 0
    mctlTxt.Text = "á"
    mboolAsc = True
Else
    .Sort = 7
    .Row = 0
    mctlTxt.Text = "â"
    mboolAsc = False
End If

'add blank row
If boolLastRowBlank = True Then
    MSFGrid.Rows = MSFGrid.Rows + 1
End If

mctlTxt.Visible = True

If boolAltFlexiColors = True Then
    Call AltFlexiColors(MSFGrid, 2, 1)
End If
Exit Sub

err:
Call fWait(0.2)
MSFGrid.Parent.Controls.Remove "txt_txt_txt"
Resume
End With
End Sub

Sub AltFlexiColors(MSFlexi As MSFlexGrid, StartRow As Integer, startCol As Integer)
Dim intX As Integer
Dim intY As Integer
Dim boolY As Boolean

boolY = True
intX = StartRow
intY = startCol
MSFlexi.Redraw = False
While intX < MSFlexi.Rows
    MSFlexi.Row = intX
    While intY < MSFlexi.Cols
        MSFlexi.Col = intY
        If boolY = True Then
            MSFlexi.CellBackColor = &HC0FFFF
        Else
            MSFlexi.CellBackColor = vbWhite
        End If
        intY = intY + 1
    Wend
    If boolY = True Then
        boolY = False
    Else
        boolY = True
    End If
    intY = startCol
    intX = intX + 1
Wend
MSFlexi.Redraw = True
End Sub

Public Function FG_AutosizeCols(myGrid As MSHFlexGrid, frmForm As Form, _
                                Optional ByVal lFirstCol As Long = -1, _
                                Optional ByVal lLastCol As Long = -1, _
                                Optional bCheckFont As Boolean = False)
  
  Dim lCol As Long, lRow As Long, lCurCol As Long, lCurRow As Long
  Dim lCellWidth As Long, lColWidth As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
  myGrid.Redraw = False
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = frmForm.FontBold
    sFontName = frmForm.FontName
    dFontSize = frmForm.FontSize
  End If
  
  With myGrid
    If bCheckFont Then
      lCurRow = .Row
      lCurCol = .Col
    End If
    
    If lFirstCol = -1 Then lFirstCol = 0
    If lLastCol = -1 Then lLastCol = .Cols - 1
    
    For lCol = lFirstCol To lLastCol
      lColWidth = 0
      If bCheckFont Then .Col = lCol
      For lRow = 1 To .Rows - 1
        If bCheckFont Then
          .Row = lRow
          frmForm.FontBold = .CellFontBold
          frmForm.FontName = .CellFontName
          frmForm.FontSize = .CellFontSize
        End If
        lCellWidth = frmForm.TextWidth(Trim(.TextMatrix(lRow, lCol)))
        
        If lCellWidth > lColWidth Then lColWidth = lCellWidth
        
      Next lRow
        If lColWidth <> 0 Then
      .ColWidth(lCol) = lColWidth + frmForm.TextWidth("W")
      Else
      .ColWidth(lCol) = 0
      End If
    Next lCol
    
    If bCheckFont Then
      .Row = lCurRow
      .Col = lCurCol
    End If
  End With
  
  If bCheckFont Then
    ' restore the forms font settings
    frmForm.FontBold = bFontBold
    frmForm.FontName = sFontName
    frmForm.FontSize = dFontSize
  End If

'Call SortFlexiArrows(myGrid, True, 1)
myGrid.Redraw = True
End Function

Public Function FG_AutosizeRows(myGrid As MSFlexGrid, frmForm As Form, _
                                Optional ByVal lFirstRow As Long = -1, _
                                Optional ByVal lLastRow As Long = -1, _
                                Optional bCheckFont As Boolean = False)
                                
  ' This will only work for Cells with a Chr(13)
  ' To have it working with WordWrap enabled
  ' you need some other routine
  ' Which has been added too
  myGrid.Redraw = False
  Dim lCol As Long, lRow As Long, lCurCol As Long, lCurRow As Long
  Dim lCellHeight As Long, lRowHeight As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
  
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = frmForm.FontBold
    sFontName = frmForm.FontName
    dFontSize = frmForm.FontSize
  End If
  
  With myGrid
    If bCheckFont Then
      lCurCol = .Col
      lCurRow = .Row
    End If
    
    If lFirstRow = -1 Then lFirstRow = 0
    If lLastRow = -1 Then lLastRow = .Rows - 1
    
    For lRow = lFirstRow To lLastRow
      lRowHeight = 0
      If bCheckFont Then .Row = lRow
      For lCol = 0 To .Cols - 1
        If bCheckFont Then
          .Col = lCol
          frmForm.FontBold = .CellFontBold
          frmForm.FontName = .CellFontName
          frmForm.FontSize = .CellFontSize
        End If
        lCellHeight = frmForm.TextHeight(.TextMatrix(lRow, lCol))
        If lCellHeight > lRowHeight Then lRowHeight = lCellHeight
      Next lCol
      .RowHeight(lRow) = lRowHeight + frmForm.TextHeight("Wg") / 5
    Next lRow
    
    If bCheckFont Then
      .Row = lCurRow
      .Col = lCurCol
    End If
  End With
  
  If bCheckFont Then
    ' restore the forms font settings
    frmForm.FontBold = bFontBold
    frmForm.FontName = sFontName
    frmForm.FontSize = dFontSize
  End If
myGrid.Redraw = True
End Function

Public Function FG_RemoveColumn(myGrid As MSFlexGrid, ByVal lColumn As Long)
  With myGrid
    .Redraw = False
    If lColumn < .Cols Then
      .ColPosition(lColumn) = .Cols - 1
      .Cols = .Cols - 1
    End If
    .Redraw = True
  End With
End Function

Function chkArrayDups(objArr() As Variant, strString As String, boolCaseSensitive As Boolean) As Boolean
chkArrayDups = False
Dim lngX As Long
lngX = 0
While lngX <= UBound(objArr)
    If boolCaseSensitive = True Then
        If objArr(lngX) = strString Then
            chkArrayDups = True
        End If
    ElseIf boolCaseSensitive = False Then
        If UCase(objArr(lngX)) = UCase(strString) Then
            chkArrayDups = True
        End If
    End If
    lngX = lngX + 1
Wend
End Function

Public Function fchkFolderPath(strFilePath As String, boolCreateFolder As Boolean) As Boolean
If (fso.FolderExists(strFilePath)) Then
    fchkFolderPath = True
Else
    fchkFolderPath = False
    If boolCreateFolder = True Then
        On Error GoTo err:
        fso.CreateFolder (strFilePath)
        fchkFolderPath = True
    End If
End If
Exit Function

err:
fchkFolderPath = False
End Function
Public Function dopr()
Set msearchResult = Nothing
Set msearchResult = fso.OpenTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", ForAppending, False, TristateUseDefault)

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "</TR>"
End Function
Public Function dono2(wid As String, pos As String, vel As String, boldi As String, rezul As String, presl As String)
Set msearchResult = Nothing
Set msearchResult = fso.OpenTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", ForAppending, False, TristateUseDefault)
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='" & presl & "%'></TD>"
msearchResult.WriteLine "<TD WIDTH='" & Trim(wid) & "%' VAlign=" & pos & ">" & boldi & "<FONT SIZE=" & vel & " FACE='Helvetica'>" & rezul & "<BR></FONT>" & IIf(boldi = "", "", "</B>") & "</TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End Function
Public Function dono(wid As String, pos As String, vel As String, boldi As String, rezul As String)
Set msearchResult = Nothing
Set msearchResult = fso.OpenTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", ForAppending, False, TristateUseDefault)

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='" & Trim(wid) & "%' VAlign=" & pos & ">" & boldi & "<FONT SIZE=" & vel & " FACE='Helvetica'>" & rezul & "<BR></FONT>" & IIf(boldi = "", "", "</B>") & "</TD>"
msearchResult.WriteLine "</TR>"
'msearchResult.WriteLine "</TABLE>"
End Function
Public Function dois(wid As String, pos As String, vel As String, boldi As String, rezul As String)
Set msearchResult = Nothing
Set msearchResult = fso.OpenTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", ForAppending, False, TristateUseDefault)

msearchResult.WriteLine "<TD WIDTH='" & Trim(wid) & "%' VAlign=" & pos & ">" & boldi & "<FONT SIZE=" & vel & " FACE='Helvetica'>" & rezul & "<BR></FONT>" & IIf(boldi = "", "", "</B>") & "</TD>"
End Function
Public Sub glava_izp()
'Set msearchResult = FSO.CreateTextFile(App.path & "\Print\Report" & Str(intInv) & ".htm", False)
Set msearchResult = Nothing
Set msearchResult = fso.OpenTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", ForAppending, False, TristateUseDefault)
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Left VAlign=Middle><B><FONT SIZE=3 FACE='Helvetica'>" & Getnazi("select glava1" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Left VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select glava2" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Left VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select glava3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Left VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"

msearchResult.WriteLine "</TABLE>"
End Sub
Public Sub crta_izp()
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End Sub
Sub Print_dob(strReportTitle As String)

'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv like '%" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' ORDER BY SIFRA"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
Call glava_izp
Call crta_izp
Call dopr
Call dono("30", "Left", "2", "<B>", "Dobavitelj")

Call dono2("30", "Left", "2", "", Getnazi("select dod0" & kib), "4")
Call dono2("30", "Left", "2", "", Getnazi("select ulica from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")
Call dono2("30", "Left", "2", "", Getnazi("select mesto from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")
Call dono2("30", "Left", "2", "", Getnazi("select davcna from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")

Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "3", "<B>", Getnazi("select naz_do" & kia) & tip_dok & xid_dok)
Call dono("30", "Left", "2", "<B>", "")
Call dono2("30", "Left", "2", "", "DATUM: " & Getnazi("select datum" & kic), "68")
Call dono2("30", "Left", "2", "", "Št. Dobavnice: " & Getnazi("select dod2" & kib), "60")
Call dono("30", "Left", "2", "", Getnazi("select opis" & kib))
Call dono("30", "Left", "2", "", "")
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='11%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select pop" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select znes" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Call crta_izp
If rs.State = 1 Then rs.Close
rs.Open "select * " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb, koli, popu, pcc As Double
skupii = 0
pcc = 0
ddva = 0
ddvb = 0
koli = 0
popu = 0
'do while pozicije glavni
Do While Not rs.EOF
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("sifra") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("naziv"), 35) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='11%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("kol"), 3) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select madaenme from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("cena"), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("pop"), 2) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber((rs.Fields("cena") * (1 - (rs.Fields("pop") / 100))), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

koli = koli + rs.Fields("kol")
popu = popu + ((rs.Fields("cena") - (rs.Fields("cena") * (1 - (rs.Fields("pop") / 100)))) * rs.Fields("kol"))
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") <> "" Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'></TD>"
msearchResult.WriteLine "<TD WIDTH='57%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
ff = ff + 1

skupii = skupii + ((rs.Fields("cena") * (1 - (rs.Fields("pop") / 100))) * rs.Fields("kol"))
pcc = pcc + (rs.Fields("cena") * rs.Fields("kol"))



rs.MoveNext
Loop


End If
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='11%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(koli, 3) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(pcc, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(popu, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(skupii, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

Call crta_izp

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

End Sub

Sub Print_zal(strReportTitle As String, skladd As String, arti As String)
If rs.State = 1 Then rs.Close
myConection.Execute "delete from zaloga where kol=0"
Dim pogo As String
pogo = ""
If skladd <> "" Then
pogo = " where skl='" & skladd & "'"
If arti <> "" Then
pogo = pogo & " and sifra='" & arti & "'"
End If
End If

If skladd = "" And arti <> "" Then
pogo = " where sifra='" & arti & "'"
End If

rs.Open "select sifra, naziv,kol,vrednost,tip_dok,id_dok from zaloga " & pogo & " order by sifra", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
End If
Dim nisif As String
nisif = ""

Do While Not rs.EOF
If Getnazi("select madasifr from mada where madasifr='" & rs.Fields("sifra") & "'") = "" Then
nisif = nisif & rs.Fields("sifra") & rs.Fields("naziv") & str(rs.Fields("vrednost")) & ","
End If
rs.MoveNext
Loop
If nisif <> "" Then
MsgBox nisif, vbInformation, "Ne najdem identa v Artiklih!!!"
End If
'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:
kia = " from izpisi where naziv='" & strReportTitle & "'"
kib = " from glavna where id_dok='" & xid_dok & "'"
kic = " from nabasif where id_dok='" & xid_dok & "' group by id_dok"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)

 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<HEAD>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "</HEAD>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
Call glava_izp
Call crta_izp
Call dopr
Call dono("30", "Middle", "3", "<B>", Getnazi("select naz_do" & kia))

Call dono("30", "Left", "2", "<B>", "Za dan: " & frmControlMain.DATDO.Value)
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")

Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='45%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select znes" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Call crta_izp
If rs.State = 1 Then rs.Close
rs.Open "select * from mada where tip_art='MAT' order by madagrup,dobavit_id ", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb, koli, popu, skgr As Double
skupii = 0
ddva = 0
ddvb = 0
skgr = 0
koli = 0
popu = 0
'do while pozicije glavni
Dim grupp, gruu, zz, skuu As Integer
grupp = 0
gruu = 0
zz = 0
skuu = 0
skuu = Getnazi("select count(madasifr) as xx from mada where tip_art='MAT'")


Dim xxn, xnazz As String
Do While Not rs.EOF
zz = zz + 1
If zz < skuu Then
izpisi.ProgressBar.Value = zz / skuu * 100
End If
If Getnazi("select cena from nabasif where tip_dok='NA' and poknj='K' and sifra='" & rs.Fields("madasifr") & "' order by datum desc") <> "" Then
rs.Fields("madanabc") = FormatNumber(Getnazi("select (cena* (1 - (pop / 100))) as xx from nabasif where poknj='K' and tip_dok='NA'  and sifra='" & rs.Fields("madasifr") & "' order by datum desc"), 4)
Else
rs.Fields("madanabc") = 0
End If
If Getnazi("select sum(kol) as pros from zaloga where sifra='" & rs.Fields("madasifr") & "' group by sifra") <> "" Then
rs.Fields("madazalo") = FormatNumber(Getnazi("select sum(kol) as pros from zaloga where sifra='" & rs.Fields("madasifr") & "' group by sifra"), 3)
Else
rs.Fields("madazalo") = 0
End If
rs.Update
grupp = rs.Fields("madagrup")
If grupp <> gruu Then
If skgr <> 0 Then
Call crta_izp
'Call dono("100", "Right", "1", "<B>", FormatNumber(skgr, 4))
Call dono2("30", "Right", "2", "", "Skupaj grupa: " & FormatNumber(skgr, 4), "70")

skgr = 0
Call crta_izp
End If
Call dono("30", "Left", "2", "<B>", Getnazi("select grupa from grupa where sifra=" & grupp))
Call crta_izp
End If
gruu = rs.Fields("madagrup")
'If UCase(Left(RS.Fields("madanazi"), 4)) = "ROLE" Then
xxn = "SELECT dokm.tekst FROM nabasif INNER JOIN dokm ON (nabasif.pozicija = dokm.atribut) AND (nabasif.id_dok = dokm.id_dok) AND (nabasif.tip_dok = dokm.tip_dok) where nabasif.sifra='" & rs.Fields("madasifr") & "' order by nabasif.datum desc"
'MsgBox xxn
xnazz = Trim(Getnazi("SELECT dokm.tekst FROM nabasif INNER JOIN dokm ON (nabasif.pozicija = dokm.atribut) AND (nabasif.id_dok = dokm.id_dok) AND (nabasif.tip_dok = dokm.tip_dok) where nabasif.sifra='" & rs.Fields("madasifr") & "' order by nabasif.datum desc"))
'Else
'xnazz = ""
'End If
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("madasifr") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("dobavit_id"), 10) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='45%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("madanazi"), 45) & "  " & xnazz & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("madazalo"), 3) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='4%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("madaenme") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("madanabc"), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("madanabc") * rs.Fields("madazalo"), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

skupii = skupii + (rs.Fields("madanabc") * rs.Fields("madazalo"))
skgr = skgr + (rs.Fields("madanabc") * rs.Fields("madazalo"))



rs.MoveNext
Loop

End If
If skgr <> 0 Then
Call crta_izp
'Call dono("100", "Right", "1", "<B>", FormatNumber(skgr, 4))
Call dono2("30", "Right", "2", "", "Skupaj grupa: " & FormatNumber(skgr, 4), "70")

skgr = 0
'Call crta_izp
End If
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='20%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'><BR>" & "Vrednost zaloge:" & "</FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(skupii, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

End Sub

Sub Print_preg(strReportTitle As String)

'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:
Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)

Dim das, des
das = Format(frmControlMain.DATOD.Value, "dd.mm.yyyy")
des = Format(frmControlMain.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv like '%" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and datum between #" & dod & "# and #" & ddo & "# group by id_dok"

 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
Call glava_izp
Call crta_izp
Call dopr
Call dono("30", "Middle", "3", "<B>", Getnazi("select naz_do" & kia) & " za obdobje od " & das & " do " & des)

Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")

Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
If tip_dok = "IZ" Then
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Tip_dok<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Id_dok<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' Align=Middle VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Opis<BR></FONT></B></TD>"

Else
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' Align=Middle VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
End If
msearchResult.WriteLine "<TD WIDTH='10%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Call crta_izp
If rs.State = 1 Then rs.Close
rs.Open "select max(datum) as datum,id_dok,sum(kol*(cena*(1-(pop/100)))) as znesek " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb, koli, popu As Double
skupii = 0
ddva = 0
ddvb = 0
koli = 0
popu = 0
'do while pozicije glavni
Do While Not rs.EOF
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & rs.Fields("id_dok") & "'"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
If tip_dok = "IZ" Then
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & tip_dok & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("id_dok") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' Align=Middle VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kib) & "<BR></FONT></B></TD>"

Else
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select sifra from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select dod0" & kib) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' Align=Middle VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kib) & "<BR></FONT></B></TD>"
End If
msearchResult.WriteLine "<TD WIDTH='10%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("datum") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("ZNESEK"), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"


skupii = skupii + rs.Fields("ZNESEK")


rs.MoveNext
Loop

End If
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='20%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'><BR>" & "SKUPAJ :" & "</FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(skupii, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

End Sub
Sub PrintFlexix(strReportTitle As String)
'msfFlexi.Redraw = False


'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv='" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<HEAD>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "</HEAD>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR><TD>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><IMG SRC='" & App.path & "\gaber.jpg'></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'naziv firme
'do while
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava1" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava2" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"

msearchResult.WriteLine "</TABLE>"
'konec naziva

'kupec id
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='17%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idk" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='12%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select sifra from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idp" & kia) & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='5%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'kupec IDST

msearchResult.WriteLine "</TABLE><TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><B><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select naziv from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select ulica from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select posta from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "   " & Getnazi("select mesto from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select davcna from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'desna stran
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='35%'></TD>"
msearchResult.WriteLine "<TD WIDTH='25%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dod0" & kib) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='5%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'datum + line
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dat" & kia) & "  " & Getnazi("select datum" & kic) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'nazil listine + izdelal
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='2%'></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=3 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='17%'></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select prod" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dod5" & kib) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH=2%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & Getnazi("select naz_do" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH=1%></TD>"
msearchResult.WriteLine "<TD WIDTH=19% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & tip_dok & xid_dok & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH=17%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select dobav" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH=0%></TD>"
msearchResult.WriteLine "<TD WIDTH=43% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select dod6" & kib) & "</FONT></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH=57%></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE></TD>"



'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'opisi poz
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Top><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select zap" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' Align=Right VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select pop" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='4%' Align=Right VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ddv" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH='7%'></TD>"
msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select znes" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
If rs.State = 1 Then rs.Close
rs.Open "select * " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb As Double
skupii = 0
ddva = 0
ddvb = 0
'do while pozicije glavni
Do While Not rs.EOF
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & str(ff) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & rs.Fields("naziv") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("kol"), 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select madaenme from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("cena"), 4) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=1>&nbsp<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='5%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("znes"), 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'do while pozicije dokm
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") <> "" Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'></TD>"
msearchResult.WriteLine "<TD WIDTH='57%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
ff = ff + 1
If Val(Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'")) = 20 Then
ddva = ddva + rs.Fields("znes")
End If
If Val(Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'")) = 8.5 Then
ddvb = ddvb + rs.Fields("znes")
End If
skupii = skupii + rs.Fields("znes")
rs.MoveNext
Loop
ddva = (ddva * 1.2) - ddva
ddvb = (ddvb * 1.085) - ddvb
End If
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"


msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup1" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(skupii, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

If ddva <> 0 Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup2" & kia) & " 20 %" & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddva, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
If ddvb <> 0 Then

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup2" & kia) & " 8.5 %" & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddvb, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If


msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddva + ddvb + skupii, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>  <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'dobavnice
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='opis'") <> "" Then

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'>Veza/opis: </TD>"
msearchResult.WriteLine "<TD WIDTH='87%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='opis'") & "<BR></FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If



msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> DIREKTOR: Mitja Lešnik <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"

frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume
End Sub
Sub Printkol(strReportTitle As String)
'msfFlexi.Redraw = False
Dim intFontSize As Integer
Dim strFontFace As String
Dim dblCellWidthTot As Double
Dim intInv As Long
Dim msearchResult

'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:
Dim kia, kib, kic As String
kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv='" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<HEAD>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "</HEAD>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR><TD>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><IMG SRC='" & App.path & "\gaber.bmp'></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'naziv firme
'do while
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava1" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava2" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"

msearchResult.WriteLine "</TABLE>"
'konec naziva

'kupec id
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='17%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idk" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select sifra from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idp" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='5%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'kupec IDST

msearchResult.WriteLine "</TABLE><TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='17%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select davcna from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"
msearchResult.WriteLine "<TD WIDTH='17%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idk" & kia) & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='2%'></TD>"
msearchResult.WriteLine "<TD WIDTH='11%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select sifra from partner where naziv='" & Getnazi("select dod1" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='10%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'desna stran
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='35%'></TD>"
msearchResult.WriteLine "<TD WIDTH='25%'></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dod0" & kib) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='5%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'datum + line
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dat" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'nazil listine + izdelal
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='2%'></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=3 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='17%'></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select prod" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>""<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH=2%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & Getnazi("select naz_do" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH=1%></TD>"
msearchResult.WriteLine "<TD WIDTH=19% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & tip_dok & xid_dok & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH=17%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select dobav" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH=0%></TD>"
msearchResult.WriteLine "<TD WIDTH=43% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select datum" & kic) & "</FONT></TD>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"

msearchResult.WriteLine "<TD WIDTH=57%></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE></TD>"
msearchResult.WriteLine "<TD WIDTH=1%></TD>"
msearchResult.WriteLine "</TR>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD></TD>"
msearchResult.WriteLine "<TD></TD>"
msearchResult.WriteLine "<TD></TD>"
msearchResult.WriteLine "<TD></TD>"
msearchResult.WriteLine "<TD></TD>"

'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'opisi poz
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Top><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select zap" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='4%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH='7%'></TD>"
msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select em" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
If rs.State = 1 Then rs.Close
rs.Open "select * " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
'do while pozicije glavni
Do While Not rs.EOF
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & str(ff) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & rs.Fields("naziv") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=1>&nbsp<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='5%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & rs.Fields("kol") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select madaenme from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'do while pozicije dokm
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") <> "" Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'></TD>"
msearchResult.WriteLine "<TD WIDTH='57%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
ff = ff + 1
rs.MoveNext
Loop
End If
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume
End Sub

Public Function GetVirtualFileName(strFilePath As String) As String
Dim arrFile() As String
If Len(strFilePath) > 0 Then
    arrFile = Split(strFilePath, "\")
    GetVirtualFileName = arrFile(UBound(arrFile))
Else
    GetVirtualFileName = ""
End If
Exit Function

err:
GetVirtualFileName = ""
End Function

Function chkFileExtension(strFileInfo As String) As String
If strFileInfo <> "" Then
    chkFileExtension = Mid(strFileInfo, Len(strFileInfo) - 2, 3)
End If
End Function

Public Function chkFilePath(strFilePath As String) As Boolean
If (fso.FileExists(strFilePath)) Then
    chkFilePath = True
Else
    chkFilePath = False
End If
End Function

Public Sub ShowAppHelp(Optional IndexID As Integer)

     On Error GoTo err

    With frmPrint.CD
        .HelpFile = App.HelpFile
        
        If IndexID = 0 Then
            .HelpCommand = cdlHelpContents
        Else
            .HelpContext = IndexID
            .HelpCommand = cdlHelpContext
        End If
        
        .ShowHelp
    End With
    
    Exit Sub
 
err:
   MsgBox err.Description, vbExclamation

End Sub

Sub createLocalDB()

End Sub




Sub Print_osn(strReportTitle As String, msfFlexi As MSHFlexGrid)
msfFlexi.Redraw = False
Dim intFontSize As Integer
Dim strFontFace As String
Dim dblCellWidthTot As Double
Dim intInv As Integer
Dim msearchResult

intFontSize = msfFlexi.Font.Size - 7
strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where naziv='" & strReportTitle & "'"
kib = " from glavna where id_dok='" & xid_dok & "'"
kic = " from zaloga where id_dok='" & xid_dok & "' group by id_dok"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "<body>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Dim intCol As Integer
Dim intRow As Integer
msearchResult.WriteLine "<font size='2' face='Arial'><center><b>" & strReportTitle & "</center></b></font>"
msearchResult.WriteLine "<BR>"
msearchResult.WriteLine "<table border='1' width='100%' cellspacing='0' cellpadding='0'>"
intCol = 0
dblCellWidthTot = 0
msfFlexi.Row = 0
Dim aaddd As Integer
aaddd = 0
'calculate space
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    'MsgBox msfFlexi.CellWidth
    aaddd = msfFlexi.CellWidth
    If aaddd < 0 Then
    aaddd = 0
    End If
    dblCellWidthTot = dblCellWidthTot + aaddd
    intCol = intCol + 1
Wend

'set col headings - ignore first contains A,B,C etc
intCol = 1
msfFlexi.Row = 0
msearchResult.WriteLine "<tr>"
While intCol < msfFlexi.Cols
    msfFlexi.Col = intCol
    msearchResult.WriteLine "<td align='center' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'><b> " & msfFlexi.Text & " </font></b></td>"
    intCol = intCol + 1
Wend
msearchResult.WriteLine "</tr>"

'add data - 2 since 1st col is added in above
intRow = 1
While intRow < msfFlexi.Rows
    msfFlexi.Row = intRow
    intCol = 1
    msearchResult.WriteLine "<tr>"
    While intCol < msfFlexi.Cols
        msfFlexi.Col = intCol
        msearchResult.WriteLine "<td align='left' width=" & Format(msfFlexi.CellWidth * 100 / dblCellWidthTot, "##0.00") & "%><font size='" & intFontSize & "' face='" & strFontFace & "'>&nbsp;" & msfFlexi.Text & " </font></td>"
        intCol = intCol + 1
    Wend
    msearchResult.WriteLine "</tr>"
    intRow = intRow + 1
Wend


msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

'frmPrint.WebBrowser1.Navigate (App.path & "\Print\osnovni.htm")
'Set msearchResult = Nothing
'frmPrint.Show
'msfFlexi.Redraw = True
'Exit Sub

'err:
'Resume
End Sub






Sub Print_zal_fifo(strReportTitle As String)

'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where naziv='" & strReportTitle & "'"
kib = " from glavna where id_dok='" & xid_dok & "'"
kic = " from zaloga where id_dok='" & xid_dok & "' group by id_dok"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<HEAD>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "</HEAD>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
Call glava_izp
Call crta_izp
Call dopr
Call dono("30", "Middle", "3", "<B>", Getnazi("select naz_do" & kia))
Call dono("30", "Left", "2", "<B>", "Za dan: " & frmControlMain.DATDO.Value)

Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")

Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Ident<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Dob.id.<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='45%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Naziv<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Kolièina<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Cena<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>Vrednost<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Call crta_izp
If rs.State = 1 Then rs.Close
rs.Open "select * from mada where tip_art='MAT' order by madagrup,dobavit_id ", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb, koli, popu, skgr As Double
skupii = 0
ddva = 0
ddvb = 0
skgr = 0
koli = 0
popu = 0
'do while pozicije glavni
Dim grupp, gruu, zz, skuu As Integer
grupp = 0
gruu = 0
zz = 0
skuu = 0
skuu = Getnazi("select count(madasifr) as xx from mada where tip_art='MAT'")
Do While Not rs.EOF
zz = zz + 1
If zz < skuu Then
izpisi.ProgressBar.Value = zz / skuu * 100
End If
grupp = rs.Fields("madagrup")
If grupp <> gruu Then
If skgr <> 0 Then
Call crta_izp
'Call dono("100", "Right", "1", "<B>", FormatNumber(skgr, 4))
Call dono2("30", "Right", "2", "", "Skupaj grupa: " & FormatNumber(skgr, 4), "70")

skgr = 0
Call crta_izp
End If
Call dono("30", "Left", "2", "<B>", Getnazi("select grupa from grupa where sifra=" & grupp))
Call crta_izp
End If
gruu = rs.Fields("madagrup")
Dim b1, b2, b3 As Double
b1 = 0
b2 = 0
b3 = 0
If Getnazi("select sum(prosta) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra") <> "" Then
'b1 = Getnazi("select sum(format(prosta,'fixed')) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra")
'(IIf([veza_td]=[tip_dok];[kol]*[cena];[prosta]*[cena])
b1 = Getnazi("select sum(format(IIf([veza_td]=[tip_dok],kol,prosta),'fixed')) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra")
End If
If Getnazi("select sum(prosta*cena/prosta) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra") <> "" Then
b2 = Getnazi("select sum(format(IIf([veza_td]=[tip_dok],kol,prosta),'fixed')*format(cena,'#####.###')/format(prosta,'fixed')) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra")
End If
If Getnazi("select sum(prosta*cena) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra") <> "" Then
b3 = Getnazi("select sum(format(IIf([veza_td]=[tip_dok],kol,prosta),'fixed')*format(cena,'#####.###')) as xx from zaloga where   sifra='" & rs.Fields("madasifr") & "' group by sifra")
'b3 = b2 * b1
End If
If b1 <> 0 Then
'b2 = b3 / b1
End If
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("madasifr") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("dobavit_id"), 10) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='45%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("madanazi"), 45) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(b1, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='4%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("madaenme") & "<BR></FONT></B></TD>"
If b3 <> 0 Then
If b1 <> 0 Then
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(b3 / b1, 4) & "<BR></FONT></B></TD>"
Else
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(0, 4) & "<BR></FONT></B></TD>"
End If
Else
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(0, 4) & "<BR></FONT></B></TD>"

End If
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(b3, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Dim bb As Double
bb = 0
If Getnazi("select sum(format(prosta,'fixed')*cena) as xx from zaloga where  sifra='" & rs.Fields("madasifr") & "' group by sifra") <> "" Then
bb = Getnazi("select sum(format(IIf([veza_td]=[tip_dok],kol,prosta),'fixed')*format(cena,'#####.###')) as xx from zaloga where sifra='" & rs.Fields("madasifr") & "' group by sifra")
'bb = b3
Else
'MsgBox (RS.Fields("madasifr"))
End If
skupii = skupii + bb
skgr = skgr + bb



rs.MoveNext
Loop
skupii = Getnazi("select Sum(IIf([veza_td]=[tip_dok],kol,prosta)*[cena]) AS aa from zaloga")
End If
If skgr <> 0 Then
Call crta_izp
'Call dono("100", "Right", "1", "<B>", FormatNumber(skgr, 4))
Call dono2("30", "Right", "2", "", "Skupaj grupa: " & FormatNumber(skgr, 4), "70")

skgr = 0
'Call crta_izp
End If
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='20%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' Align=Left VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'><BR>" & "Vrednost zaloge:" & "</FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='10%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(skupii, 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

End Sub
Sub Print_dob_les(strReportTitle As String)

'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv='" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
 
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
Call glava_izp
Call crta_izp
Call dopr
Call dono("30", "Left", "2", "<B>", "Prejemnik")

Call dono2("30", "Left", "2", "", Getnazi("select dod0" & kib), "4")
Call dono2("30", "Left", "2", "", Getnazi("select ulica from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")
Call dono2("30", "Left", "2", "", Getnazi("select mesto from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")
Call dono2("30", "Left", "2", "", Getnazi("select davcna from partner where naziv='" & LTrim(Getnazi("select dod0" & kib)) & "'"), "4")

Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "3", "<B>", "DOBAVNICA Št.: " & tip_dok & xid_dok)
Call dono("30", "Left", "2", "<B>", "")
Call dono2("30", "Left", "2", "", "DATUM: " & Getnazi("select datum" & kic), "68")
'Call dono2("30", "Left", "2", "", "Št. Dobavnice: " & Getnazi("select dod2" & kib), "60")
Call dono("30", "Left", "2", "", Getnazi("select opis" & kib))
Call dono("30", "Left", "2", "", "")
Call crta_izp
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='40%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='5%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select pop" & kia) & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select znes" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
Call crta_izp
If rs.State = 1 Then rs.Close
rs.Open "select * " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb, koli, popu, pcc As Double
skupii = 0
pcc = 0
ddva = 0
ddvb = 0
koli = 0
popu = 0
'do while pozicije glavni
Do While Not rs.EOF
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("sifra") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select madaean from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select dobavit_id from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='40%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Left(rs.Fields("naziv"), 40) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(rs.Fields("kol"), 2) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='5%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select madaenme from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='15%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber((RS.Fields("cena") * (1 - (RS.Fields("pop") / 100))), 4) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

koli = koli + rs.Fields("kol")
popu = popu + ((rs.Fields("cena") - (rs.Fields("cena") * (1 - (rs.Fields("pop") / 100)))) * rs.Fields("kol"))
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") <> "" Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'></TD>"
msearchResult.WriteLine "<TD WIDTH='57%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
ff = ff + 1

skupii = skupii + ((rs.Fields("cena") * (1 - (rs.Fields("pop") / 100))) * rs.Fields("kol"))
pcc = pcc + (rs.Fields("cena") * rs.Fields("kol"))



rs.MoveNext
Loop


End If
Call crta_izp
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Center", "2", "<B>", "PREDAL:___________________                                   PREJEL:_____________________")
Call dono("30", "Left", "2", "<B>", "")
Call dono("30", "Center", "2", "<B>", "DATUM :___________________                                   DATUM :_____________________")
Call dono("30", "Left", "2", "<B>", "")
msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"
frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume

End Sub

Sub Print_knez(strReportTitle As String)
'msfFlexi.Redraw = False


'intFontSize = msfFlexi.Font.Size - 7
'strFontFace = msfFlexi.Font.Name
intInv = 1
Call fchkFolderPath(App.path & "\Print", True)
'On Error GoTo err:

kia = " from izpisi where tip_dok='" & tip_dok & "' and naziv='" & strReportTitle & "'"
kib = " from glavna where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
kic = " from nabasif where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"

Set msearchResult = fso.CreateTextFile(App.path & "\Print\Report" & str(intInv) & ".htm", True)
msearchResult.WriteLine "<!-- saved from url=(0022)http://internet.e-mail -->"
msearchResult.WriteLine "<html>"
msearchResult.WriteLine "<HEAD>"
msearchResult.WriteLine "<title>" & strReportTitle & "</title>"
msearchResult.WriteLine "</HEAD>"
msearchResult.WriteLine "<BODY BGCOLOR=ffffff>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR><TD>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><IMG SRC='" & App.path & "\gaber.jpg'></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'naziv firme
'do while
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava1" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava2" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select glava3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><FONT SIZE=1 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"

msearchResult.WriteLine "</TABLE>"
'konec naziva

'kupec id
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='17%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idk" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='12%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select sifra from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select idp" & kia) & "<BR></FONT></B></TD>"
'msearchResult.WriteLine "<TD WIDTH='5%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'kupec IDST

msearchResult.WriteLine "</TABLE><TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><B><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select naziv from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select ulica from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select posta from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "   " & Getnazi("select mesto from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "<TR>"
'msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'><BR></FONT>&nbsp</FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Middle><FONT SIZE=1><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select davcna from partner where naziv='" & Getnazi("select dod0" & kib) & "'") & "<BR></FONT>&nbsp</FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='18%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'desna stran
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='35%'></TD>"
msearchResult.WriteLine "<TD WIDTH='25%'></TD>"
'msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dod0" & kib) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='5%'></TD>"

msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'datum + line
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dat" & kia) & "  " & Getnazi("select datum" & kic) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'nazil listine + izdelal
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"

msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='2%'></TD>"
msearchResult.WriteLine "<TD WIDTH='35%' VAlign=Middle><B><FONT SIZE=3 FACE='Helvetica'><BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='17%'></TD>"
msearchResult.WriteLine "<TD WIDTH='15%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select prod" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='30%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select dod5" & kib) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH=2%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & Getnazi("select naz_do" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH=1%></TD>"
msearchResult.WriteLine "<TD WIDTH=19% VAlign=Middle><B><FONT SIZE=3 FACE=Helvetica>" & tip_dok & xid_dok & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH=17%></TD>"
msearchResult.WriteLine "<TD WIDTH=15% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select dobav" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH=0%></TD>"
msearchResult.WriteLine "<TD WIDTH=43% VAlign=Middle><FONT SIZE=2 FACE=Helvetica>" & Getnazi("select dod6" & kib) & "</FONT></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH=57%></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE></TD>"



'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
'opisi poz
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Top><B><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select zap" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='7%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ident" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='24%' VAlign=Top><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select opis" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select kol" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select me" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select cena" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select pop" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='4%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select ddv" & kia) & "<BR></FONT></B></TD>"

msearchResult.WriteLine "<TD WIDTH='7%'></TD>"
msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & Getnazi("select znes" & kia) & "<BR></FONT></B></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'èrta
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
If rs.State = 1 Then rs.Close
rs.Open "select * " & kic, myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
Dim ff As Integer
ff = 1
rs.MoveFirst
Dim skupii, ddva, ddvb As Double
skupii = 0
ddva = 0
ddvb = 0
'do while pozicije glavni
Do While Not rs.EOF
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & str(ff) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & rs.Fields("naziv") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("kol"), 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select madaenme from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='12%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("cena"), 4) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=1>&nbsp<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "<TD WIDTH='5%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(rs.Fields("znes"), 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'do while pozicije dokm
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") <> "" Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'></TD>"
msearchResult.WriteLine "<TD WIDTH='57%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='" & rs.Fields("pozicija") & "'") & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
ff = ff + 1
If Val(Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'")) = 20 Then
ddva = ddva + rs.Fields("znes")
End If
If Val(Getnazi("select madapd from mada where madasifr='" & rs.Fields("sifra") & "'")) = 8.5 Then
ddvb = ddvb + rs.Fields("znes")
End If
skupii = skupii + rs.Fields("znes")
rs.MoveNext
Loop
ddva = (ddva * 1.2) - ddva
ddvb = (ddvb * 1.085) - ddvb
End If
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"


msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup1" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(skupii, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

If ddva <> 0 Then
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup2" & kia) & " 20 %" & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddva, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If
If ddvb <> 0 Then

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup2" & kia) & " 8.5 %" & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddvb, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If


msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='5%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"

msearchResult.WriteLine "<TD WIDTH='32%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='8%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "<TD WIDTH='6%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"

msearchResult.WriteLine "<TD WIDTH='25%' Align=Right VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select skup3" & kia) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='7%'></TD>"

msearchResult.WriteLine "<TD WIDTH='14%' Align=Right VAlign=Middle><B><FONT SIZE=2 FACE='Helvetica'>" & FormatNumber(ddva + ddvb + skupii, 2) & "<BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='1%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>  <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'><BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

'dobavnice
If Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='opis'") <> "" Then

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='13%'>Veza/opis: </TD>"
msearchResult.WriteLine "<TD WIDTH='87%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'>" & Getnazi("select tekst from dokm where tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "' and atribut='opis'") & "<BR></FONT></TD>"
'msearchResult.WriteLine "<TD WIDTH='30%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"
End If



msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "<TABLE WIDTH='100%' BORDER=0 CELLPADDING=0>"
msearchResult.WriteLine "<TR>"
msearchResult.WriteLine "<TD WIDTH='61%'></TD>"
msearchResult.WriteLine "<TD WIDTH='39%' VAlign=Middle><FONT SIZE=2 FACE='Helvetica'> DIREKTOR: Mitja Lešnik <BR></FONT></TD>"
msearchResult.WriteLine "<TD WIDTH='0%'></TD>"
msearchResult.WriteLine "</TR>"
msearchResult.WriteLine "</TABLE>"

msearchResult.WriteLine "</body>"
msearchResult.WriteLine "</html>"

frmPrint.WebBrowser1.Navigate (App.path & "\Print\Report " & intInv & ".htm")
Set msearchResult = Nothing
frmPrint.Show
' msfFlexi.Redraw = True
Exit Sub

err:
intInv = intInv + 1
Resume
End Sub

