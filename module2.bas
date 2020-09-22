Attribute VB_Name = "module2"

Option Explicit
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hwnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hwnd As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

' string used to identify which record to open in adoClsCustomer
Public idnum As Integer
' string used to identify which record to open in adoClsPurchases
Public ponum As Integer


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
''''''''''''' Scroll stuff

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long
  Dim fFrm As Form

  Select Case Lmsg
  
    Case WM_MOUSEWHEEL
    
      MouseKeys = wParam And 65535
      Rotation = wParam / 65536
      Xpos = lParam And 65535
      Ypos = lParam / 65536
      
      Set fFrm = GetForm(Lwnd)
      If fFrm Is Nothing Then
        ' it's not a form
        If Not IsOver(Lwnd, Xpos, Ypos) And IsOver(GetParent(Lwnd), Xpos, Ypos) Then
          ' it's not over the control and is over the form,
          ' so fire mousewheel on form (if it's not a dropped down combo)
          If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
            GetForm(GetParent(Lwnd)).MouseWheel MouseKeys, Rotation, Xpos, Ypos
            Exit Function ' Discard scroll message to control
          End If
        End If
      Else
        ' it's a form so fire mousewheel
        If IsOver(fFrm.hwnd, Xpos, Ypos) Then fFrm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
      End If
  End Select
  
  WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub WheelHook(ByVal hwnd As Long)
  On Error Resume Next
  SetProp hwnd, "PrevWndProc", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hwnd As Long)
  On Error Resume Next
  SetWindowLong hwnd, GWL_WNDPROC, GetProp(hwnd, "PrevWndProc")
  RemoveProp hwnd, "PrevWndProc"
End Sub

Public Sub FlexGridScroll(ByRef FG As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim NewValue As Long
  Dim Lstep As Single

  On Error Resume Next
  With FG
    Lstep = .Height / .RowHeight(0)
    Lstep = Int(Lstep)
    If .Rows < Lstep Then Exit Sub
    Do While Not (.RowIsVisible(.TopRow + Lstep))
      Lstep = Lstep - 1
    Loop
    If Rotation > 0 Then
        NewValue = .TopRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .TopRow + Lstep
        If NewValue > .Rows - 1 Then
            NewValue = .Rows - 1
        End If
    End If
    .TopRow = NewValue
  End With
End Sub

Public Sub HorFlexGridScroll(ByRef HFG As MSHFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim NewValue As Long
  Dim Lstep As Single

  On Error Resume Next
  With HFG
  
    Lstep = .Height / .RowHeight(0)
    Lstep = Int(Lstep)
    
    If .Rows < Lstep Then Exit Sub
    Do While Not (.RowIsVisible(.TopRow + Lstep))
      Lstep = Lstep - 1
    Loop
    If Rotation > 0 Then
        NewValue = .TopRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .TopRow + Lstep
        If NewValue > .Rows - 1 Then
            NewValue = .Rows - 1
        End If
    End If
    .TopRow = NewValue
  End With
End Sub

Public Sub DataGridScroll(ByRef dGrid As DataGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  ' I wrote this code. The other parts of the scroll were written by somebody else who did not put there name in their code.
  ' J. Fisher ladenv@yahoo.com 10/8/06
  Dim NewValue As Long
  Dim Lstep As Single
  Dim CurrentLoc As Variant
  On Error Resume Next
  With dGrid
   
    Lstep = .Height / .RowHeight
    Lstep = Int(Lstep)

    If .ApproxCount < Lstep Then Exit Sub
    CurrentLoc = (.FirstRow + Lstep)
    Do While Not .FirstRow <> CurrentLoc
      Lstep = Lstep - 1
      CurrentLoc = (.FirstRow + Lstep)
    Loop
   If Rotation > 0 Then
        NewValue = .FirstRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .FirstRow + Lstep
        If NewValue > .ApproxCount - 1 Then
            NewValue = .ApproxCount - 1
        End If
    End If
    .FirstRow = NewValue
  End With

End Sub

Public Function IsOver(ByVal hwnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hwnd, rectCtl
  With rectCtl
    If lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom Then IsOver = True
  End With
End Function

Private Function GetForm(ByVal hwnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hwnd = hwnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

'''''' End of scroll


''' Connect database to adodc control
Public Sub Connect_AdoControl(ByRef AdoContObj As Adodc, ByVal DatabaseLocation As String, ByVal GetRecordString As String, ByVal DataBaseHavePass As Boolean, ByVal DBPassword As String)

If DataBaseHavePass = True Then
   AdoContObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DatabaseLocation & "; Persist Security Info= False; Jet OLEDB: Database Password=" & DBPassword
Else
   AdoContObj.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DatabaseLocation & "; Persist Security Info= False"
End If

AdoContObj.CommandType = adCmdText
AdoContObj.RecordSource = GetRecordString
AdoContObj.Refresh

End Sub
''''''''''''''''''''''''''''''''''''''''''


Public Function openword()
'Static objword As Object
'  Set objword = CreateObject("word.application")
  
  Call Shell(App.path & "\repor.doc", vbMaximizedFocus)
  
'  If err.Number <> 0 Then
 '               Set objword = New WORD.Application
            
 '                   err.clear
 '           End If
    
    
  '  On Error Resume Next
  '  objword.Visible = False
End Function
Public Function FlexGrd_SaveToExcel(FG As MSHFlexGrid, Optional sHeader As String = "", Optional sFooter As String = "", Optional ColumnHeaderFontColorIndex As Long, Optional ColumnHeaderBackColorIndex As Long, Optional CoLogoPicLocation As String, Optional WorkBkBackColorIndex As Long, Optional WorkBkGridColorIndex As Long, Optional AlternateRowColorIndex1 As Long, Optional AlternateRowColorIndex2 As Long, Optional AutoColumnFitter As Boolean, Optional AutoFitLogoPic As Boolean)
 ' I wrote this and you are to free to use it in any application
 ' It should work in any excel object library. The new office 2007 I tested it on. The only problem I have is inserting the picture in
 ' right place. You have to reference the excel library or the office library in the beta 2207. J. Fisher ladenv@yahoo.com 10/8/06
 ' 10/12/06 update
 '  Changed to Autofit columns
 '  Alternating row colors in excel
 '
  Static objExcelDel As Object
  Static objWorkbookDel As Excel.Workbook
  Static objWorksheetDel As Excel.Worksheet
  Static HeadRange    As Excel.Range
  Static NewRange As Excel.Range
  Static GridRange As Range
  Static PicObject As Excel.ShapeRange
  Dim lRow As Integer, lCol As Integer
  Dim I As Integer, J As Integer
  Dim c As Integer

  Dim rowOffset As Long
  Dim TempStr() As String

  
  Set objExcelDel = CreateObject("Excel.application")
  
  
  
  If err.Number <> 0 Then
                Set objExcelDel = New Excel.Application
            
                    err.clear
            End If
        On Error Resume Next
            objExcelDel.Visible = False
  
  If Len(sHeader) > 0 Then
    TempStr = Split(sHeader, vbTab)
    rowOffset = UBound(TempStr) + 1
  End If
  
  
  
  Set objWorkbookDel = objExcelDel.Workbooks.Add
        
        'Turn off the alerts
        objExcelDel.DisplayAlerts = False
            
        'Set objWorksheet to the remaining worksheet.
        Set objWorksheetDel = objExcelDel.ActiveSheet
 
  With objWorksheetDel
       
    ' Sheet Header
   
    For lRow = 1 To rowOffset
           .PageSetup.CenterHeader = TempStr(lRow - 1)
    Next lRow

    '************************
    ' Get Column Headers
    For lRow = 1 To FG.FixedRows
      For lCol = 0 To FG.Cols
        .Cells(4, lCol - 1) = FG.TextMatrix(lRow - 1, lCol - 1)
      Next lCol
    Next lRow
   ''''''''
   'I have never seen anybody do this before.
   'Sets Background color of worksheets in workbook
   If Val(WorkBkBackColorIndex) > 0 Then
   objWorkbookDel.Styles("Normal").Interior.ColorIndex = WorkBkBackColorIndex
   End If
    'Gridlines will not be visible but you can add that to by
   If Val(WorkBkGridColorIndex) > 0 Then
    With objWorkbookDel.Styles("Normal").Borders(xlLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1  ' 1 is black
    End With
    With objWorkbookDel.Styles("Normal").Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With objWorkbookDel.Styles("Normal").Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    With objWorkbookDel.Styles("Normal").Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
    End If
   '''''''''
  
    
   
    Set HeadRange = objWorksheetDel.Range(objWorksheetDel.Cells(4, 1), _
                objWorksheetDel.Cells(4, lCol - 2))
    With HeadRange
        '*****Sets Column Header Back Color
        If Val(ColumnHeaderBackColorIndex) > 0 Then
            .Interior.ColorIndex = ColumnHeaderBackColorIndex
            Else
            ' My Default Background color for Column header index change it to what ever you want
            .Interior.ColorIndex = 5
            End If
         '************************************
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = 6
        .Interior.Pattern = xlLightHorizontal
        .Interior.ColorIndex = 20
        .Font.Name = "Rockwell"
        .Font.FontStyle = "Bold"
        .Font.Shadow = True
        '***** Sets Column header Font color*****
        If Val(ColumnHeaderFontColorIndex) > 0 Then
            .Font.ColorIndex = ColumnHeaderFontColorIndex
            Else
            ' My Default Font color for Column header index change it to what ever you want
            .Font.ColorIndex = 2
            End If
        .Font.Bold = True
        '************************************
        'Sets border colors of header. You could also add this
        'to the function but I thought I was getting carried away
        'as it was.
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 16  'grey
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 16
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 16
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 16
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 1 ' Black
        End With
    End With
    
    HeadRange = Nothing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim RowCounter As Integer ' used for all alternate row color
    RowCounter = 0    ' ditto
   ' Dim ColCounter As Integer ' used for all alternate row color
   ' ColCounter = 0
    Dim G As Integer ' ditto
    Dim Alternate As Boolean  'ditto
    '''''''''''''''''''''''''''''''''''''''
    ' Fill excel sheet with data
    ' Row data from flexgrid
    For I = 1 To FG.Rows
       
        For J = 0 To FG.Cols
            objWorksheetDel.Cells(I + 4, J) = FG.TextMatrix(I, J)
            objWorksheetDel.Cells(I + 4, J + 1).VerticalAlignment = xlTop
        Next J
        RowCounter = RowCounter + 1
    Next I
    RowCounter = RowCounter - 1  ' Getting rid of extra row
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Alternate row colors on Excel spreadsheet
    If AlternateRowColorIndex1 <> "" And AlternateRowColorIndex2 <> "" Then
   
    G = 0
    Do Until G = RowCounter ' RowCounter is figured when row data is taken
        Set NewRange = objWorksheetDel.Range(objWorksheetDel.Cells(G + 5, 1), _
            objWorksheetDel.Cells(G + 5, lCol - 2))
  
        With NewRange
        If Alternate <> True Then
            .Interior.ColorIndex = AlternateRowColorIndex1
            .Borders.ColorIndex = 31
            'Sets font color either 1 Black or 2 white for row
            Select Case AlternateRowColorIndex1
                Case 1, 3, 5, 9, 11, 13, 14, 16, 17, 21, 23, 25
                    .Font.ColorIndex = 2
                Case Else
                    .Font.ColorIndex = 1
            End Select
            Alternate = True
           Else
            .Interior.ColorIndex = AlternateRowColorIndex2
            .Borders.ColorIndex = 31
            'Sets font color either 1 Black or 2 white
            Select Case AlternateRowColorIndex2
                Case 1, 3, 5, 9, 11, 13, 14, 16, 17, 21, 23, 25
                    .Font.ColorIndex = 2
                Case Else
                    .Font.ColorIndex = 1
            End Select
            Alternate = False
            End If
        End With
        NewRange = Nothing
         G = G + 1
    Loop
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Autofit columns
    If AutoColumnFitter = True Then
        .Columns.AutoFit
        End If
        '******************************************
   'Inserts company logo or picture in Cell A1
   ' The logo has to be the right size for the effect you are looking for. I suggest taking your current logo and editing it in photoshop
   ' In Office 2007 you will have to specify the exact cell you want but for previous versions you do not have to.
   If Len(CoLogoPicLocation) > 0 Then
          Set PicObject = objWorksheetDel.Pictures.Insert(CoLogoPicLocation)
           ' PicObject.Pictures.Insert (CoLogoPicLocation)
            End If
   '******************************************
    '''''''''''''''''''''''''''''''''''''''''
    ' Fit Clogo picture to col headers does not work yet
   ' If AutoFitLogoPic = True Then
   ' Dim ColCount As Integer
   ' ColCount = FG.Cols - 1
   ' Dim cc As Integer
   ' Dim PicWidth As Double
   ' Dim msoScaleFromTopLeft As Variant
   ' msoScaleFromTopLeft = 10
   ' Do Until cc = ColCount
   '     PicWidth = .Columns(1, cc).ColumnWidth
   '     cc = cc + 1
   '     Loop
   ' PicObject.LockAspectRatio = msoFalse
   ' PicObject.Width = PicWidth
   ' PicObject.ScaleWidth PicWidth, msoFalse, msoScaleFromTopLeft
   '
   ' End If
  
   ' PicObject.Shapes.Range.Width = PicWidth
    
    
    ''''''''''''''''''''''''''''''''''''''''''
    objWorksheetDel.OLEObjects
    
    
    ' Page Footer
    If Len(sFooter) > 0 Then
      TempStr = Split(sFooter, vbTab)
      For lRow = 0 To UBound(TempStr)
          .PageSetup.CenterFooter = TempStr(lRow)
      Next lRow
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  End With
  objExcelDel.Visible = True
                       objExcelDel.DisplayAlerts = True
                       Set objWorksheetDel = Nothing
                       Set objWorkbookDel = Nothing
                       Set objExcelDel = Nothing
End Function


Public Function MSGridSort(TheGrid As MSFlexGrid)
' GENERIC SORTING WHICH YOU CAN ALSO SORT BY DATE
'J. Fisher ladenv@yahoo.com 10/8/06
 Dim sortCo As Long

 Static PrevSortCo As Integer
 Static aSortAsc As Boolean
 sortCo = TheGrid.Col
            If sortCo = PrevSortCo Then
                If aSortAsc Then
                     TheGrid.Sort = 2  ' SORT DESCENDING
                     aSortAsc = False
                Else
                    TheGrid.Sort = 1    'SORT ASCENDING
                     aSortAsc = True
                End If
            Else
                    TheGrid.Sort = 1   ' SORT ASCENDING
                    aSortAsc = True
            End If
                PrevSortCo = TheGrid.ColSel
 
End Function






Public Sub AutosizeGridColumns(ByRef msFG As MSHFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer)
Dim I, J As Integer
Dim txtString As String
Dim intTempWidth, BiggestWidth As Integer
Dim intRows As Integer
Const intPadding = 150
With msFG
 For I = 0 To .Cols - 1
' Loops through every column
.Col = I
' Set the active colunm
intRows = .Rows
' Set the number of rows
If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
' If there are more rows of data, reset
' intRows to the MaxRowsToParse constant
 
BiggestWidth = 0
' Reset some values to 0
For J = 0 To intRows - 1
 ' check up to MaxRowsToParse # of rows and obtain
 ' the greatest width of the cell contents
 
 .Row = J
 
 txtString = .Text
 intTempWidth = (Len(txtString) * 150) + intPadding
 ' The intPadding constant compensates for text insets
 ' You can adjust this value above as desired.
 
 If intTempWidth > BiggestWidth Then BiggestWidth = intTempWidth
 ' Reset intBiggestWidth to the intMaxColWidth value if necessary
Next J
.ColWidth(I) = BiggestWidth
 Next I
 ' Now check to see if the columns aren't as wide as the grid itself.
 ' If not, determine the difference and expand each column proportionately
 ' to fill the grid
 intTempWidth = 0
 
 For I = 0 To .Cols - 1
intTempWidth = intTempWidth + .ColWidth(I)
' Add up the width of all the columns
 Next I
 
 If intTempWidth < msFG.Width Then
' Compate the width of the columns to the width of the grid control
' and if necessary expand the columns.
intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
' Determine the amount od width expansion needed by each column
For I = 0 To .Cols - 1
 .ColWidth(I) = .ColWidth(I) + intTempWidth
 ' add the necessary width to each column
 
Next I
 End If
End With
End Sub
