Attribute VB_Name = "modDBMain"
Option Explicit
'************************************************************
'API
'************************************************************

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

'************************************************************
'Constants
'************************************************************

Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
 



 '************************************************************
'Variables
'************************************************************

Private hControl As Long
Private lPrevWndProc As Long

'*************************************************************
'WindowProc
'*************************************************************

'zDelta: The value of the high-order word of wParam.
'Indicates the distance that the wheel is rotated, expressed in multiples or
'divisions of WHEEL_DELTA, which is 120. A positive value indicates that the
'wheel was rotated forward, away from the user; a negative value indicates
'that the wheel was rotated backward, toward the user.
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim fwKeys As Long
    Dim zDelta As Long
    Dim Xpos As Long
    Dim Ypos As Long

    'Test if the message is WM_MOUSEWHEEL
    If Lmsg = WM_MOUSEWHEEL Then
        fwKeys = wParam And 65535
        zDelta = wParam / 65536
        Xpos = lParam And 65535
        Ypos = lParam / 65536
        'Call the Form1's Procedure to handle the MouseWheel event
       ' MouseWheel fwKeys, zDelta, Xpos, Ypos
    End If
    'Sends message to previous procedure
    'This is VERY IMPORTANT!!!
    WindowProc = CallWindowProc(lPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function

'*************************************************************
'Hook
'*************************************************************
Public Sub Hook(ByVal hControl_ As Long)
    hControl = hControl_
lPrevWndProc = SetWindowLong(hControl, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'*************************************************************
'UnHook
'*************************************************************
Public Sub UnHook()
    Call SetWindowLong(hControl, GWL_WNDPROC, lPrevWndProc)
End Sub



Public Function ConnectDB(ByRef vDB As ADODB.Connection, PathFileName As String) As Boolean

On Error GoTo errh
 
    If vDB.State = adStateOpen Then vDB.Close
        
    vDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="
    
    ConnectDB = True
    
    Exit Function
    
errh:

    'WriteErrorLog "modDBMain", "ConnectDB", err.Description
    ConnectDB = False
    
End Function

Public Function CloseDB(ByRef vDB As ADODB.Connection)
    On Error GoTo errh
    vDB.Close
errh:
End Function


Public Function ConnectRS(ByRef vDB As ADODB.Connection, ByRef vRS As ADODB.Recordset, sSQL As String, Optional sHowMSG As Boolean = True, Optional ByRef iErrNumber As Variant, Optional ByRef sErrDescription As Variant) As Boolean
    
On Error GoTo errh

    
    Set vRS = Nothing
    Set vRS = New ADODB.Recordset
  
  
    vRS.Open sSQL, vDB, adOpenStatic, adLockOptimistic
    ConnectRS = True

    
    Exit Function
    
'-------------------------------------------
errh:
    If sHowMSG = True Then
       ' WriteErrorLog "modDBMain", "ConnectRS", "Unable to connect Recordset / Err: " & err.Description
    End If
    If Not IsMissing(iErrNumber) Then
        iErrNumber = err.Number
    End If
    If Not IsMissing(sErrDescription) Then
        sErrDescription = err.Description
    End If
    ConnectRS = False
End Function


Public Function RecordNoMatch(ByRef vRS As ADODB.Recordset) As Boolean
On Error GoTo errh:

    RecordNoMatch = (vRS.BOF = True Or vRS.EOF = True)

    Exit Function
    
errh:
    RecordNoMatch = False
    
End Function


Public Function AnyRecordExisted(ByRef vRS As ADODB.Recordset) As Boolean
    If vRS.State = adStateClosed Then
        AnyRecordExisted = False
        Exit Function
    End If
    
    
    vRS.Requery
    
    If (vRS.BOF = True) And (vRS.EOF = True) Then
        AnyRecordExisted = False
    Else
        On Error GoTo errh
        vRS.MoveFirst
        AnyRecordExisted = True
    End If

    Exit Function
    '--------------------------
    
errh:
    AnyRecordExisted = False
End Function


Public Function ReadField(ByRef vField As Field) As Variant
    
    On Error GoTo errh

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
    
errh:
    ReadField = ""
End Function

Public Function getRecordCount(ByRef vRS As ADODB.Recordset) As Long
    If AnyRecordExisted(vRS) Then
        vRS.Requery
        vRS.MoveLast
        getRecordCount = vRS.RecordCount
    Else
        getRecordCount = 0
    End If
End Function

Public Function RSMoveFirst(ByRef vRS As ADODB.Recordset) As Boolean
    If AnyRecordExisted(vRS) Then
        vRS.MoveFirst
        RSMoveFirst = True
    Else
        RSMoveFirst = False
    End If
End Function




