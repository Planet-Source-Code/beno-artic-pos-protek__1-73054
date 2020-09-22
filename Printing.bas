Attribute VB_Name = "modPrinting"
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long
    cpMax As Long
End Type

Private Type FormatRange
    hdc As Long
    hdcTarget As Long
    rc As Rect
    rcPage As Rect
    chrg As CharRange
End Type
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" ( _
    ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

Private Declare Function EnumChildWindows Lib "user32.dll" ( _
    ByVal hWndParent As Long, _
    ByVal lpEnumFunc As Long, _
    ByVal lParam As Long _
) As Long

Private hWndIE As Long
Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, ByVal nIndex As Long) As Long


Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
    lp As Any) As Long

Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
    (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
    ByVal lpOutput As Long, ByVal lpInitData As Long) As Long



Public Sub PrintPreview(RTF As WebBrowser, LeftMarginWidth As Currency, _
    TopMarginHeight As Currency, RightMarginWidth As Currency, BottomMarginHeight As Currency, _
    pgOrientation As Integer)
      
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    Dim iCount As Integer

    On Error GoTo ErrHandle
    
'Set the orientation of the printer
    Printer.Orientation = pgOrientation
    Printer.ScaleMode = vbTwips

' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = CLng(LeftMarginWidth - LeftOffset)
    TopMargin = CLng(TopMarginHeight - TopOffset)
    RightMargin = CLng((Printer.Width - RightMarginWidth) - LeftOffset)
    BottomMargin = CLng((Printer.Height - BottomMarginHeight) - TopOffset)

' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin


    frmPreview.SizePreview Printer.Width, Printer.Height

    fr.hdc = frmPreview.picPreview(0).hdc
    fr.hdcTarget = frmPreview.picPreview(0).hdc
    fr.rc = rcDrawTo
    fr.rcPage = rcPage
    fr.chrg.cpMin = 0
    fr.chrg.cpMax = -1


    TextLength = Len(RTF.Document.Body.innertext)

    Dim iPage As Integer
    
    iPage = 1
    
    Do
        With frmPreview
            If iPage > 1 Then
                .AddPage iPage
                fr.hdc = .picPreview(iPage - 1).hdc
                fr.hdcTarget = .picPreview(iPage - 1).hdc
            End If
            .picPreview(iPage - 1).Print
        End With

        NextCharPosition = SendMessage(webhw, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        
        iPage = iPage + 1
    Loop

    r = SendMessage(webhw, EM_FORMATRANGE, False, ByVal CLng(0))

    frmPreview.Show vbModal

    Exit Sub
    
ErrHandle:
    Select Case err.Number
        Case 482
            MsgBox "OptiType couldn't find an installed printer.  Install a Windows-compatible printer and try again.", vbCritical, "No Windows-printer found."
            Exit Sub
        Case Else
            MsgBox err.Number & " " & err.Description
            Resume Next
    End Select
    
End Sub


Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim sClassName As String
    sClassName = String(255, vbNullChar)
    Call GetClassName(hWnd, sClassName, 255)
    sClassName = Left$(sClassName, InStr(sClassName, vbNullChar) - 1)
    If sClassName <> "Internet Explorer_Server" Then
        EnumChildProc = 1
    Else
        hWndIE = hWnd
    End If
End Function

Public Function GetBrowserHandle(ByVal hWndParent) As Long
    hWndIE = 0
    Call EnumChildWindows(hWndParent, AddressOf EnumChildProc, 1)
    GetBrowserHandle = hWndIE
End Function

