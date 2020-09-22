Attribute VB_Name = "Module1"

Option Explicit
'-----------------------------MODULE CONSTANTS & VARIABLES------------------------------



    Private Const DESIGN_HORZRES As Long = 1024  '<- CHANGE THIS VALUE TO THE RESOLUTION
                                                'YOU DESIGNED YOUR FORMS IN.
                                                '(e.g. 800 X 600 -> 800)
Private Const DESIGN_VERTRES As Long = 768   '<- CHANGE THIS VALUE TO THE RESOLUTION
                                                'YOU DESIGNED YOUR FORMS IN.
                                                '(e.g. 800 X 600 -> 600)
Private Const DESIGN_PIXELS As Long = 96        '<- CHANGE THIS VALUE TO THE DPI
                                                'SETTING YOU DESIGNED YOUR FORMS IN.
                                                '(If in doubt do not alter the
                                                'DESIGN_PIXELS setting as most
                                                'systems use 96 dpi.)
Private Const WM_HORZRES As Long = 8
Private Const WM_VERTRES As Long = 10
Private Const WM_LOGPIXELSX As Long = 88
Private Const TITLEBAR_PIXELS As Long = 18
Private Const COMMANDBAR_PIXELS As Long = 26
Private Const COMMANDBAR_LEFT As Long = 0
Private Const COMMANDBAR_TOP As Long = 1
Private OrigWindow As tWindow                   'Module level variable holds the
                                                'original window dimensions before


Private Type tRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type tDisplay
    Height As Long
    Width As Long
    DPI As Long
End Type

Private Type tWindow
    Height As Long
    Width As Long
End Type

Private Type tControl
    Name As String
    Height As Long
    Width As Long
    Top As Long
    Left As Long
End Type
'-------------------------- END MODULE CONSTANTS & VARIABLES----------------------------

'------------------------------------API DECLARATIONS-----------------------------------
Private Declare Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function WM_apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" _
() As Long

Private Declare Function WM_apiGetDC Lib "user32" Alias "GetDC" _
(ByVal hwnd As Long) As Long

Private Declare Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" _
(ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function WM_apiGetWindowRect Lib "user32.dll" Alias "GetWindowRect" _
(ByVal hwnd As Long, lpRect As tRect) As Long

Private Declare Function WM_apiMoveWindow Lib "user32.dll" Alias "MoveWindow" _
(ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function WM_apiIsZoomed Lib "user32.dll" Alias "IsZoomed" _
(ByVal hwnd As Long) As Long


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Const WM_CLOSE = &H10
Public Const GW_HWNDPREV = 3

     Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
     
     Declare Function GetWindow Lib "user32" _
       (ByVal hwnd As Long, ByVal wCmd As Long) As Long
     Declare Function SetForegroundWindow Lib "user32" _
       (ByVal hwnd As Long) As Long
Public Sub ActivatePrevInstance()
        Dim OldTitle As String
        Dim PrevHndl As Long
        Dim Result As Long

        'Save the title of the application.
        OldTitle = App.title

        'Rename the title of this application so FindWindow
        'will not find this application instance.
        App.title = "unwanted instance"

        'Attempt to get window handle using VB4 class name.
        PrevHndl = FindWindow("ThunderRTMain", OldTitle)

        'Check for no success.
        If PrevHndl = 0 Then
           'Attempt to get window handle using VB5 class name.
           PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
        End If

        'Check if found
        If PrevHndl = 0 Then
        'Attempt to get window handle using VB6 class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
        End If

        'Check if found
        If PrevHndl = 0 Then
           'No previous instance found.
           Exit Sub
        End If

        'Get handle to previous window.
        PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

        'Restore the program.
        Result = OpenIcon(PrevHndl)

        'Activate the application.
        Result = SetForegroundWindow(PrevHndl)

        'End the application.
        End
     End Sub

Public Sub KillAppPrev(app_name As String)
Dim hwnd As Long
hwnd = FindWindow(vbNullString, app_name)

'MsgBox hwnd
If hwnd <> 0 Then
PostMessage hwnd, WM_CLOSE, 0, 0
End If

End Sub

'--------------------------------- END API DECLARATIONS----------------------------------

'---------------------------------------------------------------------------------------
' Procedure : getScreenResolution
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Function returns the current height, width and dpi.
'---------------------------------------------------------------------------------------
Private Function getScreenResolution() As tDisplay

Dim hDCcaps As Long
Dim lngRtn As Long

On Error Resume Next

    'API call get current resolution:-
    hDCcaps = WM_apiGetDC(0) 'Get display context for desktop (hwnd = 0).
    With getScreenResolution
        .Height = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
        .Width = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
        .DPI = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSX)
    End With
    lngRtn = WM_apiReleaseDC(0, hDCcaps) 'Release display context.
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFactor
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Function returns the value that the form's/control's height, width, top &
'             left should be multiplied by to fit the current screen resolution.
'---------------------------------------------------------------------------------------
Public Function getFactor(blnVert As Boolean) As Single

Dim sngFactorP As Single

On Error Resume Next

    If getScreenResolution.DPI <> 0 Then
        sngFactorP = DESIGN_PIXELS / getScreenResolution.DPI
    Else
        sngFactorP = 1 'Error with dpi reported so assume 96 dpi.
    End If
    If blnVert Then 'return vertical resolution.
        getFactor = (getScreenResolution.Height / DESIGN_VERTRES) * sngFactorP
    Else 'return horizontal resolution.
        getFactor = (getScreenResolution.Width / DESIGN_HORZRES) * sngFactorP
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReSizeForm
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Routine should be called on a form's onOpen or onLoad event.
'---------------------------------------------------------------------------------------
Public Sub ReSizeForm(ByVal frm As Form)

Dim rectWindow As tRect
Dim lngWidth As Long
Dim lngHeight As Long
Dim sngVertFactor As Single
Dim sngHorzFactor As Single
Dim sngFontFactor As Single

On Error Resume Next

    sngVertFactor = getFactor(True)  'Local function returns vertical size change.
    sngHorzFactor = getFactor(False)  'Local function returns horizontal size change.
    'Choose lowest factor for resizing fonts:-
    sngFontFactor = VBA.IIf(sngHorzFactor < sngVertFactor, sngHorzFactor, sngVertFactor)
    Resize sngVertFactor, sngHorzFactor, sngFontFactor, frm 'Local procedure to resize form sections & controls.
    If WM_apiIsZoomed(frm.hwnd) = 0 Then 'Don't change window settings for max'd form.
         'Me.maximize 'Maximize the Access Window.
        
        'Store for dimensions in rectWindow:-
        Call WM_apiGetWindowRect(frm.hwnd, rectWindow)
        'Calculate and store form height and width in local variables:-
        With rectWindow
            lngWidth = .Right - .Left
            lngHeight = .Bottom - .Top
        End With
        'Resize the form window as required (don't resize this for sub forms):-
        If frm.Parent.Name = VBA.vbNullString Then
            Call WM_apiMoveWindow(frm.hwnd, ((getScreenResolution.Width - _
            (sngHorzFactor * lngWidth)) / 2) - getLeftOffset, _
            ((getScreenResolution.Height - (sngVertFactor * lngHeight)) / 2) - _
            getTopOffset, lngWidth * sngHorzFactor, lngHeight * sngVertFactor, 1)
        End If
    End If
    Set frm = Nothing 'Free up resources.
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Resize
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Routine re-scales the form sections and controls.
'---------------------------------------------------------------------------------------
Private Sub Resize(sngVertFactor As Single, sngHorzFactor As Single, sngFontFactor As _
Single, ByVal frm As Form)

Dim ctl As Control            'Form control variable.
Dim arrCtls() As tControl            'Array of Tab and Option Group control properties.
Dim lngI As Long                     'Loop counter.
Dim lngJ As Long                     'Loop counter.
Dim lngWidth As Long                 'Stores form's new width.
Dim lngHeaderHeight As Long          'Stores header's new height.
Dim lngDetailHeight As Long          'Stores detail's new height.
Dim lngFooterHeight As Long          'Stores footer's new height.
Dim blnHeaderVisible As Boolean      'True if form header visible before resize.
Dim blnDetailVisible As Boolean      'True if form detail visible before resize.
Dim blnFooterVisible As Boolean      'True if form footer visible before resize.
Const FORM_MAX As Long = 31680       'Maximum possible form width & section height.

On Error Resume Next
    
    With frm
        .Painting = False 'Turn off form painting.
        'Calculate form's new with and section heights and store in local variables
        'for later use:-
        lngWidth = .Width * sngHorzFactor
        lngHeaderHeight = .Section.Height * sngVertFactor
        lngDetailHeight = .Sectio.Height * sngVertFactor
        lngFooterHeight = .Section.Height * sngVertFactor
        'Now maximize the form's width and height while controls are being resized:-
        .Width = FORM_MAX
        .Section.Height = FORM_MAX
        .SectionHeight = FORM_MAX
        .SectionHeight = FORM_MAX
        'Hiding form sections during resize prevents invalid page fault after
        'resizing column widths for list boxes on forms with a header/footer:-
        blnHeaderVisible = .Section.Visible
        blnDetailVisible = .Section.Visible
        blnFooterVisible = .Section.Visible
        .Section.Visible = False
        .Section.Visible = False
        .Section.Visible = False
    End With
    'Resize array to hold 1 element:-
    ReDim arrCtls(0)
    'Gather properties for Tabs and Option Groups to recify height/width problems:-
  
    'Resize and locate each control:-
    For Each ctl In frm.Controls
        'If ctl.ControlType <> acPage Then 'Ignore pages in Tab controls.
            
            With ctl
            If TypeOf ctl Is xcKeypad Then
            Else
                .Height = .Height * sngVertFactor
                .Left = .Left * sngHorzFactor
                .Top = .Top * sngVertFactor
                .Width = .Width * sngHorzFactor
                .FontSize = .FontSize * sngFontFactor
                .Font.Size = .Font.Size * sngFontFactor
                'Enhancement by Myke Myers --------------------------------------->
                'Fix certain Combo Box, List Box and Tab control properties:-
                    
                        .ColumnWidths = adjustColumnWidths(.ColumnWidths, sngHorzFactor)
                        .ListWidth = .ListWidth * sngHorzFactor
                If TypeOf ctl Is TextBox Then
                 .Height = .Height * sngVertFactor * 0.7
               
                End If
               
                If TypeOf ctl Is MSHFlexGrid Then
                .FontFixed.Size = .FontFixed.Size * sngFontFactor
                
                End If
                '------------------------------------> End enhancement by Myke Myers.
            End If
            End With
        'End If
    Next ctl
    '********************************************************
    '* Note if scaling form up: If Tab controls or Option   *
    '* Groups are too near the bottom or right side of the  *
    '* form they WILL distort due to the way that Access    *
    '* keeps the child controls within the control frame.   *
    '* Try moving these controls left or up if possible.    *
    '* The opposite is true for scaling down so in this     *
    '* case try moving these controls right or down.        *
    '********************************************************
    'Now try to rectify Tabs and Option Groups height/widths:-
    For lngJ = 0 To lngI
        With frm.Controls.Item(arrCtls(lngJ).Name)
            .Left = arrCtls(lngJ).Left * sngHorzFactor
            .Top = arrCtls(lngJ).Top * sngVertFactor
            .Height = arrCtls(lngJ).Height * sngVertFactor
            .Width = arrCtls(lngJ).Width * sngHorzFactor
        End With
    Next lngJ
    'Now resize height for each section and form width using stored values:-
    With frm
        .Width = lngWidth
        .Section.Height = lngHeaderHeight
        .Section.Height = lngDetailHeight
        .Section.Height = lngFooterHeight
        'Now unhide form sections:-
        .Section.Visible = blnHeaderVisible
        .Section.Visible = blnDetailVisible
        .Section.Visible = blnFooterVisible
        .Painting = True 'Turn form painting on.
    End With
    Erase arrCtls 'Destory array.
    Set ctl = Nothing 'Free up resources.

End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTopOffset
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Function returns the total size in pixels of menu/toolbars at the top of
'             the Access window allowing the form to be positioned in the centre of the
'             screen.
'---------------------------------------------------------------------------------------
Private Function getTopOffset() As Long

Dim cmdBar As Object
Dim lngI As Long

On Error GoTo err

     
exit_fun:
    Exit Function
    
err:
    'Assume only 1 visible command bar plus the title bar:
    getTopOffset = TITLEBAR_PIXELS + COMMANDBAR_PIXELS
    Resume exit_fun
     
End Function

'---------------------------------------------------------------------------------------
' Procedure : getLeftOffset
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Function returns the total size in pixels of menu/toolbars at the left of
'             the Access window allowing the form to be positioned in the centre of the
'             screen.
'---------------------------------------------------------------------------------------
Private Function getLeftOffset() As Long

Dim cmdBar As Object
Dim lngI As Long

On Error GoTo err

     
exit_fun:
    Exit Function
    
err:
    'Assume no visible command bars:-
    getLeftOffset = 0
    Resume exit_fun
     
End Function
 
'---------------------------------------------------------------------------------------
' Procedure : adjustColumnWidths
' DateTime  : 27/01/2003
' Author    : Myke Myers [Split() replacement for Access 97 by Jamie Czernik]
' Purpose   : Adjusts column widths for list boxes and combo boxes.
' Called By : modResize/Resize().
' Event Modification Information:
'   1. Chris Garland    02/07/2006
'   The event was modified to check if there is any column size entry, and if not, the
'   property is left blank on the control.
'---------------------------------------------------------------------------------------
Private Function adjustColumnWidths(strColumnWidths As String, sngFactor As Single) As String
On Error GoTo Err_adjustColumnWidths

Dim astrColumnWidths() As String                'Array to hold the individual column widths
Dim strTemp As String                           'Holds the recombined columnwidths string
Dim lngI As Long                                'For Loop counter
Dim lngJ As Long                                'Columnwidths counter

    'Get the column widths:-
    'THIS CODE BY JAMIE CZERNIK------------------------------------------->
    'Replace the Split() function as not available in Access 97:
    'Sets the array to one entry.
    ReDim astrColumnWidths(0)
    'Loops through each character in the Column Widths String passed in by the calling code.
    For lngI = 1 To VBA.Len(strColumnWidths)
        'Looks for each semicolon, which is what separates the individual Column Widths.
        Select Case VBA.Mid(strColumnWidths, lngI, 1)
            'If a semicolon is not found, the character is added to the any characters
            ' already in the columnwidths entry in the array.  If it is found, the
            ' Columnwidths Counter is incremented by one and the array is increased by
            ' one while retaining entered data so that the next columnwidth can be entered.
            Case Is <> ";"
                astrColumnWidths(lngJ) = astrColumnWidths(lngJ) & VBA.Mid( _
                strColumnWidths, lngI, 1)
            Case ";"
                lngJ = lngJ + 1
                ReDim Preserve astrColumnWidths(lngJ) 'Resize the array.
        End Select
    Next lngI
    'Resets the loop counter to 0.
    lngI = 0
    '--------------------------------------------> END CODE BY JAMIE CZERNIK.
    'Access 2000/2002 users can uncomment the line below and remove the split() code
    'replacement above.
    'astrColumnWidths = Split(strColumnWidths, ";")'Available in Access 2000/2002 only
    strTemp = VBA.vbNullString 'Sets the temp variable to a null string
    'Loops through the all the columnwidths in the array, converting them to the new sizes
    ' (using the Width Size Conversion Factor that was passed-in), and recombining them
    ' into a single string to pass back to the calling code. (If there is no Column Width,
    ' the value is left blank.)
    Do Until lngI > UBound(astrColumnWidths)
        If Not IsNull(astrColumnWidths(lngI)) And astrColumnWidths(lngI) <> "" Then
            strTemp = strTemp & CSng(astrColumnWidths(lngI)) * sngFactor & ";"
        End If
        lngI = lngI + 1
    Loop
    'Returns the combined columnwidths string to the calling code.
    adjustColumnWidths = strTemp
    Erase astrColumnWidths 'Destroy array.
    
Exit_adjustColumnWidths:
    On Error Resume Next
    Exit Function

Err_adjustColumnWidths:
    Erase astrColumnWidths 'Destroy array.
    Resume Exit_adjustColumnWidths
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : getOrigWindow
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Routine stores the original window dimensions before resizing call it
'             when form loads. (before calling ResizeForm Me!).
'             Call it: Form_Load()
'             [More info in "Important Points" - point 5 - in help file.]
'---------------------------------------------------------------------------------------
Public Sub getOrigWindow(frm As Form)

On Error Resume Next

    OrigWindow.Height = frm.WindowHeight
    OrigWindow.Width = frm.WindowWidth

End Sub

'---------------------------------------------------------------------------------------
' Procedure : RestoreWindow
' DateTime  : 27/01/2003
' Author    : Jamie Czernik
' Purpose   : Routine restores the original window dimensions call it when form closes.
'             Call it: Form_Close()
'             [More info in "Important Points" - point 5 - in help file.]
'---------------------------------------------------------------------------------------
Public Sub RestoreWindow()

On Error Resume Next

    'DoCmd.MoveSize , , OrigWindow.Width, OrigWindow.Height
   ' Access.DoCmd.Save
    
End Sub



