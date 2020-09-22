VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1998-2003
'/*************************************/

'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit
'
' --- PRIVATE VARIABLES
'
Private intMinPage    As Integer
Private intMaxPage    As Integer
Private intFromPage   As Integer
Private intToPage     As Integer
Private intCopies     As Integer
Private intMaxCopies  As Integer
Private OwnerhWnd     As Long
Private OrientationOk As Boolean
Private ShowPrintToFile As Boolean
'Private strNewPrinterName As String
'
' --- PUBLIC VARIABLES
'
Public flags As PrinterConstants
Public CancelError As Boolean
'
' --- API CONSTANTS
'
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
'
' --- API TYPES DEFINITION
'
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
'
' --- API DECLARATIONS
'
Private Declare Function PrintDialog Lib "COMDLG32.DLL" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'
' --- PUBLIC ENUM
'
Public Enum PrinterConstants
    PD_ALLPAGES = &H0
    PD_COLLATE = &H10
    PD_DISABLEPRINTTOFILE = &H80000
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
    PD_NOPAGENUMS = &H8
    PD_NOSELECTION = &H4
    PD_NOWARNING = &H80
    PD_PAGENUMS = &H2
    PD_PRINTSETUP = &H40
    PD_PRINTTOFILE = &H20
    PD_RETURNDC = &H100
    PD_RETURNDEFAULT = &H400
    PD_RETURNIC = &H200
    PD_SELECTION = &H1
    PD_SHOWHELP = &H800
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    DLG_PRINT = 0
    DLG_PRINTSETUP = &H40
End Enum

Public Enum ErrorConstants
    cdlCancel = 32755
End Enum
Private Sub Class_Initialize()
  Dim objPrinter As Printer
  Dim sPrinterName As String
 
    intMinPage = 0
    intMaxPage = 0
    intFromPage = 0
    intToPage = 0
    intCopies = 1
    intMaxCopies = 1
    OwnerhWnd = 0
    OrientationOk = True
    CancelError = False
    ShowPrintToFile = False
    
    On Error Resume Next
    sPrinterName = GetSetting(App.title, "Options", "Printer", "None")
    'strNewPrinterName = UCase(Printer.DeviceName)
    
'    For Each objPrinter In Printers
'        If objPrinter.DeviceName = sPrinterName Then
'            Set Printer = objPrinter
'        End If
'    Next
    On Error GoTo 0

End Sub
'
' -- PUBLIC MEMBERS
'
Property Get MinPage() As Integer
  MinPage = intMinPage
End Property
Property Let MinPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMinPage = intNewValue
  If intNewValue > intFromPage Then intFromPage = intNewValue
  If intNewValue > intToPage Then intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
End Property
Property Get FromPage() As Integer
  FromPage = intFromPage
End Property
Property Let FromPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intFromPage = intNewValue
  If intNewValue > intToPage Then intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
End Property
Property Get ToPage() As Integer
  ToPage = intToPage
End Property
Property Let ToPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
  If intNewValue < intFromPage Then intFromPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
End Property
Property Get MaxPage() As Integer
  MaxPage = intMaxPage
End Property
Property Let MaxPage(ByVal intNewValue As Integer)
  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMaxPage = intNewValue
  If intNewValue < intToPage Then intToPage = intNewValue
  If intNewValue < intFromPage Then intFromPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
End Property
Public Function ShowPrinter(ByVal flags As PrinterConstants) As Boolean
  Dim PrintDlg As PRINTDLG_TYPE
  Dim DevMode As DEVMODE_TYPE
  Dim DevName As DEVNAMES_TYPE
  Dim lpDevMode As Long, lpDevName As Long
  Dim intReturn As Integer
  Dim objPrinter As Printer
  Dim blnCancel As Boolean
  Dim strNewPrinterName As String
  
    If Not ShowPrintToFile Then flags = flags + PD_HIDEPRINTTOFILE
    
    blnCancel = False
    'Use PrintDialog to get the handle to a memory
    'block with a DevMode and DevName structures
    With PrintDlg
        .lStructSize = Len(PrintDlg)
        .hWndOwner = OwnerhWnd
        .flags = flags
        .nMinPage = intMinPage
        .nFromPage = intFromPage
        .nToPage = intToPage
        .nMaxPage = intMaxPage
        .nCopies = intCopies
    End With
    
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmColor = Printer.ColorMode
    
    On Error Resume Next
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0
    
    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        intReturn = GlobalUnlock(lpDevMode)
    End If
    
    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        intReturn = GlobalUnlock(lpDevName)
    End If
    
    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) Then
        
        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        intReturn = GlobalUnlock(lpDevName)
        
        With PrintDlg
            flags = .flags
            intFromPage = .nFromPage
            intToPage = .nToPage
        End With
        GlobalFree PrintDlg.hDevNames
        
        'Get the DevMode structure and set the printer properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        intReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        strNewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        
        On Error Resume Next
            'Set the printer
            If UCase(Printer.DeviceName) <> strNewPrinterName Then
                For Each objPrinter In Printers
                    If UCase$(objPrinter.DeviceName) = strNewPrinterName Then Set Printer = objPrinter
                Next
            End If
            'Set the printer properties modified by the user
            With Printer
                .ColorMode = DevMode.dmColor
                .Copies = IIf(DevMode.dmCopies > intMaxCopies, intMaxCopies, DevMode.dmCopies)
                .PaperBin = DevMode.dmDefaultSource
                .Duplex = DevMode.dmDuplex
                If OrientationOk Then .Orientation = DevMode.dmOrientation
                .PaperSize = DevMode.dmPaperSize
                .PrintQuality = DevMode.dmPrintQuality
                .Zoom = DevMode.dmScale
            End With
        On Error GoTo 0
      
    Else
        
        GlobalFree PrintDlg.hDevMode
        GlobalFree PrintDlg.hDevNames
        blnCancel = True
        If CancelError Then err.Raise cdlCancel, "PrintDialog", "Cancel."
    
    End If
    
    ShowPrinter = blnCancel
  
End Function



Public Property Get Copies() As Variant
    Copies = intCopies
End Property

Public Property Let Copies(ByVal vNewValue As Variant)
    intCopies = IIf(vNewValue > 1, vNewValue, 1)
End Property

Public Property Get Owner_hWnd() As Variant

End Property

Public Property Let Owner_hWnd(ByVal vNewValue As Variant)
    OwnerhWnd = vNewValue
End Property

Public Property Get MaxCopies() As Variant
    MaxCopies = intMaxCopies
End Property

Public Property Let MaxCopies(ByVal vNewValue As Variant)
    intMaxCopies = IIf(vNewValue > 1, vNewValue, 1)
End Property

Public Property Get AllowOrientation() As Variant
    AllowOrientation = OrientationOk
End Property

Public Property Let AllowOrientation(ByVal vNewValue As Variant)
    OrientationOk = CBool(vNewValue)
End Property

Private Sub Class_Terminate()
    SaveSetting App.title, "Options", "Printer", Printer.DeviceName
End Sub



Public Property Get AllowPrintToFile() As Variant
    AllowOrientation = ShowPrintToFile
End Property

Public Property Let AllowPrintToFile(ByVal vNewValue As Variant)
    ShowPrintToFile = CBool(vNewValue)
End Property
