VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRegOCX 
   Caption         =   "Registriraj/Unregistriraj OCX/DLL"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDialogFileSpec 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Dialog for file spec"
      Top             =   720
      Width           =   315
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Izhod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4110
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnregister 
      Caption         =   "Odregistriraj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Registriraj"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtOCXPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   690
      Width           =   5535
   End
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H80000018&
      Caption         =   "Zacni"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000018&
      Caption         =   "Preklici"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblPleaseVerify 
      Caption         =   "¸Preveri verzijo ce obstaja"
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   2925
   End
   Begin VB.Label lblFileInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   180
      TabIndex        =   6
      Top             =   1560
      Width           =   5505
   End
   Begin VB.Label lblOCXFileSpec 
      Caption         =   "File doloèi fele ocx/dll"
      Height          =   465
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   5505
   End
End
Attribute VB_Name = "frmRegOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RegOCX.frm
'
' By Herman Liu
'
' To register or unregister OCX/DLL controls (1) With prior confirmation of the
' Product Version and (2) Without using Regsvr32.exe.  This code has the advantages
' of (a) user verification of the specific version being registered; (b) being free
' from the existence of a Regsvr32.exe file, speedier and a better error handling.
'
' Output: Entries entered into/removed from HKEY_CLASSES_ROOT in the registry.
'
Option Explicit

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpdata As Any) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
  (ByVal lpLibFileName As String) As Long
  
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
   
'Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, _
   ByVal dwExitCode As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
   
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim mCompanyName As String
Dim mProductVersion As String
Dim RegFlag As Boolean
Dim UnregFlag As Boolean
Dim mresult
Dim gcdg As Object



Private Sub Form_Load()
    lblOCXFileSpec = "File specification of OCX/DLL" & vbCrLf & _
       "    e.g. c:\Windows\System\mswinsck.ocx"
    lblPleaseVerify.Visible = False
    Set gcdg = CommonDialog1
    setButtons True
End Sub



Private Sub cmdRegister_Click()
    If Trim(txtOCXPath.Text) = "" Then
        MsgBox "No path and filename yet"
        Exit Sub
    ElseIf IsFileThere(txtOCXPath.Text) = False Then
        MsgBox "File not found"
        Exit Sub
    End If
    DispProdVersion txtOCXPath.Text
    RegFlag = True
    UnregFlag = False
    setButtons False
End Sub



Private Sub cmdUnregister_Click()
    If Trim(txtOCXPath.Text) = "" Then
        MsgBox "No path and filename yet"
        Exit Sub
    ElseIf IsFileThere(txtOCXPath.Text) = False Then
        MsgBox "File not found"
        Exit Sub
    End If
    RegFlag = False
    UnregFlag = True
    setButtons False
End Sub
    
    
    
Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdProceed_click()
    If RegFlag Then
         RegUnReg txtOCXPath.Text
         lblPleaseVerify.Visible = False
         lblFileInfo.Caption = ""
    Else
         RegUnReg txtOCXPath.Text, "u"
    End If
    setButtons True
End Sub



Private Sub cmdCancel_click()
    txtOCXPath_Change
    setButtons True
End Sub



Private Sub setButtons(OnOff)
    cmdRegister.Visible = OnOff
    cmdUnregister.Visible = OnOff
    cmdExit.Visible = OnOff
    cmdProceed.Visible = Not OnOff
    cmdCancel.Visible = Not OnOff
    If OnOff = True Then
        cmdDialogFileSpec.Enabled = True
        txtOCXPath.Locked = False
    Else
        cmdDialogFileSpec.Enabled = False
        txtOCXPath.Locked = True
    End If
End Sub



Private Sub txtOCXPath_Change()
    lblPleaseVerify.Visible = False
    lblFileInfo.Caption = ""
End Sub

    
    
Private Sub cmdDialogFileSpec_Click()
    On Error GoTo errHandler
    gcdg.Filter = "(*.ocx)|*.ocx|(*.dll)|*.dll|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.DefaultExt = "ocx"
    gcdg.Flags = cdlOFNFileMustExist
    gcdg.FileName = ""
    gcdg.CancelError = True
    gcdg.ShowOpen
    If gcdg.FileName = "" Then
        txtOCXPath.Text = ""
        Exit Sub
    End If
    txtOCXPath.Text = gcdg.FileName
    Exit Sub
    
errHandler:
    If Err <> 32755 Then
         ErrMsgProc "cmdDialogFileSpec_Click"
    End If
End Sub



Private Sub DispProdVersion(inFile As String)
    If Not GetFileInfo(inFile) Then
        lblPleaseVerify.Visible = False
        lblFileInfo.Caption = "(No Product Version available for this file)"
    Else
        lblPleaseVerify.Visible = True
        lblFileInfo.Caption = "Company Name:  " & mCompanyName & vbCrLf & _
             "Product Version:  " & mProductVersion
    End If
End Sub



Private Function GetFileInfo(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim lInfoSize As Long
    Dim lpHandle As Long
    Dim strFileInfoString As String
    Dim i As Integer
    
    GetFileInfo = False                                ' Assume
    
     ' GetFileVersionInfoSize determines if system can obtain version info
     ' about the specified file.  If yes, it returns its size in bytes and
     ' a handle to the data.
    lpHandle = 0
    lInfoSize = GetFileVersionInfoSize(inFileSpec, lpHandle)
    If lInfoSize = 0 Then
        Exit Function
    End If

     ' We pass the file name, size(ignored), size of buffer and the buffer of
     ' version info to GetFileVersionInfo, which will fill the buffer with
     ' version info about the file. (Modified here).
    strFileInfoString = String(lInfoSize, 0)
    mresult = GetFileVersionInfo(ByVal inFileSpec, 0&, ByVal lInfoSize, _
          ByVal strFileInfoString)
    If mresult = 0 Then
        Exit Function
    End If

     ' We now have a block of version data, in an unreadable format though. If you
     ' wish, you may check the existence of "StringFileInfo" with InStr function.
     ' Normally we must call VerQueryValue to read selected pieces of data of the
     ' above, with arguments such as "\VarFileInfo\Translation" or "\StringFileInfo
     ' \lang-codepage\string-name" where lang-codepage is a code which has yet to be
     ' obtained from first 2 words(high-low) returned by "\VarFileInfo\Translation"
     ' from the strFileInfoString (and padded to fixed 8-digit), and string-name is
     ' one of predefined string names such as "CompanyName" & "FileDescription", etc.
     ' However, the following simple alternative is OK for our purpose.

     mCompanyName = ""
     mProductVersion = ""
     i = InStr(strFileInfoString, "CompanyName")
     If i > 0 Then
         i = i + 12
         mCompanyName = Mid$(strFileInfoString, i, 21)
     End If
     i = InStr(strFileInfoString, "FileDescription")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "FileVersion")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "InternalName")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "LegalCopyright")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "OriginalFilename")
     If i > 0 Then
         i = i + 20
     End If
     i = InStr(strFileInfoString, "ProductName")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "ProductVersion")
     If i > 0 Then
         i = i + 16
         mProductVersion = Mid$(strFileInfoString, i)
     End If

     If Trim(mProductVersion) <> "" Then
         GetFileInfo = True
     End If
End Function


    
    
Private Sub RegUnReg(ByVal inFileSpec As String, Optional inHandle As String = "")
    On Error Resume Next
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread
    
      ' Load the control DLL, i. e. map the specified DLL file into the
      ' address space of the calling process
    lLib = LoadLibrary(inFileSpec)
    If lLib = 0 Then
         ' e.g. file not exists or not a valid DLL file
        MsgBox "Failure loading control DLL"
        Exit Sub
    End If
    
      ' Find and store the DLL entry point, i.e. obtain the address of the
      ' “DllRegisterServer” or "DllUnregisterServer" function (to register
      ' or deregister the server’s components in the registry).
      '
    If inHandle = "" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf inHandle = "U" Or inHandle = "u" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
        MsgBox "Unknown command handle"
        Exit Sub
    End If
    If lpDLLEntryPoint = vbNull Then
        GoTo earlyExit1
    End If
    
    Screen.MousePointer = vbHourglass
    
      ' Create a thread to execute within the virtual address space of the calling process
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
      ' Use WaitForSingleObject to check the return state (i) when the specified object
      ' is in the signaled state or (ii) when the time-out interval elapses.  This
      ' function can be used to test Process and Thread.
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
      ' We don't call the dangerous TerminateThread(); after the last handle
      ' to an object is closed, the object is removed from the system.
    CloseHandle mThread
    FreeLibrary lLib
    
    Screen.MousePointer = vbDefault
    MsgBox "Process completed"
    Exit Sub
    
    
earlyExit1:
    Screen.MousePointer = vbDefault
    MsgBox "Process failed in obtaining entry point or creating thread."
     ' Decrements the reference count of loaded DLL module before leaving
    FreeLibrary lLib
    Exit Sub
    
earlyExit2:
    Screen.MousePointer = vbDefault
    MsgBox "Process failed in signaled state or time-out."
    FreeLibrary lLib
     ' Terminate the thread to free up resources that are used by the thread
     ' NB Calling ExitThread for an application's primary thread will cause
     ' the application to terminate
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub


