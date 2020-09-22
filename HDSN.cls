VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'

' Costants
Private Const VER_PLATFORM_WIN32S = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Costants for driver IDE
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

' Costants to create file
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const CREATE_NEW = 1

' enumeration for CmnGetHDData
Private Enum HDINFO
    HD_MODEL_NUMBER
    HD_SERIAL_NUMBER
    HD_FIRMWARE_REVISION
End Enum

' structure for OS Info data
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' Structure for SENDCMDINPARAMS
Private Type IDEREGS
    bFeaturesReg As Byte
    bSectorCountReg As Byte
    bSectorNumberReg As Byte
    bCylLowReg As Byte
    bCylHighReg As Byte
    bDriveHeadReg As Byte
    bCommandReg As Byte
    bReserved As Byte
End Type

' Structure for driver of IDE
Private Type SENDCMDINPARAMS
    cBufferSize As Long
    irDriveRegs As IDEREGS
    bDriveNumber As Byte
    bReserved(1 To 3) As Byte
    dwReserved(1 To 4) As Long
End Type

' Structure for SENDCMDOUTPARAMS
Private Type DRIVERSTATUS
    bDriveError As Byte
    bIDEStatus As Byte
    bReserved(1 To 2) As Byte
    dwReserved(1 To 2) As Long
End Type

' Structure for driver IDE
Private Type SENDCMDOUTPARAMS
    cBufferSize As Long
    DStatus As DRIVERSTATUS
    bBuffer(1 To 512) As Byte
End Type

'
Private Declare Function GetVersionEx _
    Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

'
Private Declare Function CreateFile _
    Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFilename As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

'
Private Declare Function CloseHandle _
    Lib "kernel32" _
    (ByVal hObject As Long) As Long

'
Private Declare Function DeviceIoControl _
    Lib "kernel32" _
    (ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, _
    ByVal lpOverlapped As Long) As Long
    
'
Private Declare Sub ZeroMemory _
    Lib "kernel32" Alias "RtlZeroMemory" _
    (dest As Any, _
    ByVal numBytes As Long)

'
Private Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Declare Function GetLastError _
    Lib "kernel32" () As Long

Private mvarCurrentDrive As Byte    '
Private mvarPlatform As String      '

Public Property Get Copyright() As String
    
    ' Copyright
    Copyright = "Every VB Developer"
    
End Property

' Method GetModelNumber
Public Function GetModelNumber() As String
    
    '
    GetModelNumber = CmnGetHDData(HD_MODEL_NUMBER)
    
End Function

' Method GetSerialNumber
Public Function GetSerialNumber() As Long
      Dim DriveLetter As String
      Dim DriveSerial As Long
      DriveLetter = ""
          Dim fso As Object, Drv As Object
          
          'Create a FileSystemObject object
          Set fso = CreateObject("Scripting.FileSystemObject")
          
          'Assign the current drive letter if not specified
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(App.path))
          End If
      
          With Drv
              If .IsReady Then
                  DriveSerial = Abs(.SerialNumber)
              Else    '"Drive Not Ready!"
                  DriveSerial = -1
              End If
          End With
          
          'Clean up
          Set Drv = Nothing
          Set fso = Nothing
          
          GetSerialNumber = DriveSerial
          
      End Function



' Method GetFirmwareRevision
Public Function GetFirmwareRevision() As String
    
    '
    GetFirmwareRevision = CmnGetHDData(HD_FIRMWARE_REVISION)
    
End Function

' Property CurrentDrive
Public Property Let CurrentDrive(ByVal vData As Byte)
    
    '
    If vData < 0 Or vData > 3 Then
        err.Raise 10000, , "Illegal drive number"   ' IDE drive 0..3
    End If
    
    '
    mvarCurrentDrive = vData

End Property

' Property CurrentDrive
Public Property Get CurrentDrive() As Byte
    
    '
    CurrentDrive = mvarCurrentDrive

End Property

' Property Platform
Public Property Get Platform() As String
    
    '
    Platform = mvarPlatform

End Property

Private Sub Class_Initialize()

    '
    Dim os As OSVERSIONINFO
        
    os.dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)
    mvarPlatform = "Unk"
    Select Case os.dwPlatformId
        Case Is = VER_PLATFORM_WIN32S
            mvarPlatform = "32S"                ' Win32S
        Case Is = VER_PLATFORM_WIN32_WINDOWS
            If os.dwMinorVersion = 0 Then
                mvarPlatform = "W95"            ' Win 95
            Else
                mvarPlatform = "W98"            ' Win 98
            End If
        Case Is = VER_PLATFORM_WIN32_NT
            mvarPlatform = "WNT"                ' Win NT/2000
    End Select

End Sub

Private Function CmnGetHDData(hdi As HDINFO) As String

    ' IDE
    
    Dim bin As SENDCMDINPARAMS
    Dim bout As SENDCMDOUTPARAMS
    Dim hdh As Long
    Dim br As Long
    Dim ix As Long
    Dim hddfr As Long
    Dim hddln As Long
    Dim s As String
    
    Select Case hdi
        Case HD_MODEL_NUMBER
            hddfr = 55
            hddln = 40
        Case HD_SERIAL_NUMBER
            hddfr = 21
            hddln = 20
        Case HD_FIRMWARE_REVISION
            hddfr = 47
            hddln = 8
        Case Else
            err.Raise 10001, "Illegal HD Data type"
    End Select
    
    Select Case mvarPlatform
        Case "WNT"

            hdh = CreateFile("\\.\PhysicalDrive" & mvarCurrentDrive, _
                GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ + FILE_SHARE_WRITE, _
                0, OPEN_EXISTING, 0, 0)
        Case "W95", "W98"

            hdh = CreateFile("\\.\Smartvsd", _
                0, 0, 0, CREATE_NEW, 0, 0)
        Case Else

            err.Raise 10002, , "Illegal platform (only WNT, W98 or W95)"    ' Altre piattaforme non gestite
    End Select

    If hdh = 0 Then
        err.Raise 10003, , "Error on CreateFile"
    End If
    
    
    ZeroMemory bin, Len(bin)
    ZeroMemory bout, Len(bout)
    

    With bin
        .bDriveNumber = mvarCurrentDrive
        .cBufferSize = 512
        With .irDriveRegs
            If (mvarCurrentDrive And 1) Then
                .bDriveHeadReg = &HB0
            Else
                .bDriveHeadReg = &HA0
            End If
            .bCommandReg = &HEC
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
        End With
    End With
    

    DeviceIoControl hdh, DFP_RECEIVE_DRIVE_DATA, _
                    bin, Len(bin), bout, Len(bout), br, 0
    



    s = ""
    For ix = hddfr To hddfr + hddln - 1 Step 2
        If bout.bBuffer(ix + 1) = 0 Then Exit For
        s = s & Chr(bout.bBuffer(ix + 1))
        If bout.bBuffer(ix) = 0 Then Exit For
        s = s & Chr(bout.bBuffer(ix))
    Next ix
    

    CloseHandle hdh


    CmnGetHDData = Trim(s)
    
End Function

