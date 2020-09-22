Attribute VB_Name = "Module3"

      Private Declare Function GetDesktopWindow Lib "user32" () As Long
 Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long


Private Declare Function GetTempPath Lib "kernel32" _
Alias "GetTempPathA" (ByVal nBufferLength As Long, _
ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName _
Lib "kernel32" Alias "GetTempFileNameA" _
(ByVal lpszPath As String, _
ByVal lpPrefixString As String, _
ByVal wUnique As Long, _
ByVal lpTempFileName As String) As Long
Private Declare Function ConvertUncompressedSnapshot Lib "StrStorage.dll" _
    (ByVal UnCompressedSnapShotName As String, _
    ByVal OutputPDFname As String, _
    Optional ByVal CompressionLevel As Long = 0, _
    Optional ByVal PasswordOpen As String = "", _
    Optional ByVal PasswordOwner As String = "", _
    Optional ByVal PasswordRestrictions As Long = 0, _
    Optional ByVal PDFNoFontEmbedding As Long = 0, _
    Optional ByVal PDFUnicodeFlags As Long = 0 _
    ) As Boolean
    Private Declare Function FreeLibrary Lib "kernel32" _
(ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" _
Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function SetupDecompressOrCopyFile _
Lib "setupAPI" _
Alias "SetupDecompressOrCopyFileA" ( _
ByVal SourceFileName As String, _
ByVal TargetFileName As String, _
ByVal CompressionType As Integer) As Long

Private Declare Function SetupGetFileCompressionInfo _
Lib "setupAPI" _
Alias "SetupGetFileCompressionInfoA" ( _
ByVal SourceFileName As String, _
TargetFileName As String, _
SourceFileSize As Long, _
DestinationFileSize As Long, _
CompressionType As Integer _
) As Long

 
'Compression types
Private Const FILE_COMPRESSION_NONE = 0
Private Const FILE_COMPRESSION_WINLZA = 1
Private Const FILE_COMPRESSION_MSZIP = 2

Private Const Pathlen = 256
Private Const MaxPath = 256

' Note: I converted the Enums to Constants to allow for use in Access 97.

'Enum TDocumentInfo 'Coming Soon!
 '  diAuthor
 '  diCreator
 '  diKeywords
 '  diProducer
 '  diSubject
 '  diTitle
 '  diCompany
 '  diPDFX_Ver ' GetInDocInfo() only -> The PDF/X version is set by SetPDFVersion()!
 '  diCustom   ' User defined key
'End Enum








'  Device Parameters for GetDeviceCaps()
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

' ***********************************************
'       Font, DC and TextWidth stuff

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
 
Private Const LF_FACESIZE = 32
 
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" _
(ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
 
Private Declare Function apiCreateFontIndirect Lib "gdi32" Alias _
        "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
 
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" _
(ByVal hdc As Long, _
ByVal hObject As Long) As Long
 
Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long
 
Private Declare Function apiMulDiv Lib "kernel32" Alias "MulDiv" _
(ByVal nNumber As Long, _
ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
 
Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hwnd As Long) As Long
 
Private Declare Function apiReleaseDC Lib "user32" _
 Alias "ReleaseDC" (ByVal hwnd As Long, _
 ByVal hdc As Long) As Long
  
Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" _
(ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function CreateDCbyNum Lib "gdi32" Alias "CreateDCA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As Long, ByVal lpInitData As Long) As Long  'DEVMODE) As Long




' CONSTANTS
Private Const TWIPSPERINCH = 1440
' Used to ask System for the Logical pixels/inch in X & Y axis
'Private Const LOGPIXELSY = 90
'Private Const LOGPIXELSX = 88
 
' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_NOCLIP = &H100



' Font stuff
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_PRECIS = 4
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_OUTLINE_PRECIS = 8

Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_MASK = &HF
Private Const CLIP_LH_ANGLES = 16
Private Const CLIP_TT_ALWAYS = 32
Private Const CLIP_EMBEDDED = 128

Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255

' ***********************************************




' Allow user to set FileName instead
' of using API Temp Filename or
' popping File Dialog Window
Private mSaveFileName As String

' Full path and name of uncompressed SnapShot file
Private mUncompressedSnapFile As String

' Name of the Report we ' working with
Private mReportName As String

' Instance returned from LoadLibrary calls
Private hLibDynaPDF As Long
Private hLibStrStorage As Long

    
    



Private Function LoadLib() As Boolean
Dim s As String
Dim blRet As Boolean

On Error Resume Next

' *** Please Note ***
' If you are going to process many reports at once then to improve performance you
' should only call LoadLib once.

' May 16/2008
' Always look in the folder where this MDB resides First before checking the System folder.

LoadLib = False

' If we aready loaded then free the library
If hLibDynaPDF <> 0 Then
    hLibDynaPDF = FreeLibrary(hLibDynaPDF)
End If


' Our error string
s = "Sorry...cannot find the DynaPDF.dll file" & vbCrLf
s = s & "Please copy the DynaPDF.dll file into the same folder as this Access MDB or your Windows System32 folder."

' OK Try to load the DLL assuming it is in the same folder as this MDB.
' CurrentDB works with both A97 and A2K or higher
hLibDynaPDF = LoadLibrary(App.path & "DynaPDF.dll")
    
If hLibDynaPDF = 0 Then
    ' OK Try to load the DLL assuming it is in the Window System folder
    hLibDynaPDF = LoadLibrary(App.path & "\DynaPDF.dll")
End If

If hLibDynaPDF = 0 Then
    MsgBox s, vbOKOnly, "MISSING DynaPDF.dll FILE"
    LoadLib = False
    Exit Function
End If



'' ** Commented out for Debugging only - Must be active
'' ***************************************************************************
'
' Load StrStorage.DLL
' If we aready loaded then free the library
If hLibStrStorage <> 0 Then
    hLibStrStorage = FreeLibrary(hLibStrStorage)
End If


' Our error string
s = "Sorry...cannot find the StrStorage.dll file" & vbCrLf
s = s & "Please copy the StrStorage.dll file into the same folder as this Access MDB or your Windows System32 folder."

' OK Try to load the DLL assuming it is in the same folder as this MDB.
' CurrentDB works with both A97 and A2K or higher
hLibStrStorage = LoadLibrary(App.path & "StrStorage.dll")

If hLibStrStorage = 0 Then
    ' OK Try to load the DLL assuming it is in the Window System folder
    hLibStrStorage = LoadLibrary(App.path & "\StrStorage.dll")
End If

If hLibStrStorage = 0 Then
    MsgBox s, vbOKOnly, "MISSING StrStorage.dll FILE"
    LoadLib = False
    Exit Function
End If

' RETURN SUCCESS
LoadLib = True
End Function
Public Function ConvertReportToPDF( _
Optional RptName As String = "", _
Optional SnapshotName As String = "", _
Optional OutputPDFname As String = "", _
Optional ShowSaveFileDialog As Boolean = False, _
Optional StartPDFViewer As Boolean = True, _
Optional CompressionLevel As Long = 0, _
Optional PasswordOpen As String = "", _
Optional PasswordOwner As String = "", _
Optional PasswordRestrictions As Long = 0, _
Optional PDFNoFontEmbedding As Long = 0, _
Optional PDFUnicodeFlags As Long = 0 _
) As Boolean


' RptName is the name of a report contained within this MDB
' SnapshotName is the name of an existing Snapshot file
' OutputPDFname is the name you select for the output PDF file
' ShowSaveFileDialog is a boolean param to specify whether or not to display
' the standard windows File Dialog window to select an exisiting Snapshot file
' CompressionLevel - not hooked up yet
' PasswordOwner  - not hooked up yet
' PasswordOpen - not hooked up yet
' PasswordRestrictions - not hooked up yet
' PDFNoFontEmbedding - Do not Embed fonts in PDF. Set to 1 to stop the
' default process of embedding all fonts in the output PDF. If you are
' using ONLY - any of the standard Windows fonts
' using ONLY - any of the standard 14 Fonts natively supported by the PDF spec
'The 14 Standard Fonts
'All version of Adobe's Acrobat support 14 standard fonts. These fonts are always available
'independent whether they're embedded or not.
'Family name PostScript name Style
'Courier Courier fsNone
'Courier Courier-Bold fsBold
'Courier Courier-Oblique fsItalic
'Courier Courier-BoldOblique fsBold + fsItalic
'Helvetica Helvetica fsNone
'Helvetica Helvetica-Bold fsBold
'Helvetica Helvetica-Oblique fsItalic
'Helvetica Helvetica-BoldOblique fsBold + fsItalic
'Times Times-Roman fsNone
'Times Times-Bold fsBold
'Times Times-Italic fsItalic
'Times Times-BoldItalic fsBold + fsItalic
'Symbol Symbol fsNone, other styles are emulated only
'ZapfDingbats ZapfDingbats fsNone, other styles are emulated only




Dim s As String
Dim blRet As Boolean
' Let's see if the DynaPDF.DLL is available.
blRet = LoadLib()
If blRet = False Then
    ' Cannot find DynaPDF.dll or StrStorage.dll file
    Exit Function
End If

On Error GoTo ERR_CREATSNAP

Dim strPath  As String
Dim strPathandFileName  As String
Dim strEMFUncompressed As String

Dim sOutFile As String
Dim lngRet As Long

' Init our string buffer
strPath = Space(Pathlen)

'Save the ReportName to a local var
mReportName = RptName

' Let's kill any existing Temp SnapShot file
If Len(mUncompressedSnapFile & vbNullString) > 0 Then
    Kill mUncompressedSnapFile
    mUncompressedSnapFile = ""
End If

' If we have been passed the name of a Snapshot file then
' skip the Snapshot creation process below
If Len(SnapshotName & vbNullString) = 0 Then
      
    ' Make sure we were passed a ReportName
    If Len(RptName & vbNullString) = 0 Then
        ' No valid parameters - FAIL AND EXIT!!
        ConvertReportToPDF = ""
        Exit Function
    End If
        
    ' Get the Systems Temp path
    ' Returns Length of path(num characters in path)
    lngRet = GetTempPath(Pathlen, strPath)
    ' Chop off NULLS and trailing "\"
    strPath = Left(strPath, lngRet) & Chr(0)
    
    ' Now need a unique Filename
    ' locked from a previous aborted attemp.
    ' Needs more work!
    strPathandFileName = GetUniqueFilename(strPath, "SNP" & Chr(0), "snp")
    
    ' Export the selected Report to SnapShot format
    DoCmd.OutputTo acOutputReport, RptName, "SnapshotFormat(*.snp)", _
       strPathandFileName
    ' Make sure the process has time to complete
    DoEvents

Else
    strPathandFileName = SnapshotName
 
End If

' Let's decompress into same filename but change type to ".tmp"
'strEMFUncompressed = Mid(strPathandFileName, 1, Len(strPathandFileName) - 3)
'strEMFUncompressed = strEMFUncompressed & "tmp"
Dim sPath As String * 512
lngRet = GetTempPath(512, sPath)

strEMFUncompressed = GetUniqueFilename(sPath, "SNP", "tmp")

lngRet = SetupDecompressOrCopyFile(App.path & "\" & strPathandFileName, strEMFUncompressed, 0&)

If lngRet <> 0 Then
    err.Raise vbObjectError + 525, "ConvertReportToPDF.SetupDecompressOrCopyFile", _
    "Sorry...cannot Decompress SnapShot File" & vbCrLf & _
    "Please select a different Report to Export"
End If

' Set our uncompressed SnapShot file name var
mUncompressedSnapFile = strEMFUncompressed

' Remember to Cleanup our Temp SnapShot File if we were NOT passed the
' Snapshot file as the optional param
If Len(SnapshotName & vbNullString) = 0 Then
    Kill strPathandFileName
End If


' Do we name output file the same as the input file name
' and simply change the file extension to .PDF or

    ' let's decompress into same filename but change type to ".tmp"
    ' But first let's see if we were passed an output PDF file name
    If Len(OutputPDFname & vbNullString) = 0 Then
        sOutFile = Mid(strPathandFileName, 1, Len(strPathandFileName) - 3)
        sOutFile = sOutFile & "PDF"
    Else
        sOutFile = OutputPDFname
    End If



' Call our function in the StrStorage DLL
' Note the Compression and Password params are not hooked up yet.
blRet = ConvertUncompressedSnapshot(mUncompressedSnapFile, sOutFile, _
CompressionLevel, PasswordOpen, PasswordOwner, PasswordRestrictions, PDFNoFontEmbedding, PDFUnicodeFlags)

If blRet = False Then
err.Raise vbObjectError + 526, "ConvertReportToPDF.ConvertUncompressedSnaphot", _
    "Sorry...damaged SnapShot File" & vbCrLf & _
    "Please select a different Report to Export"
End If

' Do we open new PDF in registered PDF viewer on this system?
If StartPDFViewer = True Then
'MsgBox (sOutFile)
Dim Scr_hDC As Long
          Scr_hDC = GetDesktopWindow()
          StartDoc = ShellExecute(Scr_hDC, "Open", sOutFile, _
          "", App.path, SW_SHOWNORMAL)
End If

' Success
ConvertReportToPDF = True


EXIT_CREATESNAP:

' Let's kill any existing Temp SnapShot file
'If Len(mUncompressedSnapFile & vbNullString) > 0 Then
     On Error Resume Next
   Kill mUncompressedSnapFile
    mUncompressedSnapFile = ""
'End If

' If we aready loaded then free the library
If hLibStrStorage <> 0 Then
    hLibStrStorage = FreeLibrary(hLibStrStorage)
End If

If hLibDynaPDF <> 0 Then
    hLibDynaPDF = FreeLibrary(hLibDynaPDF)
End If

Exit Function

ERR_CREATSNAP:
MsgBox err.Description, vbOKOnly, err.Source & ":" & err.Number
mUncompressedSnapFile = ""
ConvertReportToPDF = False
Resume EXIT_CREATESNAP

End Function



Private Function CurrentDBDir() As String
Dim strDBPath As String
Dim strDBFile As String
    strDBPath = CurrentDb.Name
    strDBFile = Dir(strDBPath)
    CurrentDBDir = Left$(strDBPath, Len(strDBPath) - Len(strDBFile))
End Function
'******************** Code End ****************



Private Function GetUniqueFilename(Optional path As String = "", _
Optional Prefix As String = "", _
Optional UseExtension As String = "") _
As String

' originally Posted by Terry Kreft
' to: comp.Databases.ms -Access
' Subject:  Re: Creating Unique filename ??? (Dev code)
' Date: 01/15/2000
' Author: Terry Kreft <terry.kreft@mps.co.uk>

' SL Note: Input strings must be NULL terminated.
' Here it is done by the calling function.

  Dim wUnique As Long
  Dim lpTempFileName As String
  Dim lngRet As Long

  wUnique = 0
  If path = "" Then path = CurDir
  lpTempFileName = String(MaxPath, 0)
  lngRet = GetTempFileName(path, Prefix, _
                            wUnique, lpTempFileName)

  lpTempFileName = Left(lpTempFileName, _
                        InStr(lpTempFileName, Chr(0)) - 1)
  Call Kill(lpTempFileName)
  If Len(UseExtension) > 0 Then
    lpTempFileName = Left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
  End If
  GetUniqueFilename = lpTempFileName
End Function







