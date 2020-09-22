VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form slike 
   Caption         =   "Slika"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Shrani sliko"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Text            =   "c:\xampp\htdocs\pice"
      Top             =   6840
      Width           =   4095
   End
   Begin VB.PictureBox picLoad4 
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picLoad3 
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picLoad2 
      Height          =   495
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCrop4 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   9840
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox picCrop3 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   6720
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   11
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox picCrop2 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   9840
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   10
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdSaveCrop 
      Caption         =   "Shrani porezane slike"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10800
      TabIndex        =   8
      Top             =   7320
      Width           =   2175
   End
   Begin VB.PictureBox picLoad 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCrop 
      Caption         =   "Poreži sliko"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOpenPicture 
      Caption         =   "Odpri sliko"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCrop 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   6720
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.PictureBox picResize 
      Height          =   6135
      Left            =   240
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   6075
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      Begin VB.Shape shp4 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   3015
         Left            =   3120
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Shape shp3 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   3015
         Left            =   120
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Shape shp2 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   3015
         Left            =   3120
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sliko lahko potegneš notri iz windowsa"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   5640
         Width           =   2730
      End
      Begin VB.Shape shp1 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   3015
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "benc"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   8280
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porezana slika za picerije"
      Height          =   195
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "glavna slika"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "slike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private XX As Long
Private YY As Long
Private XX2 As Long
Private tempslika As String
Private YY2 As Long
Private isBoxExist As Boolean
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
Private Type PicBmp
    PicSize As Long
    PicType As Long
    PichBmp As Long
    PichPal As Long
    PicReserved As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Const RASTERCAPS As Long = 38
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC&, ByVal iCapabilitiy&) As Long
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC&, ByVal HPALETTE&, ByVal bForceBackground&) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal dwRop&) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC&) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdCrop_Click()
    On Error Resume Next
    Dim intR As Integer
    
    ' after opening the picture, click crop to get the coordinates and put these in a new picture
    If isBoxExist = False Then
        'MsgBox "The section to crop has not been selected yet, please select the area to crop first.", , "Picture Crop Error"
        err.clear
        Exit Sub
    End If
    Set picCrop.Picture = Nothing
    picCrop.Height = shp1.Height
    picCrop.Width = shp1.Width
     picCrop2.Height = shp2.Height
    picCrop2.Width = shp2.Width
     picCrop3.Height = shp3.Height
    picCrop3.Width = shp3.Width
     picCrop4.Height = shp4.Height
    picCrop4.Width = shp4.Width
    picCrop.Refresh
    picCrop2.Refresh
    picCrop3.Refresh
    picCrop4.Refresh
    Set picLoad.Picture = LoadPicture(picResize.Tag)
    PictureCopy picResize, picLoad
    
    picCrop.PaintPicture picResize.Image, 0, 0, shp1.Width, shp1.Height, shp1.Left, shp1.Top, shp1.Width, shp1.Height, vbSrcCopy
    picCrop2.PaintPicture picResize.Image, 0, 0, shp2.Width, shp2.Height, shp2.Left, shp2.Top, shp2.Width, shp2.Height, vbSrcCopy
    picCrop3.PaintPicture picResize.Image, 0, 0, shp3.Width, shp3.Height, shp3.Left, shp3.Top, shp3.Width, shp3.Height, vbSrcCopy
    picCrop4.PaintPicture picResize.Image, 0, 0, shp4.Width, shp4.Height, shp4.Left, shp4.Top, shp4.Width, shp4.Height, vbSrcCopy
    DoEvents
    cmdSaveCrop.Enabled = True
    
    'intR = MsgBox("Do you want to reload the cropped image?", vbYesNo + vbQuestion, "Confirm Reload")
    'If intR = vbNo Then Exit Sub
    
    SavePictureAPI picCrop, App.path & "\cropped.jpg"
    SavePictureAPI picCrop2, App.path & "\cropped2.jpg"
    SavePictureAPI picCrop3, App.path & "\cropped3.jpg"
    SavePictureAPI picCrop4, App.path & "\cropped4.jpg"
    Set picLoad.Picture = LoadPicture(App.path & "\cropped.jpg")
    Set picLoad2.Picture = LoadPicture(App.path & "\cropped2.jpg")
    Set picLoad3.Picture = LoadPicture(App.path & "\cropped3.jpg")
    Set picLoad4.Picture = LoadPicture(App.path & "\cropped4.jpg")
    PictureCopy picResize, picLoad
    PictureCopy picResize, picLoad2
    PictureCopy picResize, picLoad3
    PictureCopy picResize, picLoad4
    isBoxExist = False
    'picResize.Tag = App.path & "\cropped.jpg"
    Set picLoad.Picture = LoadPicture(tempslika)
    'Set picCrop.Picture = Nothing
    PictureCopy picResize, picLoad
        picResize.Refresh
    err.clear
End Sub
Private Sub cmdOpenPicture_Click()
    On Error Resume Next
    ' open dialog box and choose picture file to crop
    Dim strFileName As String
    strFileName = DialogOpen(CommonDialog1, "Odpri sliko", , "*.jpg")
    If Len(strFileName) = 0 Then Exit Sub
    Set picLoad.Picture = LoadPicture(strFileName)
    Set picCrop.Picture = Nothing
    PictureCopy picResize, picLoad
    isBoxExist = False
    cmdSaveCrop.Enabled = False
    cmdCrop.Enabled = True
    picResize.Tag = strFileName
        tempslika = strFileName
    err.clear
End Sub
Private Function DialogOpen(CD As CommonDialog, Optional ByVal Title As String = "Open Existing File", Optional ByVal InitDir As String = "...", Optional ByVal DefaultExt As String = "*.*") As String
    On Error GoTo ErrHandler
    ' returns the filename selected in an open dialog box
    Dim filName As String
    Dim filCnt As Integer
    Dim spFilter() As String
    Dim strFilter As String
    strFilter = File_Filters
    filCnt = 0
    StrParse spFilter, strFilter, "|"
    filCnt = ArraySearch(spFilter, DefaultExt)
    If filCnt <> 0 Then
        filCnt = filCnt / 2
    End If
    filName = vbNullString
    With CD
        .CancelError = True
        '.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        .Filter = strFilter
        .DialogTitle = Title
        .InitDir = InitDir
        .FileName = vbNullString
        .DefaultExt = DefaultExt
        .FilterIndex = filCnt
        .ShowOpen
        filName = .FileName
        If Len(filName) = 0 Then
            err.clear
            Exit Function
        End If
    End With
    DialogOpen = filName
    err.clear
    Exit Function
ErrHandler:
    err.clear
    Exit Function
    err.clear
End Function
Public Function File_Filters() As String
    On Error Resume Next
    ' this stores file extensions and their names
    Dim s_result As String
    s_result = "All Files (*.*)|*.*|Template File (*.tem)|*.tem"
    s_result = s_result & "|Temporal File (*.tmp)|*.tmp|Transaction File (*.trn)|*.trn"
    s_result = s_result & "|Data File (*.dat)|*.dat|Settings File (*.ini)|*.ini"
    s_result = s_result & "|Wave File (*.wav)|*.wav|Mpeg 3 File (*.mp3)|*.mp3"
    s_result = s_result & "|Help File Creator (*.hfc)|*.hfc|Bible File (*.bib)|*.bib"
    s_result = s_result & "|Dictionary File (*.dic)|*.dic|Topic Note File (*.top)|*.top"
    s_result = s_result & "|Study Node File (*.stu)|*.stu|Commentary File (*.com)|*.com"
    s_result = s_result & "|Graphics File (*.gra)|*.gra|Audio File (*.aud)|*.aud"
    s_result = s_result & "|Comma-Separated Values (*.csv)|*.csv|Sequential Access (*.seq)|*.seq"
    s_result = s_result & "|Excel (*.xls)|*.xls|Lotus 123 (*.wks)|*.wks"
    s_result = s_result & "|Rich Text Format (*.rtf)|*.rtf|Text (*.txt)|*.txt"
    s_result = s_result & "|Word for Windows (*.doc)|*.doc|Microsoft Access (*.mdb)|*.mdb"
    s_result = s_result & "|Adobe Acrobat (*.pdf)|*.pdf|Project Show (*.ps)|*.ps"
    s_result = s_result & "|Tree File (*.tree)|*.tree|Dictionary File (*.dict)|*.dict"
    s_result = s_result & "|Visual Basic Project File (*.vbp)|*.vbp|Visual Basic Project Group File (*.vbg)|*.vbg"
    s_result = s_result & "|Visual Basic Mak File (*.mak)|*.mak|Visual Basic Form File(*.frm)|*.frm"
    s_result = s_result & "|Visual Basic Module File (*.mod)|*.mod|Visual Basic Class Module (*.cls)|*.cls"
    s_result = s_result & "|Bitmap File (*.bmp)|*.bmp|Tif File (*.tif)|*.tif"
    s_result = s_result & "|Tiff File (*.tif)|*.tif|Jpeg File (*.jpg)|*.jpg"
    s_result = s_result & "|DBase File (*.dbf)|*.dbf|Ms Project File (*.mpp)|*.mpp"
    s_result = s_result & "|Gif File (*.gif)|*.gif|Png File (*.png)|*.png"
    s_result = s_result & "|Batch File (*.bat)|*.bat|Executable File (*.exe)|*.exe"
    s_result = s_result & "|Icon File (*.ico|*.ico|Configuration (*.cfg)|*.cfg"
    s_result = s_result & "|Visual Basic Setup File (*.lst)|*.lst|Inno Setup Script (*.iss)|*.iss"
    s_result = s_result & "|Spell Check Log (*.spl)|*.spl|Document Tracking System (*.dts)|*.dts"
    s_result = s_result & "|Archive File (*.arc)|*.arc|List View File (*.lvf)|*.lvf"
    s_result = s_result & "|Executable File (*.exe)|*.exe|Dynamic Link Library File (*.dll)|*.dll"
    File_Filters = s_result
    err.clear
End Function
Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delim As String, Optional PreserveSize As Long = -1) As Long
    On Error Resume Next
    ' this is the same as the split function, however starts at 1 and returns the number
    ' of delimiters counted
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    ReDim retarray(VarE + 1)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
    Next
    If PreserveSize <> -1 Then ReDim Preserve retarray(PreserveSize)
    StrParse = UBound(retarray)
    err.clear
End Function
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ' search an array and return the index position
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
    err.clear
End Function
Private Sub PictureCopy(NewPicBox As Variant, ByVal ActualPic As StdPicture)
    On Error Resume Next
    ' this copies a picturebox frome one to another and resizes the picture to fit the height and width of new picture box
    Dim NewH As Long ' new height
    Dim NewW As Long 'New Width
    NewH = NewPicBox.Height 'actual image height
    NewW = NewPicBox.Width 'actual image width
    With NewPicBox
        .AutoRedraw = True 'set needed properties
        .Cls 'clear picture box
        .PaintPicture ActualPic, 0, 0, NewW, NewH        'paint new picture size in picturebox
    End With
    err.clear
End Sub
Private Sub cmdSaveCrop_Click()
    On Error Resume Next
    ' save the picture, the vb savepicture works if the source picturebox is linked to a file
    ' the current cropped image is not sourced from a file, thus use the api to save the picture
    If DirektorijExists(Text1.Text) = True Then Else MkDir (Text1.Text)
    SavePictureAPI picCrop, Text1.Text & "\kot1_" & trenslika & ".jpg"
       SavePictureAPI picCrop2, Text1.Text & "\kot3_" & trenslika & ".jpg"
          SavePictureAPI picCrop3, Text1.Text & "\kot2_" & trenslika & ".jpg"
             SavePictureAPI picCrop4, Text1.Text & "\kot4_" & trenslika & ".jpg"
   ' picResize.Tag = App.path & "\cropped.jpg"
    ' open the saved file with associated default program
   ' ShellExecute Me.hwnd, "open", App.Path & "\cropped.jpg", "", "", 1
    err.clear
End Sub

Private Sub Command1_Click()
  If DirektorijExists(App.path & "\slike") = True Then Else MkDir (App.path & "\slike")
  If DirektorijExists(Text1.Text) = True Then Else MkDir (Text1.Text)
    SavePictureAPI picResize, App.path & "\slike\ma_" & trenslika & ".jpg"
    SavePictureAPI picResize, Trim(Text1.Text) & "\ma_" & trenslika & ".jpg"
frmProdEntry.slikka.Picture = LoadPicture(App.path & "\slike\ma_" & trenslika & ".jpg")
   Unload Me
End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        XX = X
        YY = y
        XX2 = X
        YY2 = y
        shp1.shape = 0
        shp1.Visible = True
       ' shp1.Left = X
       ' shp1.Top = y
       ' shp1.Width = 0
       ' shp1.Height = 0
    End If
    err.clear
End Sub
Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    'Drawing the rectangle
    If Button = 1 Then
        If isBoxExist = True Then
            isBoxExist = False
            PictureCopy picResize, picLoad
        End If
        XX2 = X
        YY2 = y
       ' shp1.Left = IIf(X > XX, XX, X)
       ' shp1.Top = IIf(y > YY, YY, y)
       ' shp1.Width = Abs(X - XX)
       ' shp1.Height = Abs(y - YY)
    End If
    err.clear
End Sub
Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    ' the user has finished dragging the crop area
    If Button = 1 Then
      '  picResize.Line (XX, YY)-(XX2, YY2), &H0, B
        shp1.Visible = False
        isBoxExist = True
    End If
    err.clear
End Sub
Private Sub picResize_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    ' if a file is dropped here, load it
    Set picLoad.Picture = LoadPicture(Data.Files(1))
    PictureCopy picResize, picLoad
    cmdCrop.Enabled = True
    picResize.Tag = Data.Files(1)
    err.clear
End Sub
Private Sub SavePictureAPI(picHwnd As PictureBox, ByVal strFileName As String)
    On Error Resume Next
    Dim myX As Picture
    Dim RectActive As RECT
    Dim R As Long
    R = GetWindowRect(picHwnd.hwnd, RectActive)
    Set myX = CaptureWindow(picHwnd.hwnd, False, 1, 1, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
    SavePicture myX, strFileName
    err.clear
End Sub
Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error Resume Next
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim R As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
    ' Depending on the value of Client get the proper device context.
    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
        ' window.
    End If
    ' Create a memory device context for the copy process.
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC.
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    ' Get screen properties.
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    ' capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
    ' support.
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    ' palette.
    ' If the screen has a palette make a copy and realize it.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        ' Create a copy of the system palette.
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If
    ' Copy the on-screen image into the memory DC.
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    ' If the screen has a palette get back the palette that was
    ' selected in previously.
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    ' Release the device context resources back to the system.
    R = DeleteDC(hDCMemory)
    R = ReleaseDC(hWndSrc, hDCSrc)
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
    err.clear
End Function
Private Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    On Error Resume Next
    Dim Pic As PicBmp
    ' IPicture requires a reference to "Standard OLE Types."
    Dim iPic As IPicture
    Dim IID_IDispatch As GUID
    Dim R As Long
    ' Fill in with IDispatch Interface ID.
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    ' Fill Pic with necessary parts.
    With Pic
        .PicSize = Len(Pic)          ' Length of structure.
        .PicType = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .PichBmp = hBmp              ' Handle to bitmap.
        .PichPal = hPal              ' Handle to palette (may be null).
    End With
    ' Create Picture object.
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, iPic)
    ' Return the new Picture object.
    Set CreateBitmapPicture = iPic
    err.clear
End Function
Private Function DirektorijExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirektorijExists = fs.FolderExists(OrigFile)
End Function
