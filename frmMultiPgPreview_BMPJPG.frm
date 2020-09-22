VERSION 5.00
Begin VB.Form frmMultiPgPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   1530
   ClientTop       =   1725
   ClientWidth     =   4920
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMultiPgPreview_BMPJPG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGoto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   3090
      TabIndex        =   28
      Top             =   5550
      Visible         =   0   'False
      Width           =   3120
      Begin VB.CommandButton cmdGotoOK 
         Caption         =   "Ok"
         Height          =   255
         Left            =   2100
         TabIndex        =   30
         Top             =   495
         Width           =   780
      End
      Begin VB.TextBox txtGoto 
         Height          =   315
         Left            =   1305
         TabIndex        =   29
         Text            =   "1"
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pojdi na stran#"
         ForeColor       =   &H80000014&
         Height          =   465
         Left            =   120
         TabIndex        =   31
         Top             =   165
         Width           =   1080
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000F&
         Height          =   750
         Left            =   15
         Top             =   15
         Width           =   3045
      End
   End
   Begin VB.PictureBox picFullPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3840
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   18
      Top             =   2340
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picPrintPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3885
      ScaleHeight     =   435
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6450
      Left            =   4365
      ScaleHeight     =   6450
      ScaleWidth      =   555
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   555
      Begin VB.CheckBox cmdFullPage 
         Caption         =   "Ured"
         Height          =   510
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Fit Page"
         Top             =   1215
         Width           =   525
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "&Pejt"
         Height          =   510
         Left            =   45
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Goto Page"
         Top             =   2955
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   285
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":0A0E
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Next Page"
         Top             =   2580
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   45
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":0D89
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Prev. Page"
         Top             =   2580
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1260
         LargeChange     =   10
         Left            =   105
         Max             =   100
         Min             =   -20
         TabIndex        =   32
         Top             =   3495
         Width           =   330
      End
      Begin VB.CommandButton cmd_quit 
         Cancel          =   -1  'True
         Caption         =   "Izhod"
         Height          =   630
         Left            =   30
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1129
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   525
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   585
         Left            =   30
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1813
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Send to Printer"
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "P 1"
         Height          =   600
         Left            =   45
         TabIndex        =   37
         Top             =   1830
         Width           =   465
      End
   End
   Begin VB.PictureBox picHScroll 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   4920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6450
      Visible         =   0   'False
      Width           =   4920
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   0
         Max             =   100
         TabIndex        =   3
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.PictureBox picPrintOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H000000FF&
      Height          =   2640
      Left            =   555
      ScaleHeight     =   2610
      ScaleWidth      =   3150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   3180
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1695
         TabIndex        =   7
         Text            =   "1"
         Top             =   1350
         Width           =   420
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2475
         TabIndex        =   8
         Text            =   "1"
         Top             =   1350
         Width           =   420
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Ok"
         Height          =   360
         Left            =   2145
         TabIndex        =   10
         Top             =   2070
         Width           =   705
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopiraj vse strani v mapo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   4
         Left            =   585
         TabIndex        =   19
         Top             =   420
         Width           =   2250
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   4
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1C0C
         Top             =   390
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1CA9
         Top             =   705
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopiraj vse strani v clipb."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   0
         Left            =   585
         TabIndex        =   4
         Top             =   735
         Width           =   2250
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Natisni trenutno stran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   1
         Left            =   585
         TabIndex        =   5
         Top             =   1065
         Width           =   1965
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1D46
         Top             =   1035
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   2
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1DE3
         Top             =   1335
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   3
         Left            =   270
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1E80
         Top             =   1665
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Natisni vse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   3
         Left            =   585
         TabIndex        =   9
         Top             =   1695
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "do"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2175
         TabIndex        =   16
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label lblPrintingPg 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   255
         TabIndex        =   15
         Top             =   2250
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nastavitve printanja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Left            =   135
         TabIndex        =   13
         Top             =   30
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000F&
         Height          =   2535
         Left            =   30
         Top             =   30
         Width           =   3090
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Natisni stran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Index           =   2
         Left            =   585
         TabIndex        =   6
         Top             =   1365
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3765
   End
   Begin VB.PictureBox picGetFolder 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4440
      Left            =   1245
      ScaleHeight     =   4410
      ScaleWidth      =   6375
      TabIndex        =   20
      Top             =   615
      Visible         =   0   'False
      Width           =   6405
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1530
         TabIndex        =   26
         Top             =   45
         Width           =   3930
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   30
         TabIndex        =   25
         Top             =   450
         Width           =   6315
      End
      Begin VB.CommandButton cmdNewFolder 
         Height          =   345
         Left            =   5955
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":1F1D
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "New Folder"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUpOne 
         Height          =   345
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMultiPgPreview_BMPJPG.frx":226B
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Back Up"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4830
         TabIndex        =   22
         Top             =   3975
         Width           =   1470
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3255
         TabIndex        =   21
         Top             =   3975
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " Select a Directory: "
         Height          =   195
         Left            =   75
         TabIndex        =   27
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.Image imgFit 
      Height          =   240
      Index           =   1
      Left            =   360
      Picture         =   "frmMultiPgPreview_BMPJPG.frx":251D
      Top             =   5235
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFit 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmMultiPgPreview_BMPJPG.frx":2AA7
      Top             =   5220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   0
      Picture         =   "frmMultiPgPreview_BMPJPG.frx":3031
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   555
      Picture         =   "frmMultiPgPreview_BMPJPG.frx":30DE
      Top             =   4875
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1999-2003
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
'/* Used for Manifest files (Win XP)
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long


'/* Copy BMP to JPG
'Private Declare Function DIBmpToJpg Lib "DIjpg.dll" (ByVal SrcPath As String, ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long

Public PageNumber As Integer
Private ViewPage As Integer
Private TempDir As String
Private OptionV As Integer
Private FitToPage As Boolean

Private Type PanState
   x As Long
   Y As Long
End Type
Dim PanSet As PanState

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long
    
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type
Private UseStretchBit As Boolean

Private Sub cmd_print_Click()
    txtTo.Text = PageNumber + 1
    OptionV = 3
    Call optText_Click(OptionV)
    picGoto.Visible = False
    picPrintOptions.Left = Me.Width - (Picture2.Width + picPrintOptions.Width + 50)
    picGetFolder.Left = Me.Width - (Picture2.Width + picGetFolder.Width + 50)
    picPrintOptions.Visible = True
End Sub


Private Sub cmd_quit_Click()
'Dim cPrint As clsMultiPgPreview
    cPrint.SendToPrinter = False
    Unload Me
End Sub

Private Sub cmdFullPage_Click()
  Dim xmin As Single
  Dim ymin As Single
  Dim wid As Single
  Dim hgt As Single
  Dim aspect As Single
 
    '/* If already here then restore original
    If cmdFullPage.Value = 0 Then
        Picture1.Visible = True
        Picture1.SetFocus
        picFullPage.Visible = False
        cmdFullPage.Picture = imgFit(0).Picture
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    cmdFullPage.Picture = imgFit(1).Picture
    
    '/* Clear any picture and set the size and loaction
    Set picFullPage.Picture = Nothing
    If Not picHScroll.Visible Then
        picFullPage.Height = Me.Height - 100
        picFullPage.Width = picFullPage.Height * 0.773
        picFullPage.Move ((Me.Width - Picture2.Width) - picFullPage.Width) \ 2, 0
    Else
        picFullPage.Top = 50
        picFullPage.Left = 50
        picFullPage.Width = Me.Width - Picture2.Width - 100
        picFullPage.Height = picFullPage.Width * 0.773
    End If
        
    '/* Get the scale values
    aspect = Picture1.ScaleHeight / Picture1.ScaleWidth
    wid = picFullPage.ScaleWidth
    hgt = picFullPage.ScaleHeight
    
    '/* MaintainRatio
    If hgt / wid > aspect Then
        hgt = aspect * wid
        xmin = picFullPage.ScaleLeft
        ymin = (picFullPage.ScaleHeight - hgt) / 2
    Else
        wid = hgt / aspect
        xmin = (picFullPage.ScaleWidth - wid) / 2
        ymin = picFullPage.ScaleTop
    End If
    
    If UseStretchBit Then '/* NT platform
        StretchBlt picFullPage.hdc, _
            xmin, ymin, wid, hgt, _
            Picture1.hdc, _
            0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
    Else
        picFullPage.PaintPicture Picture1.Picture, _
          xmin, ymin, wid, hgt, _
          0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
    End If

    picGoto.Visible = False
    Picture1.Visible = False
    picFullPage.Visible = True
    picFullPage.SetFocus
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdGoTo_Click()
    picGoto.Top = cmdGoTo.Top
    picGoto.Left = Me.Width - (Picture2.Width + picGoto.Width + 50)
    picGoto.Visible = True
    picGoto.ZOrder
    txtGoto = CStr(ViewPage + 1)
    txtGoto.SelStart = 0
    txtGoto.SelLength = Len(txtGoto)
    txtGoto.SetFocus
End Sub

Private Sub cmdGotoOK_Click()
  Dim NewPageNo As Integer
    
    On Local Error Resume Next
    
    txtGoto.SetFocus
    NewPageNo = Val(txtGoto)
    If NewPageNo = 0 Then Exit Sub
    
    NewPageNo = NewPageNo - 1
    If NewPageNo > PageNumber Then NewPageNo = PageNumber
    ViewPage = NewPageNo
        
    Picture1.SetFocus
    Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    
    picPrintOptions.Visible = False
    picGetFolder.Visible = False
    picGoto.Visible = False
    
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages

End Sub

Private Sub cmdPrint_Click()
  Dim i As Integer
  Dim cPrint As clsMultiPgPreview
    '/* Prevent printing again until done
    picPrintOptions.Enabled = False
    lblPrintingPg.Visible = True
    cmdPrint.Visible = False
    
    Select Case OptionV
    Case 0 '/* Copy to clipboard
        Clipboard.clear
        Clipboard.SetData Picture1.Picture, vbCFBitmap
    Case 1 '/* Print current page
        lblPrintingPg.Caption = "Printing page " & ViewPage + 1
        lblPrintingPg.Refresh
        Call PrintPictureBox(Picture1, True, False)
    Case 2 '/* Print range
        For i = Val(txtFrom) - 1 To Val(txtTo) - 1
            lblPrintingPg.Caption = "Printing page " & CStr(i + 1) & " of " & txtTo
            lblPrintingPg.Refresh
            Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(i) & ".bmp")
            Call PrintPictureBox(Picture1, True, False)
        Next i
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Case 4
        picGetFolder.Visible = True
        picGetFolder.ZOrder
    Case Else '/* Print all
       ' cPrint.SendToPrinter = True '/* Send to Printer */
        Unload Me
    End Select
    
    '/* Restore normal view
    picPrintOptions.Enabled = True
    cmdPrint.Visible = True
    picPrintOptions.Visible = False
    lblPrintingPg.Visible = False
    
End Sub

Private Sub Command1_Click(Index As Integer)
    On Local Error Resume Next
    If Index = 0 Then
        ViewPage = ViewPage - 1
        If ViewPage < 0 Then ViewPage = 0
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Else
        ViewPage = ViewPage + 1
        If ViewPage > PageNumber Then ViewPage = PageNumber
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    End If
    
    Picture1.Top = 0
    picPrintOptions.Visible = False
    picGoto.Visible = False
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Call DisplayPages
    If Picture1.Width < Me.Width - Picture2.Width Then
        Picture1.Move ((Me.Width - Picture2.Width) - Picture1.Width) \ 2, 0
    End If
    cmdFullPage.Picture = imgFit(0).Picture
    Label5 = "Goto Page#" & vbCrLf & "(1 to " & CStr(PageNumber + 1) & ")"
    Picture1.SetFocus
End Sub

Private Sub Form_Click()
    picPrintOptions.Visible = False
    picGetFolder.Visible = False
    picGoto.Visible = False
End Sub


Private Sub Form_Initialize()
    '/* Used for Manifest files (Win XP)
    Call InitCommonControls
    'MakeXPButton cmd_quit
    'MakeXPButton cmd_print
    'MakeXPButton cmdFullPage
    'MakeXPButton cmdGoTo
    'MakeXPButton Command1(0)
    'MakeXPButton Command1(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 71 Or KeyAscii = 103 Then cmdGoTo_Click
End Sub

Private Sub Form_Load()
  Dim osv As OSVERSIONINFO
  Const VER_PLATFORM_WIN32_NT = 2
    
    osv.OSVSize = Len(osv)
    If GetVersionEx(osv) = 1 Then
        If osv.PlatformID = VER_PLATFORM_WIN32_NT Then
            UseStretchBit = True
        Else
            UseStretchBit = False
        End If
    End If
    
    Me.Move 0, 0, Screen.Width, Screen.Height
    Picture1.Move 0, 0

    VScroll1.Height = Me.Height - cmdGoTo.Top - cmdGoTo.Height - 500
    HScroll1.Width = Me.Width - Picture2.Width - 500
    
    TempDir = Environ("TEMP") & "\"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim tFilename As String
    
    '/* Remove preview pages
    tFilename = Dir(TempDir & "PPview*.bmp")
    If tFilename > vbNullString Then
        Do
            Kill TempDir & tFilename
            tFilename = Dir(TempDir & "PPview*.bmp")
        Loop Until tFilename = vbNullString
    End If
    
    PageNumber = 0
    ViewPage = 0
    Set frmMultiPgPreview = Nothing
End Sub

Private Sub HScroll1_Change()
    On Local Error Resume Next
    Picture1.Left = -(HScroll1.Value)
    'HScroll1.SetFocus
    Picture1.SetFocus
    On Local Error GoTo 0
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 33 '/* PageUp
        Call Command1_Click(0)
    Case 34 '/* PageDown
        Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        picGetFolder.Visible = False
        picGoto.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select

End Sub

Private Sub optPrint_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 4
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub

Private Sub optText_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 4
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub

Private Sub picFullPage_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Decode_KeyUp(KeyCode, Shift)
End Sub
Private Sub Decode_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 37 '/* Arrow left
        If HScroll1.Visible = False Then
            Call Command1_Click(0)
        Else
            HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    Case 39 '/* Arrow right
        If HScroll1.Visible = False Then
            Call Command1_Click(1)
        Else
            HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    Case 33 '/* PageUp
        Call Command1_Click(0)
    Case 34 '/* PageDown
        Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        picGetFolder.Visible = False
        picGoto.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select
End Sub

Private Sub cmdNewFolder_Click()
  Dim NewFolderName As String
  Dim Security As SECURITY_ATTRIBUTES
  
    NewFolderName = InputBox("Enter Folder Name", , "New Folder")
    NewFolderName = Trim(NewFolderName)
    If NewFolderName > vbNullString Then
        CreateDirectory Dir1.path & "\" & NewFolderName, Security
        NewFolderName = Dir1.path & "\" & NewFolderName
        Dir1.Refresh
        Dir1.path = NewFolderName
    End If
        
End Sub

Private Sub cmdOpen_Click()
  Dim FolderName As String
  Dim ReportTitle As String
    Dim cPrint As clsMultiPgPreview
  Dim i As Integer
  
    FolderName = Dir1.path & "\"
    picGetFolder.Visible = False
    
    picPrintOptions.Visible = True
    picPrintOptions.Enabled = False
    lblPrintingPg.Visible = True
    cmdPrint.Visible = False
    
    On Local Error GoTo CopyError:

    DoEvents
    ReportTitle = Trim(cPrint.ReportTitle)
    If ReportTitle = vbNullString Or InStr(ReportTitle, "\") Then
        ReportTitle = "PPview"
    End If
    
    For i = 0 To PageNumber
       lblPrintingPg.Caption = "Copying page " & i + 1
       lblPrintingPg.Refresh
       'DIBmpToJpg TempDir & "PPview" & CStr(i) & ".bmp", FolderName & ReportTitle & CStr(i + 1) & ".jpg", 100, 1
    Next
        
    '/* Restore normal view
    picPrintOptions.Enabled = True
    cmdPrint.Visible = True
    picPrintOptions.Visible = False
    lblPrintingPg.Visible = False

Exit Sub

CopyError:
    If err.Number = 76 Then
        ReportTitle = "PPview"
        Resume
    End If
End Sub

Private Sub cmdQuit_Click()
    picGetFolder.Visible = False
    '/* Restore normal view
    picPrintOptions.Enabled = True
    cmdPrint.Visible = True
    picPrintOptions.Visible = False
    lblPrintingPg.Visible = False
End Sub

Private Sub cmdUpOne_Click()
    Dir1.path = Dir1.List(-2)
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dir1.path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub Picture1_Click()
    picPrintOptions.Visible = False
    picGetFolder.Visible = False
    picGoto.Visible = False
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Decode_KeyUp(KeyCode, Shift)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.x = x
      PanSet.Y = Y
      MousePointer = vbSizePointer
   End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim nTop As Integer, nLeft As Integer

   On Local Error Resume Next

   If Button = vbLeftButton And Shift = 0 Then

      '/* new coordinates?
      With Picture1
         nTop = -(.Top + (Y - PanSet.Y))
         nLeft = -(.Left + (x - PanSet.x))
      End With

      '/* Check limits
      With VScroll1
         If .Visible Then
            If nTop < .min Then
               nTop = .min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -Picture1.Top
         End If
      End With

      With HScroll1
         If .Visible Then
            If nLeft < .min Then
               nLeft = .min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -Picture1.Left
         End If
      End With

      Picture1.Move -nLeft, -nTop

   End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      If VScroll1.Visible Then VScroll1.Value = -(Picture1.Top)
      If HScroll1.Visible Then HScroll1.Value = -(Picture1.Left)
   End If
   MousePointer = vbDefault
End Sub

Private Sub txtFrom_Change()
    If Val(txtFrom) < 1 Then txtFrom = 1
    If Val(txtFrom) > Val(txtTo) Then txtFrom = txtTo
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.SelStart = 0
    txtFrom.SelLength = Len(txtFrom)
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtFrom = txtFrom + 1
        KeyCode = False
    Case 40  '/* "-"
        txtFrom = txtFrom - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    IsNumber txtFrom, KeyAscii, False, False
End Sub

Private Sub txtGoto_Change()
    If Val(txtGoto) > PageNumber + 1 Then txtGoto = PageNumber + 1
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdGotoOK_Click
    ElseIf (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub


Private Sub txtTo_Change()
    If Val(txtTo) > PageNumber + 1 Then txtTo = PageNumber + 1
    If Val(txtTo) < Val(txtFrom) Then txtTo = txtFrom
End Sub

Private Sub txtTo_GotFocus()
    txtTo.SelStart = 0
    txtTo.SelLength = Len(txtTo)
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtTo = txtTo + 1
        KeyCode = False
    Case 40  '/* "-"
        txtTo = txtTo - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    IsNumber txtTo, KeyAscii, False, False
End Sub

Private Sub VScroll1_Change()
    On Local Error Resume Next
    Picture1.Top = -(VScroll1.Value)
    'VScroll1.SetFocus
    Picture1.SetFocus
    On Local Error GoTo 0
End Sub

Private Sub DisplayPages()
    Label1 = CStr(ViewPage + 1) & vbNewLine & "-- of --" & vbNewLine & CStr(PageNumber + 1)
    
    If Picture1.Width > Me.Width - Picture2.Width Then
        picHScroll.Visible = True
    Else
        picHScroll.Visible = False
    End If

    If Picture1.Height >= Me.Height Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    
    If picFullPage.Visible Then cmdFullPage_Click
    
End Sub
Private Sub PrintPictureBox(pBox As PictureBox, _
                           Optional ScaleToFit As Boolean = True, _
                           Optional MaintainRatio As Boolean = True)
   Dim cPrint As clsMultiPgPreview
 Dim xmin As Single
 Dim ymin As Single
 Dim wid As Single
 Dim hgt As Single
 Dim aspect As Single
 
    Screen.MousePointer = vbHourglass
    
    If Not ScaleToFit Then
        wid = Printer.ScaleX(pBox.ScaleWidth, pBox.ScaleMode, Printer.ScaleMode)
        hgt = Printer.ScaleY(pBox.ScaleHeight, pBox.ScaleMode, Printer.ScaleMode)
        xmin = (Printer.ScaleWidth - wid) / 2
        ymin = (Printer.ScaleHeight - hgt) / 2
    Else
        aspect = pBox.ScaleHeight / pBox.ScaleWidth
        wid = Printer.ScaleWidth
        hgt = Printer.ScaleHeight
        
        If MaintainRatio Then
            If hgt / wid > aspect Then
                hgt = aspect * wid
                xmin = Printer.ScaleLeft
                ymin = (Printer.ScaleHeight - hgt) / 2
            Else
                wid = hgt / aspect
                xmin = (Printer.ScaleWidth - wid) / 2
                ymin = Printer.ScaleTop
            End If
        End If
    End If
    
    Printer.PaintPicture pBox.Picture, xmin, ymin, wid, hgt, , , , , vbSrcCopy
    Printer.EndDoc
    
'    Printer.Orientation = cPrint.Orientation

    Screen.MousePointer = vbDefault

End Sub

Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 37, 33 '/* Arrow left, PageUp
        If HScroll1.Visible = False Then
            Call Command1_Click(0)
        Else
            HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    Case 39, 34 '/* Arrow right, PageDown
        If HScroll1.Visible = False Then
            Call Command1_Click(1)
        Else
            HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        picGetFolder.Visible = False
        picGoto.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select
End Sub

