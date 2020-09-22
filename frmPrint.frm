VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PREDOGLED"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   HelpContextID   =   1025
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmPrint.frx":038A
   ScaleHeight     =   9240
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton cmdSave 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   8640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "SHRANI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPrint.frx":3389B
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdPrint 
      Height          =   375
      Left            =   1840
      TabIndex        =   4
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "PRINTAJ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPrint.frx":338B7
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.PictureBox cmdUp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14880
      ScaleHeight     =   1035
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   135
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14760
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   8295
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   13095
         ExtentX         =   23098
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox cmdDown 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14880
      ScaleHeight     =   1035
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sh_word As String
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long


Private Sub cmdDown_Click()
Frame1.Enabled = True
WebBrowser1.SetFocus
Sendkeys "{PGDN}", True
Frame1.Enabled = False
cmdDown.SetFocus
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1025)
End Sub

Private Sub cmdPrevious_Click()
'frmReport.Visible = True
'Me.Hide
End Sub

Private Sub cmdPrinta_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim strFooter As String
Dim strHeader As String

'store HEADER & FOOTER
strFooter = QueryValue("Software\Microsoft\Internet Explorer\PageSetup", "footer")
strHeader = QueryValue("Software\Microsoft\Internet Explorer\PageSetup", "header")

'our HEADER & FOOTER
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "header", "", REG_SZ
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "footer", "Izdelano poroèilo z dne " & Format(Now, "dd-MM-yyyy"), REG_SZ

WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER

'replace HEADER & FOOTER with old value
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "footer", strFooter, REG_SZ
SetKeyValue "Software\Microsoft\Internet Explorer\PageSetup", "header", strHeader, REG_SZ

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub cmdSave_Click()
Dim strFileName As String
Dim boolOnce As Boolean
boolOnce = False
strFileName = Replace(LCase(WebBrowser1.LocationURL), "file:///", "")
strFileName = Replace(strFileName, "%20", " ")
strFileName = Replace(strFileName, "/", "\")

CD.CancelError = True
CD.DialogTitle = "Shrani report..."
CD.Filter = "MS Excel (*.xls)|*.xls|HTM (*.htm)|*.htm|MS Word (*.doc)|*.doc"

CD.FileName = Replace(GetVirtualFileName(strFileName), "." & chkFileExtension(strFileName), "")
'If Len(Trim(frmReport.txtReportTitle.text)) > 0 Then
'    CD.FileName = frmReport.txtReportTitle.text
'End If
On Error GoTo err1
CD.ShowSave

If Len(CD.FileName) > 0 Then
    If chkFilePath(CD.FileName) = False Then
        fso.CopyFile strFileName, CD.FileName
    ElseIf chkFilePath(CD.FileName) = True Then
        boolConfirm = MsgBox("Ta datoteka že obstaja prepišem? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
            On Error GoTo err1
            fso.CopyFile strFileName, CD.FileName, True
        Else
            Exit Sub
        End If
    End If
    MsgBox "File saved to " & CD.FileName, vbInformation
End If
Exit Sub
    
err1:
If boolOnce = False And UCase(err.Description) <> "Preklicano." Then
    CD.FileName = Replace(GetVirtualFileName(strFileName), "." & chkFileExtension(strFileName), "")
    boolOnce = True
    Resume
End If

MsgBox err.Description, vbExclamation
Exit Sub

End Sub

Private Sub cmdUp_Click()
Frame1.Enabled = True
WebBrowser1.SetFocus
Sendkeys "{PGUP}", True
Frame1.Enabled = False
cmdUp.SetFocus
End Sub


Private Sub Form_Load()
Me.WebBrowser1.Navigate App.path & "\tempx.htm"
End Sub

Private Sub LaVolpeButton1_Click()
Dim hwnd As Long
    hwnd = GetBrowserHandle(Me.hwnd)
    webhw = hwnd

PrintPreview WebBrowser1, 1, 1, 1, 1, 1
End Sub

Private Sub toword_Click()
Dim strFileName As String
Dim boolOnce As Boolean
WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, 0, 0
'MsgBox (strFileName)
sh_word = App.path & "\repor.doc"
'CD.CancelError = True
'CD.DialogTitle = "Shrani report..."
'CD.Filter = "MS Excel (*.xls)|*.xls|HTM (*.htm)|*.htm|MS Word (*.doc)|*.doc"

'CD.FileName = Replace(GetVirtualFileName(strFileName), "." & "doc", "")
'If Len(Trim(frmReport.txtReportTitle.text)) > 0 Then
'    CD.FileName = frmReport.txtReportTitle.text
'End If
On Error GoTo err1
'CD.ShowSave

'If Len(C) > 0 Then
'    If chkFilePath(CD.FileName) = False Then
        fso.CopyFile strFileName, sh_word
 '   ElseIf chkFilePath(CD.FileName) = True Then
  '      boolConfirm = MsgBox("Ta datoteka že obstaja prepišem? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
   '     If boolConfirm = vbYes Then
   '         On Error GoTo err1
   '         FSO.CopyFile strFileName, CD.FileName, True
   '     Else
   '         Exit Sub
   '     End If
   ' End If
   ' MsgBox "File saved to " & CD.FileName, vbInformation
'End If
Dim wrdApp As New WORD.Application
    wrdApp.Documents.Open App.path & "\tempx.doc"
    wrdApp.Visible = True
'Call Shell(sh_word, vbMaximizedFocus)
Exit Sub
    
err1:
If boolOnce = False And UCase(err.Description) <> "Preklicano." Then
    CD.FileName = Replace(GetVirtualFileName(strFileName), "." & chkFileExtension(strFileName), "")
    boolOnce = True
    Resume
End If

MsgBox err.Description, vbExclamation
Exit Sub

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress = 0 Then
'    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, Null, Null
'    Unload Me
End If
End Sub
