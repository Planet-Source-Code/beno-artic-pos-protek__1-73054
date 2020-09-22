VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmBackup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BackUp"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Preklici"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
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
      MICON           =   "frmBackup.frx":0000
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
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   2
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Baze"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmBackup.frx":001C
         Top             =   0
         Width           =   480
      End
   End
   Begin LVbuttons.LaVolpeButton cmdBackup 
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Naredi backup"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
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
      MICON           =   "frmBackup.frx":08E6
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
   Begin MSComctlLib.ProgressBar progStat 
      Height          =   420
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblCBK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delam..."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim m_AutoBackup As Boolean
Public WithEvents clsBKU As clsHuffman
Attribute clsBKU.VB_VarHelpID = -1

Public Function ShowForm(Optional ByVal bAutoBackup As Boolean = False)
    
    m_AutoBackup = bAutoBackup
    
    
    If bAutoBackup = False Then
        Me.Show vbModal
    Else
        cmdBackup.Enabled = False
        BackUpDB
    End If
    
End Function

Private Sub cmdBackup_Click()
cmdBackup.Enabled = False
    BackUpDB
End Sub


Private Sub clsBKU_EncodeFinish()
    
    progStat.Visible = False
    lblCBK.Visible = False
    DoEvents
    
    If m_AutoBackup = False Then
        MsgBox "Backup je uspešno narejen.", vbInformation
    End If
    
    'close this form
    Unload Me
    
End Sub

Private Sub clsBKU_Progress(Procent As Integer)

    progStat.Value = Procent

End Sub
Private Sub Form_Activate()
    If m_AutoBackup = False Then
        cmdBackup.Enabled = True
    End If
End Sub

Private Sub LaVolpeButton1_click()
 
End Sub

Private Sub LaVolpeButton2_Click()
 Unload Me
End Sub
Private Sub BackUpDB()

    Dim FSO As New FileSystemObject
    
    Dim sDBFN As String
    Dim sDBTmpFN As String
    
    If FSO.FolderExists(App.path & "\Backup") = False Then
        FSO.CreateFolder App.path & "\Backup"
    End If
    
    'set backup file path filename
    sDBFN = App.path & "\Backup\" & Format$(Date, "yyyymmdd") & ".bak"
    
    'set temporary file
    sDBTmpFN = sDBFN
    '& Now - DateValue(Now)
    
   ' If FSO.FileExists(sDBTmpFN) = True Then
   '     FSO.DeleteFile sDBTmpFN
   ' End If
    
    'show ctl
    progStat.Visible = True
    lblCBK.Visible = True
    DoEvents
    
    'start backup
    Set frmBackup.clsBKU = New clsHuffman
    frmBackup.clsBKU.EncodeFile DBPathFileName, sDBTmpFN
   ' MsgBox (DBPathFileName)
    'frmBackup.cl
    'rename file
  '  If FSO.FileExists(sDBFN) = True Then
  '      FSO.DeleteFile sDBFN
  '  End If
    'MsgBox (sDBTmpFN & "     " & sDBFN)
   ' FSO.MoveFile sDBTmpFN, sDBFN
    
    
    Set FSO = Nothing
    progStat.Visible = False
   lblCBK.Visible = False
End Sub

