VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "View Detail of:"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4935
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WBrow 
      CausesValidation=   0   'False
      Height          =   3615
      Left            =   -15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "frmViewDetail.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPrint_Click()
    Me.WindowState = 2
    WBrow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, 100, 200
    Me.WindowState = 0
End Sub

Private Sub Form_Load()
    frmView.WBrow.Navigate "about:blank"
End Sub

Private Sub WBrow_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    MsgBox IsChildWindow
End Sub
