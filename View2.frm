VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Predogled"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12405
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser Wbrow 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   12375
      ExtentX         =   21828
      ExtentY         =   15690
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
   Begin VB.CommandButton Command2 
      Caption         =   "Shrani report"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Predogled"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Printaj"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "View2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12375
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
 Wbrow.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command1_Click()
  Me.WindowState = 2
  
      Wbrow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER, 100, 200
    
      Me.WindowState = 0
End Sub

Private Sub Command2_Click()
    Wbrow.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub Command3_Click()
    Wbrow.ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DODEFAULT, -1
End Sub

