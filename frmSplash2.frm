VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash2.frx":000C
   ScaleHeight     =   4590
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2160
      Top             =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed to Lipa Solid Auto Supply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "frmSplash2.frx":220CA
      Top             =   0
      Width           =   7560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Me.Enabled = False
    MakeTopMost Me.hWnd
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
frmMAIN.Show
Me.Enabled = True
Unload Me
End Sub

