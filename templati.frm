VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{46E98F52-504C-4B1B-B951-CE2725A20438}#1.1#0"; "gdpicturepro4.ocx"
Begin VB.Form templati 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TEMPLATI"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtZoom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2745
      TabIndex        =   11
      Top             =   7440
      Width           =   690
   End
   Begin VB.CommandButton btZoomIn 
      Caption         =   "+"
      Height          =   285
      Left            =   3465
      TabIndex        =   10
      Top             =   7440
      Width           =   300
   End
   Begin VB.CommandButton btZoomOut 
      Caption         =   "-"
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   7440
      Width           =   300
   End
   Begin VB.CommandButton btlastpage 
      Caption         =   ">>"
      Height          =   255
      Left            =   4125
      TabIndex        =   7
      Top             =   7080
      Width           =   675
   End
   Begin VB.CommandButton btFirstpage 
      Caption         =   "<<"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   7080
      Width           =   675
   End
   Begin VB.CommandButton btNextPage 
      Caption         =   ">"
      Height          =   255
      Left            =   3390
      TabIndex        =   5
      Top             =   7080
      Width           =   675
   End
   Begin VB.CommandButton btPreviousPage 
      Caption         =   "<"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   7080
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   480
   End
   Begin GdPicturePro4.GdViewer GdViewer1 
      Height          =   6615
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11668
      LicenseKEY      =   "1519740135762015145551548"
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PREKLICI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
      MICON           =   "templati.frx":0000
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "DODAJ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
      MICON           =   "templati.frx":001C
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   7320
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin GdPicturePro4.Imaging Imaging1 
      Left            =   120
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X,Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2925
      TabIndex        =   13
      Top             =   8025
      Width           =   345
   End
   Begin VB.Label lbPos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0,0)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   12
      Top             =   7770
      Width           =   735
   End
   Begin VB.Label lbPage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2820
      TabIndex        =   8
      Top             =   7125
      Width           =   495
   End
End
Attribute VB_Name = "templati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub Form_Load()
ReSizeForm Me
 SQL = "Select naziv,pozicija from izpisi where tip_dok=' '"
  Filll List1, SQL
  Me.List1.Refresh
Me.GdViewer1.CloseImage
   Me.GdViewer1.DisplayFromFile (App.path & "\print\Report 1.htm")
     Me.GdViewer1.SetZoomFitControl
End Sub
Private Sub Form_Initialize()
  On Error Resume Next
  InitCommonControls 'To get the XP theme
End Sub

Private Sub btFirstpage_Click()
   Me.GdViewer1.DisplayFirstFrame
End Sub

Private Sub btlastpage_Click()
   Me.GdViewer1.DisplayLastFrame
End Sub

Private Sub btNextPage_Click()
   Me.GdViewer1.DisplayNextFrame
End Sub

Private Sub btPreviousPage_Click()
   Me.GdViewer1.DisplayPreviousFrame
End Sub

Private Sub btZoomIn_Click()
   Me.GdViewer1.ZoomIN
End Sub

Private Sub btZoomOut_Click()
   Me.GdViewer1.ZoomOUT
End Sub

Private Sub gdViewer1_MouseMoveControl(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.lbPos.Caption = "(" & Me.GdViewer1.GetMouseX & "," & Me.GdViewer1.GetMouseY & ")"
  
   
End Sub

Private Sub gdViewer1_PageChange()
   Me.lbPage.Caption = Me.GdViewer1.CurrentPage & "/" & Me.GdViewer1.NumPages
End Sub

Private Sub gdViewer1_ZoomChanged()
   Me.txtZoom.text = GdViewer1.Zoom * 100
End Sub


