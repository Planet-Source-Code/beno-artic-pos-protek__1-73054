VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmWindow 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome Window"
   ClientHeight    =   5130
   ClientLeft      =   1095
   ClientTop       =   105
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWindow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   2040
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security"
         Height          =   195
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmWindow.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance"
         Height          =   195
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmWindow.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction"
         Height          =   195
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmWindow.frx":0A56
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Entry"
         Height          =   195
         Index           =   4
         Left            =   240
         MouseIcon       =   "frmWindow.frx":0D60
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reports"
         Height          =   195
         Index           =   3
         Left            =   240
         MouseIcon       =   "frmWindow.frx":106A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Choose Tasks"
         Top             =   300
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C7BDAD&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   60
         Top             =   660
         Visible         =   0   'False
         Width           =   1845
      End
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Don't Show Me This Window"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   2145
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7858
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483639
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "This is a classic Example of Optimizing Bitmpas and Pictures ...... This will be implemented in all the application "
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   4440
      Width           =   3615
   End
End
Attribute VB_Name = "frmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''' to make the most of Images we will use RESource
''' for now just put bitmaps and icons and stirng
'''''' oh by the way its very easy to edit res in .net
'''''command are LoadResPicture,LoadResData,LoadResString
'''id is the picture id when you input it in res::::::You can Customize it and Export Log File of Bitmaps so that
'' we will know waht is that picture is
''by default the id starts from 101 but you can always edit that cooly in .net C++
'''''you can also put you custome dialog boxes messagebox etc etc//// don't try to put GIF and JEPEG
'' that's where it gets very tricky...
'''' but as i say don't waste time here coz its so vast topic:: later


Private Sub Form_Load()
MakeTopMost Me.hWnd

Dim j As Integer
ImageList1.ImageHeight = 64
ImageList1.ImageWidth = 75
For j = 101 To 106
'    ImageList1.ListImages.Add , , LoadResPicture(j, 0)
Next j

'    ListView1.Icons = ImageList1
 '   ListView1.SmallIcons = ImageList1

    
For i = 1 To ImageList1.ListImages.Count
    ListView1.ListItems.Add , , i, i, i
Next i




End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Check1.Value = 1 Then
   Call SaveSetting(App.Title, "FrmWindow", "Show", False)
Else
   Call SaveSetting(App.Title, "FrmWindow", "Show", True)
End If
End Sub

'Private Sub Label1_Click(Index As Integer)
'If Index = 2 Or Index = 3 Or Index = 1 Then
'    ListView1.Icons = Nothing
'    ListView1.SmallIcons = Nothing
'    ListView1.ListItems.clear
'i = 0
'j = 0
'ImageList1.ListImages.clear
'ImageList1.ImageHeight = 55
'ImageList1.ImageWidth = 55
'For j = 101 To 106
'    ImageList1.ListImages.Add , , LoadResPicture(j, 0)
'Next j
'    ListView1.Icons = ImageList1
'    ListView1.SmallIcons = ImageList1
'For i = 1 To ImageList1.ListImages.Count
'    ListView1.ListItems.Add , , i, i, i
'Next i
'ListView1.SelectedItem.Selected = False
'End If
'''''''''''''' this is really messy
'If Index = 0 Or Index = 4 Then
'    ListView1.Icons = Nothing
'    ListView1.SmallIcons = Nothing
'    ListView1.ListItems.clear
'i = 0
'j = 0
'ImageList1.ListImages.clear
'ImageList1.ImageHeight = 64
'ImageList1.ImageWidth = 85
'For j = 101 To 106
'    ImageList1.ListImages.Add , , LoadResPicture(j, 0)
'Next j
'
'    ListView1.Icons = ImageList1
'    ListView1.SmallIcons = ImageList1
'For i = 1 To ImageList1.ListImages.Count
'    ListView1.ListItems.Add , , i, i, i
'Next i
'ListView1.SelectedItem.Selected = False ''not for putting that mark on window
'End If
'
'
'End Sub

