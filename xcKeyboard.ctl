VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl xcKeyboard 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3615
   ScaleWidth      =   10875
   Begin VB.CommandButton cmdShift 
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdShift 
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2880
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imlLEDs 
      Left            =   0
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   8
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xcKeyboard.ctx":0000
            Key             =   "grey"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xcKeyboard.ctx":005E
            Key             =   "red"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xcKeyboard.ctx":00BC
            Key             =   "yellow"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xcKeyboard.ctx":011A
            Key             =   "green"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "xcKeyboard.ctx":0178
            Key             =   "cyan"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHyphen 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCapsLock 
      Caption         =   "Caps"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdSpace 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   5175
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdComma 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdDot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAlpha 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdNumeric 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdBackSpace 
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdApost 
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "xcKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
' copyright(c) 2000, 2006 Original Software Designs
'-------------------------------------------------------------------------

Option Explicit

Private m_fCapsLock As Boolean
Private m_fNumLock As Boolean
Private m_fShift As Boolean

Public Event Keyboard(KeyPressed As String)

Private Sub cmdApost_Click()

    RaiseEvent Keyboard("'")
  
End Sub

Private Sub cmdBackSpace_Click()

    RaiseEvent Keyboard("BS")
    
End Sub

Private Sub cmdCapsLock_Click()

    If CapsLock Then
        Set cmdCapsLock.Picture = imlLEDs.ListImages(1).Picture
    Else
        Set cmdCapsLock.Picture = imlLEDs.ListImages(2).Picture
    End If

    CapsLock = Not CapsLock

End Sub

Private Sub cmdComma_Click()

    RaiseEvent Keyboard(",")
  
End Sub

Private Sub cmdDot_Click()

    RaiseEvent Keyboard(".")
  
End Sub

Private Sub cmdEnter_Click()

    RaiseEvent Keyboard("CR")

End Sub

Private Sub cmdHyphen_Click()

    RaiseEvent Keyboard("-")
  
End Sub

Private Sub cmdShift_Click(Index As Integer)

    If Shift Then
        Set cmdShift(0).Picture = imlLEDs.ListImages(1).Picture
        Set cmdShift(1).Picture = imlLEDs.ListImages(1).Picture
    Else
        Set cmdShift(0).Picture = imlLEDs.ListImages(4).Picture
        Set cmdShift(1).Picture = imlLEDs.ListImages(4).Picture
    End If
    
    Shift = Not Shift
    
End Sub

Private Sub cmdSpace_Click()

    RaiseEvent Keyboard(" ")
  
End Sub

Private Sub cmdAlpha_Click(Index As Integer)

    If m_fCapsLock And Not m_fShift Or Not m_fCapsLock And m_fShift Then
        RaiseEvent Keyboard(UCase$(cmdAlpha(Index).Caption))
    Else
        RaiseEvent Keyboard(LCase$(cmdAlpha(Index).Caption))
    End If
    
    If Shift Then cmdShift_Click (0)
  
End Sub

Public Property Get CapsLock() As Boolean

    CapsLock = m_fCapsLock
    
End Property

Public Property Let CapsLock(ByVal NewValue As Boolean)

    If NewValue = m_fCapsLock Then
        Exit Property
    Else
        m_fCapsLock = NewValue
        PropertyChanged "CapsLock"
    End If

End Property

Public Property Get Shift() As Boolean

    Shift = m_fShift
    
End Property

Public Property Let Shift(ByVal NewValue As Boolean)

    If NewValue = m_fShift Then
        Exit Property
    Else
        m_fShift = NewValue
        PropertyChanged "Shift"
    End If

End Property

Private Sub cmdNumeric_Click(Index As Integer)

    RaiseEvent Keyboard(CStr(Index))

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_fCapsLock = PropBag.ReadProperty("CapsLock", False)
    m_fNumLock = PropBag.ReadProperty("NumLock", False)
    m_fShift = PropBag.ReadProperty("Shift", False)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "CapsLock", m_fCapsLock, False
    PropBag.WriteProperty "NumLock", m_fNumLock, False
    PropBag.WriteProperty "Shift", m_fShift, False
    
End Sub



