VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form2 
   Caption         =   "Stock Analysis"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form2"
   ScaleHeight     =   3150
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   960
      ScaleHeight     =   1155
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Not In Stock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reorder Level Reached"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60489729
      CurrentDate     =   38580
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   60489729
      CurrentDate     =   38580
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 frmControlMain.WBrow.Visible = True
    frmControlMain.MSHFlexGrid1.Visible = False



SQL = "Select * From dprodstata where [date] between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "#"


Call frmControlMain.CreateSubPage(SQL, "Stock Analysis")
End Sub

Private Sub Image1_Click()

End Sub

