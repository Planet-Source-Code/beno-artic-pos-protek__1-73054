VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FormPurchaseOrder 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "s"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   585
      Left            =   6240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7200
      Width           =   840
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   11850
      TabIndex        =   4
      Top             =   0
      Width           =   11910
      Begin VB.TextBox TXTNUM 
         Enabled         =   0   'False
         Height          =   375
         Left            =   9480
         TabIndex        =   25
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtQTY 
         Height          =   375
         Left            =   2910
         TabIndex        =   17
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Height          =   375
         Left            =   4065
         TabIndex        =   16
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox text5 
         Height          =   375
         Left            =   5625
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbCust 
         Height          =   315
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   405
         Left            =   7125
         TabIndex        =   1
         Top             =   1635
         Width           =   540
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9360
         TabIndex        =   21
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483639
         CalendarTitleBackColor=   16777215
         Format          =   58064897
         CurrentDate     =   38530
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   9390
         TabIndex        =   23
         Top             =   1605
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58064897
         CurrentDate     =   38530
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   8760
         TabIndex        =   24
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Date"
         Height          =   195
         Left            =   7920
         TabIndex        =   22
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order  ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   7395
         TabIndex        =   13
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lipa Solid Auto Supply"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Left            =   7920
         TabIndex        =   5
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9F0EB&
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2190
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1320
         Left            =   7755
         Shape           =   4  'Rounded Rectangle
         Top             =   765
         Width           =   3930
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   0
         Picture         =   "2_frmPurchaseOrder.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12330
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   5745
         TabIndex        =   9
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name "
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   1395
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Left            =   60
         Top             =   1290
         Width           =   6900
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4455
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvLook 
      Height          =   4455
      Left            =   0
      TabIndex        =   26
      Top             =   2280
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16380139
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblwords 
      Caption         =   "Label4"
      Height          =   615
      Left            =   480
      TabIndex        =   19
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   7170
   End
   Begin MSForms.TextBox TextAmount 
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   7200
      Width           =   1905
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3360;661"
      BorderColor     =   -2147483647
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   8040
      TabIndex        =   10
      Top             =   7200
      Width           =   1530
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000001&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F9F0EB&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   6885
      Width           =   3690
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   7815
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3945
   End
End
Attribute VB_Name = "FormPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
