VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPR 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Artik vnos nabave"
   ClientHeight    =   9960
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15300
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   17.568
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.988
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Preklici"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Shrani"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nov"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   7320
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   2640
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00F9F0EB&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   15240
      TabIndex        =   12
      Top             =   0
      Width           =   15300
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   9840
         TabIndex        =   34
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63504385
         CurrentDate     =   39332
      End
      Begin VB.TextBox DTPICKER41 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   9800
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox emb 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   255
         Left            =   6840
         TabIndex        =   30
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox cmbSupp 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9800
         TabIndex        =   29
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtDis 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtRate 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtQTY 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2300
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   255
         TabIndex        =   1
         Top             =   1560
         Width           =   2000
      End
      Begin VB.ComboBox cmbCust 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPR.frx":0000
         Left            =   9800
         List            =   "frmPR.frx":0002
         TabIndex        =   0
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Dodaj"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7000
         TabIndex        =   6
         Top             =   1635
         Width           =   650
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalaza"
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
         Left            =   3195
         TabIndex        =   31
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "St. dokumenta"
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
         Left            =   8500
         TabIndex        =   27
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Popust"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cena"
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
         Left            =   4200
         TabIndex        =   22
         Top             =   1350
         Width           =   450
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vnesi artikel"
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
         TabIndex        =   21
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dobavitelj"
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
         Left            =   8500
         TabIndex        =   13
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxx"
         ForeColor       =   &H00F9F0EB&
         Height          =   240
         Left            =   9510
         TabIndex        =   16
         Top             =   255
         Width           =   2235
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vnos Nabave"
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
         TabIndex        =   15
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datum"
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
         Left            =   8500
         TabIndex        =   14
         Top             =   1320
         Width           =   570
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000001&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1320
         Left            =   8500
         Shape           =   4  'Rounded Rectangle
         Top             =   765
         Width           =   3930
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Znesek"
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
         Left            =   5745
         TabIndex        =   19
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kolicina"
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
         Left            =   2300
         TabIndex        =   18
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel"
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
         Left            =   195
         TabIndex        =   17
         Top             =   1395
         Width           =   555
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   0
         Picture         =   "frmPR.frx":0004
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12330
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
      DragMode        =   1  'Automatic
      Height          =   4455
      Left            =   2520
      TabIndex        =   10
      Top             =   2280
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Naziv"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cena"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Popust"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Znesek"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Embalaza"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvLook 
      Height          =   4455
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16380139
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Naziv"
         Object.Width           =   7000
      EndProperty
   End
   Begin VB.Label lblwords 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   28
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   390
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
      Left            =   8504
      TabIndex        =   11
      Top             =   7080
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
      Caption         =   "Znesek:"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   8504
      TabIndex        =   20
      Top             =   7140
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000003&
      BorderColor     =   &H80000001&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F9F0EB&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   8504
      Shape           =   4  'Rounded Rectangle
      Top             =   7005
      Width           =   3690
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BorderColor     =   &H80000001&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   8504
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3945
   End
   Begin VB.Menu mnuLook 
      Caption         =   "Look"
      Visible         =   0   'False
      Begin VB.Menu mnuUnit 
         Caption         =   "Unit In Stock"
      End
   End
End
Attribute VB_Name = "frmPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PurchReg As Boolean
Dim PurchOrder As Boolean
Dim PurchRet As Boolean
Dim Suppmadanazi As String
Dim SuppID As String
Dim PurchID As String
Dim ReturnID As String


Dim TxtLen As Integer
Dim STRT As Integer
Dim PRILEN As Integer
Private Sub cmbCus_lostfocus()
If Me.cmbCust.text = "" Then
MsgBox ("St.partnerja je obvezen podatek!")
Me.cmbCust.Enabled = True
Me.cmbCust.SetFocus
End If
End Sub
Private Sub cmbCust_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select sifra from partner where naziv='" & cmbCust & "'")

If Rs1.RecordCount <> 0 Then
    SuppID = Rs1!sifra
   
End If
Set Rs1 = Nothing
Set DCON = Nothing

'Call CMB1("v", "stdok", cmbSupp, "where naziv='" & cmbCust.text & "' and Deliver=0", True)
    
'Call CMB2("Select DIstinct stdok from nabasif where naziv='" & cmbCust.text & "'", cmbSupp)
'cmbSupp.AddItem ("<ALL>")

lvMain.ListItems.clear
Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""

EDT = False

If cmbCust.text <> "" Then
    Text4.Enabled = True
Else
    Text4.Enabled = False
End If



End Sub

Private Sub RegModify()
Dim lst As ListItem
Dim tempSql As String
GetNewConnection2
Set Rs1 = New ADODB.Recordset
lvMain.ListItems.clear

'If cmbSupp.text = "<ALL>" Then
'
'     tempSql = "select * from nabasif Where naziv='" & Trim(cmbCust.text) & "'and deliver=0"
'Else
     tempSql = "select * from nabasif Where stdok='" & MODIFYID & "'"
'End If

Set Rs1 = DCON.Execute(tempSql)
    While Rs1.EOF <> True
       Set lst = lvMain.ListItems.Add(, , Rs1!sifra)
        lst.SubItems(1) = Getnazi("select madanazi from mada where madasifr='" & Rs1!sifra & "'")
        
        lst.SubItems(2) = Rs1!kol
        lst.SubItems(3) = Rs1!nabcena
         lst.SubItems(4) = Rs1!pop
        lst.SubItems(5) = Val(Rs1!kol) * Val(Rs1!nabcena)
        lst.SubItems(6) = Rs1!embalaza
        
        Rs1.MoveNext
    Wend

    Set DCON = Nothing
    Set Rs1 = Nothing
End Sub
Private Sub cmdAdd_Click()
Dim LSTITEM As ListItem
Dim CNT As Boolean
Dim dd As Integer
If text5.text <> "" Then
   
    CNT = False
  

    Call GetNewConnection2
        Set Rs1 = New Recordset
        Set Rs1 = DCON.Execute("Select * from mada where madasifr='" & (Text4) & "' OR madanazi like'" & Text4 & "%'")

If Not Rs1.EOF Then
    
     
'Set LSTITEM = ListView1.FindItem(RS1!madasifr, lvwText, , lvwPartial)
'       If LSTITEM Is Nothing Then
            
       
       'LBL_DES.Caption = RS1!madasifr & ", " & RS1!madanazi & ""
       txtRate = Rs1!madanabc
            
  With lvMain
        TextAmount.text = ""
        
        If .ListItems.Count <> 0 Then
          
            For dd = 1 To .ListItems.Count
               
                If InStr(1, .ListItems(dd).text, Rs1!madasifr) = 1 Then
                    If InStr(1, .ListItems(dd).SubItems(1), Rs1!MADANAZI) = 1 Then
              
                        If EDT = True Then
                            .ListItems(dd).Selected = True
                            .ListItems(dd).SubItems(2) = Val(txtQTY.text)
                            .ListItems(dd).SubItems(3) = Val(txtRate.text)
                            If txtDis = ".00  %" Then
                            txtDis = ""
                            End If
                            .ListItems(dd).SubItems(4) = Format(Val(txtDis) / 100, ".00  %")
                            .ListItems(dd).SubItems(5) = text5.text
                            .ListItems(dd).SubItems(6) = emb.text
                        '  .ListItems(DD).SubItems(6) = Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5))) * Val(.ListItems(DD).SubItems(4)))
                        
                             '     .ListItems(DD).SubItems(6) = Val(Val(txtRate.text) * Val(txtQTY.text) - Val(txtRate.text) * Val(txtQTY.text) * Val(txtDis) / 100) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
                        ' .ListItems(dd).SubItems(6) = Val(Val(txtRate.text) * Val(txtQTY.text) - Val(txtRate.text) * Val(txtQTY.text) * Val(txtDis) / 100) '* 'Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5)) * Val(.ListItems(DD).SubItems(4))))
                     
                        Else
                            .ListItems(dd).Selected = True
                            .ListItems(dd).SubItems(2) = Val(.ListItems(dd).SubItems(2)) + Val(txtQTY.text)
                            .ListItems(dd).SubItems(3) = Val(txtRate.text)
                              If txtDis = ".00  %" Then
                            txtDis = ""
                            End If
                            .ListItems(dd).SubItems(4) = Format(Val(txtDis) / 100, ".00  %")
                           .ListItems(dd).SubItems(5) = Val(.ListItems(dd).SubItems(2)) * Val(.ListItems(dd).SubItems(3))
                           ' .ListItems(DD).SubItems(6) = Val(.ListItems(DD).SubItems(5)) - Val(Val(Val(.ListItems(DD).SubItems(5))) * Val(.ListItems(DD).SubItems(4)))
                          .ListItems(dd).SubItems(6) = emb.text
            
            
                     
                    
                        End If
                             
                    CNT = True
                    
                    End If
                End If
                   
            Next
       
         End If
            
        If CNT = False Then

         .ListItems.Add , , Rs1!madasifr
            .ListItems(.ListItems.Count).SubItems(1) = Rs1!MADANAZI
            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
            If txtDis.text = ".00  %" Then
            txtDis.text = 0
            End If
            
            .ListItems(.ListItems.Count).SubItems(4) = Format(Val(txtDis) / 100, ".00  %")
                           
            .ListItems(.ListItems.Count).SubItems(5) = text5.text
             
                 
            .ListItems(.ListItems.Count).SubItems(6) = emb.text
             ' TextAmount.text = Val(TextAmount.text) + Val(text5.text)
        
        End If
           
            
             
        
'        If .ListItems.Count <= 0 Then
'
'            .ListItems.Add 1, , RS1!madasifr
'            .ListItems(.ListItems.Count).SubItems(1) = RS1!madanazi
'            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
'
'
'        End If


            For dd = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(dd).SubItems(5)) + Val(TextAmount.text)
            Next
            
        ' lblunit.Caption = RS1!madazalo
       
        
       Set Rs1 = DCON.Execute("Select * from Product")
      '  Set DataGrid1.DataSource = RS1
         
        Set Rs1 = Nothing
        Set DCON = Nothing
   
  End With

Else
    MsgBox "mada Not Found", vbInformation, "Product"
    
    
End If
    
Text4.text = ""
text5.text = ""
txtQTY.text = "1"
txtRate.text = ""
txtDis.text = ""
emb.text = ""
EDT = False

'If Me.Enabled = True Then txtQTY.SetFocus:
Else


'
'txtQTY.SetFocus


End If
Me.Text4.SetFocus

End Sub

Private Sub cmdAdd_GotFocus()
'Call GetNewConnection2
'
'Set RS1 = New Recordset
'SQL = "Select TOP 5 * from mada where madasifr like '" & Text4 & "%' OR madanazi like'" & Text4 & "%'"
'
'Set RS1 = DCON.Execute(SQL)
'
'
'
'
'    SQL = "UPDATE mada set UnitSellingPrice=" & Val(txtrate.text) & " where (madasifr=" & RS1!madasifr & "' AND madanabc <" & Val(txtrate.text) & ")"
'    'MsgBox SQL
'    DCON.Execute SQL
'
'
'Set RS1 = Nothing
'Set DCON = Nothing
'
'Text5.text = Val(txtqty.text) * Val(txtrate.text)
vvvv = 1
End Sub

Private Sub cmdClear_Click()
EDT = False
lvMain.ListItems.clear

picTop.Enabled = False

lvLook.Enabled = False
lvMain.Enabled = False
Text4.text = ""
text5.text = ""
txtQTY.text = "1"
txtRate.text = ""
txtDis.text = ""
emb.text = ""

cmdOK.Enabled = False
cmdClear.Enabled = False
Command1.Enabled = True

cmbCust.ListIndex = -1

End Sub

Private Sub cmdOK_Click()
If Me.cmbCust.text = "" Then
MsgBox ("St.partnerja je obvezen podatek!")
Me.cmbCust.Enabled = True
Me.cmbCust.SetFocus
Else
If Me.cmbSupp.text = "" Then
MsgBox ("St.dokumenta je obvezen podatek!")
Me.cmbSupp.Enabled = True
Me.cmbSupp.SetFocus

Else
If lvMain.ListItems.Count > 0 Then
cmdClear.Enabled = False
cmdOK.Enabled = False
Command1.Enabled = True
picTop.Enabled = False
lvLook.Enabled = False
lvMain.Enabled = False

Call PurchaseReg
    MsgBox "Shranjeno", vbInformation
    
Else
    MsgBox "Ni podatkov", vbInformation
End If
End If
End If

End Sub





'Private Sub Combo5_Click()
'On Error Resume Next
'
'Call GetNewConnection2
'
'Set Rs1 = New Recordset
'Set Rs1 = DCON.Execute("Select * from PurchaseRegistryDetail where PurchaseRegistryID='" & Combo5.text & "'")
'
'While Not Rs1.EOF
'
'    Dim LVITEM As ListItem
'    Dim ProdID As String
'
'    With lvMain
'    .ListItems.clear
'
'    ProdID = Rs1!madasifr
'    Set LVITEM = .ListItems.Add(, , Rs1!madasifr)
'        LVITEM.SubItems(2) = Rs1!kol
'        LVITEM.SubItems(4) = Rs1!discount
'        LVITEM.SubItems(3) = Rs1!Rate
'
'    Set RS2 = New Recordset
'    Set RS2 = DCON.Execute("SElect * from mada where madasifr=" & ProdID & "'")
'        LVITEM.SubItems(1) = RS2!madanazi
'
'
'    End With
'
'    Rs1.MoveNext
'Wend
'
'
'
'Set Rs1 = Nothing
'Set RS2 = Nothing
'Set DCON = Nothing
'End Sub


Private Sub Command1_Click()
Dim CDb As CDbase
Dim CIns As New CInsert


Call GetNewConnection(CIns)
Set CDb = CIns



'PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseRegistryHeader", "PurchaseRegistryID", "PR", Label1)


Set CIns = Nothing

'cmbSupp.ListIndex = -1
cmdClear.Enabled = True
lvMain.ListItems.clear
cmdOK.Enabled = True
Command1.Enabled = False
picTop.Enabled = True
lvLook.Enabled = True
lvMain.Enabled = True

Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""
cmbCust.SetFocus


End Sub

Private Sub DTPICKER1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Command2_Click()
'MsgBox (cmbCust.text)
'MsgBox (DTPicker4.Value)
End Sub

Private Sub Form_Load()

'ReSizeForm Me


PurchOrder = True
PurchReg = False
PurchRet = False
Timer2.Enabled = True
Timer2.Interval = 100
 '   cmbCust.AddItem "Cash"
    
    
    Call CMB1("partner", "naziv", cmbCust)
cmbCust.ListIndex = -1
    Me.DTPicker4.Value = Date
If edi = 1 Then
Dim CDb As CDbase
Dim CIns As New CInsert


Call GetNewConnection(CIns)
Set CDb = CIns



'PurchID = CIns.AUTONUM(CDb.OpenDb, "PurchaseRegistryHeader", "PurchaseRegistryID", "PR", Label1)


Set CIns = Nothing

'cmbSupp.ListIndex = -1
cmdClear.Enabled = True
lvMain.ListItems.clear
cmdOK.Enabled = True
Command1.Enabled = False
picTop.Enabled = True
lvLook.Enabled = True
lvMain.Enabled = True

Text4.text = ""
text5.text = ""
txtQTY.text = ""
txtRate.text = ""
txtDis.text = ""
'cmbCust.SetFocus

edi = 0
Set Rs1 = myConection.Execute("Select * from nabasif where stdok='" & std & "' and sifrapart=" & dob)

If Not Rs1.EOF Then
Me.DTPicker4.Value = Rs1!datum
cmbCust.text = Getnazi("select naziv from partner where sifra=" & Rs1!sifrapart)
Me.cmbSupp.text = Rs1!stdok
Rs1.MoveFirst
cmbCust.Enabled = False
cmbSupp.Enabled = False
DTPicker4.Enabled = False

Dim dd As Integer
dd = 1
Do While Not Rs1.EOF
 With lvMain
  'If .ListItems.Count <> 0 Then
          
           
               .ListItems.Add , , Rs1!sifra
               .ListItems(.ListItems.Count).SubItems(1) = Getnazi("select madanazi from mada where madasifr='" & Rs1!sifra & "'")
              .ListItems(.ListItems.Count).SubItems(2) = Rs1!kol
                          .ListItems(.ListItems.Count).SubItems(3) = Rs1!nabcena
                         
                          .ListItems(.ListItems.Count).SubItems(4) = Format(Rs1!pop, ".00  %")
                          .ListItems(.ListItems.Count).SubItems(5) = Rs1!kol * Rs1!nabcena
                          .ListItems(.ListItems.Count).SubItems(6) = Rs1!embalaza
           
       ' End If
Rs1.MoveNext
End With
Loop

End If
End If
End Sub

Private Sub Form_Resize()
'picBody.Height = Me.ScaleHeight - picTop.Height - picBottom

End Sub

Private Sub Form_Unload(Cancel As Integer)

       'frmMAIN.WindowState = 0
End Sub


Private Sub LvHeads()
    lvMain.ColumnHeaders(1).Width = lvMain.Width * 0.1
    lvMain.ColumnHeaders(2).Width = lvMain.Width * 0.2
    lvMain.ColumnHeaders(3).Width = lvMain.Width * 0.2

End Sub


Private Sub ReOrder()
Call GetNewConnection2

With lvLook

SQL = "Select * from mada where madazalo <= ReOrderLevel"

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)

While Not Rs1.EOF

 .ListItems.Add , , Rs1!MADANAZI
 


Rs1.MoveNext
Wend






End With

End Sub


Private Function GetProduct(ProdID As String) As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select madanazi from mada where madasifr='" & ProdID & "'")

If Not Rs1.EOF Then

'Getmada = Rs1!MADANAZI

Else
    MsgBox "Ne najdem"
    Exit Function
    
End If


Set Rs1 = Nothing
Set DCON = Nothing



End Function


Private Sub PurchaseReg()

Dim CNT1 As Integer
myConection.Execute "delete from nabasif where stdok='" & cmbSupp.text & "' and sifrapart=" & Getnazi("select sifra from partner where naziv='" & Me.cmbCust.text & "'")
With lvMain
If .ListItems.Count > 0 Then
Call GetNewConnection2
   
   ' If CRED = True Then
  '' If cmbCust.text <> "Cash" Then
   
        'PurchaseRegistryHeader
        'PurchaseRegistryDetail
        
      
    
     '  ' DCON.Execute SQL
      
        
  ' ' End If

        For CNT1 = 1 To .ListItems.Count
            Dim sii As Integer
            
               Dim popust As String
popust = Replace(.ListItems(CNT1).SubItems(3), "%", 1, Len(.ListItems(CNT1).SubItems(3)), 1, vbTextCompare)
        sii = Getnazi("select madasifr from mada where madanazi='" & .ListItems(CNT1).SubItems(1) & "'")
        SQL = "Insert into nabasif (datum,stdok,sifra,kol,nabcena,sifrapart,embalaza,pop) values ('" & Me.Label8.Caption & "','" & cmbSupp.text & "'" & "," & Round(sii, 0) & "," & .ListItems(CNT1).SubItems(2) & "," & .ListItems(CNT1).SubItems(3) & "," & Getnazi("select sifra from partner where naziv='" & Me.cmbCust.text & "'") & "," & .ListItems(CNT1).SubItems(6) & "," & Left(.ListItems(CNT1).SubItems(4), 4) & ")"
        DCON.Execute SQL

    'Set Rs1 = New Recordset

       ' sql1 = "Select * from mada where madasifr=" & .ListItems(CNT1).text

       'Set Rs1 = DCON.Execute(sql)

      ' sql = "update mada set madazalo=" & Val(Val(Rs1!madazalo) + Val(.ListItems(CNT1).SubItems(2))) _
       '             & " WHERE madasifr=" & .ListItems(CNT1).text

       
      ' DCON.Execute sql
       
   
        
    Next
        
        
    
  
       
        
Set Rs1 = Nothing
Set DCON = Nothing

End If

End With

End Sub


Private Sub lvLook_DblClick()
'Dim LVindex As Integer
'
'Dim LSTITEM As ListItem
'
'
'
'Dim CNT As Boolean
'Dim DD As Integer
'
'''' query in kol is not yet included
'
'
'
'    CNT = False
'
'
'    Call GetNewConnection2
'        Set Rs1 = New Recordset
'        Set Rs1 = DCON.Execute("Select * from mada where madasifr=" & lvLook.SelectedItem.text & "'")
'
'If Not Rs1.EOF Then
'
'
'
'       'LBL_DES.Caption = RS1!madasifr & ", " & RS1!madanazi & ""
'       txtRate = Rs1!madanabc
'     '  txtQTY.text = RS1!ReOrderkol
'
'  With lvMain
'        TextAmount.text = ""
'
'        If .ListItems.Count <> 0 Then
'
'            For DD = 1 To .ListItems.Count
'
'              If InStr(1, .ListItems(DD).text, Rs1!madasifr) = 1 Then
'                If InStr(1, .ListItems(DD).SubItems(1), Rs1!madanazi) = 1 Then
'
'                        If EDT = True Then
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = text5.text
'                        Else
'                            .ListItems(DD).Selected = True
'                            .ListItems(DD).SubItems(2) = Val(.ListItems(DD).SubItems(2)) + Val(txtQTY.text)
'                            .ListItems(DD).SubItems(3) = Val(txtRate.text)
'                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(2)) * Val(.ListItems(DD).SubItems(3))
'
'                        End If
'
'                    CNT = True
'
'                End If
'
'               End If
'
'            Next
'
'         End If
'
'        If CNT = False Then
'
'         .ListItems.Add , , Rs1!madasifr
'            .ListItems(.ListItems.Count).SubItems(1) = Rs1!madanazi
'            .ListItems(.ListItems.Count).SubItems(2) = txtQTY.text
'            .ListItems(.ListItems.Count).SubItems(3) = txtRate.text
'            .ListItems(.ListItems.Count).SubItems(4) = "1"
'            .ListItems(.ListItems.Count).SubItems(5) = text5.text
'
'
'             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
'
'        End If
'
'
'
'
''        If .ListItems.Count <= 0 Then
''
''            .ListItems.Add 1, , RS1!madasifr
''            .ListItems(.ListItems.Count).SubItems(1) = RS1!madanazi
''            .ListItems(.ListItems.Count).SubItems(2) = txtqty.text
''            .ListItems(.ListItems.Count).SubItems(3) = txtrate.text
''            .ListItems(.ListItems.Count).SubItems(4) = "1"
''            .ListItems(.ListItems.Count).SubItems(5) = Text5.text
''
''
''        End If
'
'
'            For DD = 1 To .ListItems.Count
'                  TextAmount.text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.text)
'            Next
'
'        ' lblunit.Caption = RS1!madazalo
'
'
'       Set Rs1 = DCON.Execute("Select * from Product")
'      '  Set DataGrid1.DataSource = RS1
'
'        Set Rs1 = Nothing
'        Set DCON = Nothing
'
'  End With
'
'Else
'    MsgBox "mada Not Found", vbInformation, "Product"
'
'
'End If
End Sub

Private Sub lvLook_ItemClick(ByVal Item As MSComctlLib.ListItem)
Timer2.Enabled = False
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from mada where madasifr='" & lvLook.SelectedItem.text & "'")

    If Not Rs1.EOF Then
        txtRate.text = Rs1!madanabc
    End If
    

If lvLook.ListItems.Count > 0 Then
    Text4.text = lvLook.SelectedItem.text
    
End If

End Sub

Private Sub lvLook_LostFocus()
Timer2.Enabled = True
End Sub

Private Sub lvLook_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
    PopupMenu mnuLook
'End If
End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
If PurchRet = True Then
    Timer2.Enabled = False
    

Else

End If
EDT = True
If lvMain.ListItems.Count <> 0 Then
      Text4.text = lvMain.ListItems(lvMain.SelectedItem.Index).text
    txtQTY.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(2)
    txtRate.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(3)
    txtDis = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4)
    text5.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(5)
    emb.text = lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(6)
        
End If




End Sub

Private Sub lvMain_KeyDown(KeyCode As Integer, Shift As Integer)
Dim dd As Integer

With lvMain
If .ListItems.Count <> 0 Then
If KeyCode = vbKeyDelete Then
TextAmount.text = ""
       
  .ListItems.Remove (.SelectedItem.Index)
  
Text4.text = ""

  'lblcat.Caption = ""
EDT = False
txtRate.text = ""
text5.text = ""
txtQTY.text = "1"
emb.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""

            For dd = 1 To .ListItems.Count
                  TextAmount.text = Val(.ListItems(dd).SubItems(5)) + Val(TextAmount.text)
            Next
 End If
End If
End With

End Sub

Private Sub mnuUnit_Click()
Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from mada where madasifr='" & lvLook.SelectedItem.text & "'")

If Rs1.RecordCount <> 0 Then
    MsgBox "Na zalogi: " & Rs1!madazalo
End If
Set Rs1 = Nothing
Set DCON = Nothing
End Sub



Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Getnazi("select madaenme from mada where madasifr=" & Me.Text4.text) = "KOM" Then
Me.txtRate.SetFocus
'SendKeys "{enter}"
Else
    emb.SetFocus
    End If
End If
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQTY.SetFocus
End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)


'MsgBox ("5")

Select Case KeyCode
 Case vbKeyA To vbKeyZ
Dim idar As String
'zap = Indx
'opp = Me.cmbItmcode.Top
'oppa = Me.cmbItmcode.Left
idar = Chr(KeyCode)
   DoSQL3 "mada", "madasifr", "madanazi"
       
Case Else
    End Select
    End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)

If xxre <> "" Then

Me.Text4 = xxre
SendKeys "{enter}"
xxre = ""
End If
End Sub

Private Sub Text4_Change()

 

Me.emb.text = "1"
If Len(Text4.text) = 0 Then
    Timer2.Enabled = True
End If

'If EDT = False Then
Timer2.Interval = 100


TxtLen = Len(Text4.text)
STRT = 0

'End If
'EDT = True
txtQTY.text = "1"

If Len(Text4.text) <> 0 Then
    txtQTY.Enabled = True
Else
    txtQTY.Enabled = False
    
End If
End Sub

Private Sub TextAmount_Change()
lblwords.Caption = NumToWord(TextAmount.text)
End Sub





Private Sub Timer2_Timer()
Me.Label8.Caption = Me.DTPicker4.Value
'Static c As Integer


STRT = STRT + 1



If STRT = 3 Then
Timer2.Interval = 0
  




Dim FVAL As String
Dim dd As Integer
Dim LISTITM As ListItem

Call GetNewConnection2

Set Rs1 = New Recordset

'SQL = "Select TOP 10 * from mada where madasifr like '" & Text4 & "%' OR madanazi like'" & Text4 & "%'"
SQL = "Select  * from mada where (madasifr=" & Val(Text4) & " OR madanazi like'" & Text4 & "%')"
'SQL = "Select Top 20 * from lowINstock order by Total"
Set Rs1 = DCON.Execute(SQL)
 Set RS2 = New Recordset
        Set RS2 = DCON.Execute(SQL)
        lvLook.ListItems.clear
        While Not RS2.EOF
        
            Set LISTITM = lvLook.ListItems.Add(, , RS2!madasifr)
                LISTITM.SubItems(1) = RS2!MADANAZI
               
                If RS2!madazalo <= 0 Then
                    LISTITM.ForeColor = vbRed
                    LISTITM.ListSubItems(1).ForeColor = vbRed
                Else
                    If RS2!madazalo <= RS2!madaminz Then
                    LISTITM.ForeColor = vbBlue
                    LISTITM.ListSubItems(1).ForeColor = vbBlue
                    End If
                End If
        
            RS2.MoveNext
        Wend

If Text4.text <> "" Then

    If Not Rs1.EOF Then
'        TXT_CODE.SelStart = PRILEN
'        TXT_CODE.text = RS1!madanazi
'        TXT_CODE.SelLength = Len(TXT_CODE.text)
'
      
        FVAL = Rs1!madasifr
        
      
       txtRate.text = Rs1!madanabc
       
        
       text5.text = Val(txtQTY.text) * Val(txtRate.text)
        
        'lblselling.Caption = RS1!UnitSellingPrice
        'lblunit.Caption = RS1!madazalo

  With lvMain
        If .ListItems.Count <> 0 Then
          
            For dd = 1 To .ListItems.Count
                  
                If InStr(1, .ListItems(dd).SubItems(1), Rs1!madasifr) = 1 Then
                  
                 
                            .ListItems(dd).Selected = True
                         '   lblunit.Caption = Val(lblunit.Caption) - Val(.ListItems(DD).SubItems(3))
                    
                    
                End If
            
               
            Next
         End If
    End With

  

    Else
      txtRate.text = ""
        text5.text = ""
'        TXT_QTY.text = ""
'        lblselling.Caption = ""
'        lblunit.Caption = ""
'        lblcat.Caption = ""
'        PRILEN = 0

    End If

    
    Set Rs1 = Nothing
    Set RS2 = Nothing
    Set DCON = Nothing


ElseIf Text4.text = "" Then
   txtRate.text = ""
        text5.text = ""
        txtQTY.text = "1"
        emb.text = ""
'lblcat.Caption = ""
'EDT = False
'TXT_RATE.text = ""
'TXT_AMT.text = ""
'TXT_QTY.text = ""
'lblselling.Caption = ""
'lblunit.Caption = ""
'TXT_CODE.SetFocus
End If

End If



End Sub

Private Sub txtQty_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)

If Len(txtQTY.text) <> 0 Then
    txtRate.Enabled = True
Else
    txtRate.Enabled = False
    
End If

If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
End Sub
Private Sub txtQTY_GotFocus()
SendKeys "{RIGHT}"
End Sub
Private Sub txtQTY_KeyPress1(KeyAscii As Integer)
'Call OFFCHar(KeyAscii, txtQTY)

End Sub

Private Sub txtRate_Change()
text5.text = Val(txtQTY.text) * Val(txtRate.text)
If Len(txtRate.text) <> 0 And Val(txtRate.text) <> 0 And Val(txtQTY.text) <> 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
    
End If
If Len(txtRate.text) <> 0 Then
    txtDis.Enabled = True
Else
    txtDis.Enabled = False
    
End If
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

Call Decimals(KeyAscii, txtRate, 2)

If KeyAscii = 13 Then

    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from mada where madasifr=" & Text4 & " OR madanazi='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from mada where (madasifr=" & RS1!madasifr & "' AND madanabc <" & Val(txtrate.text) & ")"
'
            Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                                     If Val(Rs1!madanabc) < Val(txtRate.text) Then
                    
                         '   If MsgBox("The unit cost price has increase, do you want to update " & vbCrLf _
                         '   & "unit selling price?   ", vbQuestion + vbYesNo) = vbYes Then
                                Dim SelPrice As String
                            '         SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.Text) * 1.1)
                                
                              '  While IsNumeric(SelPrice) = False Or Val(txtRate.Text) > Val(SelPrice)
                             '       SelPrice = InputBox("Input New Unit Selling Price for this Product:", , Val(txtRate.Text) * 1.1)
                                          
                              '  Wend
                                  'SQL = "UPDATE mada set madanabc=" & Val(SelPrice) & " where madasifr=" & Rs1!madasifr
   
                                   '         DCON.Execute SQL
                          '   Else
                           '     txtRate.SetFocus
                                
                           'End If
                    End If
                            
'                SQL = "UPDATE mada set madanabc=" & Val(txtrate.text) & " where (madasifr=" & RS1!madasifr & "' AND madanabc <" & Val(txtrate.text) & ")"
                    SQL = "UPDATE mada set madanabc=" & Val(txtRate.text) & " where madasifr=" & Rs1!madasifr
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from mada where madasifr like '" & Text4 & "%' OR madanazi like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!madanabc <> txtrate.text Then
'
'                         MsgBox "Cannot update madanabc" & vbTab, vbInformation, "madanabc"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!madanabc
'                         txtrate.SetFocus
                   End If
Me.cmdAdd.Enabled = True
     Me.cmdAdd.SetFocus
         

Set Rs1 = Nothing
Set DCON = Nothing



End If
End Sub

Private Sub txtrate_LostFocus()
  
    Call GetNewConnection2

        Set Rs1 = New Recordset
            SQL = "Select * from mada where madasifr=" & Text4 & " OR madanazi='" & Text4 & "'"

'        Set RS1 = DCON.Execute(SQL)
'
'            SQL = "Select * from mada where (madasifr=" & RS1!madasifr & "' AND madanabc <" & Val(txtrate.text) & ")"
'
        Set Rs1 = DCON.Execute(SQL)
'
          If Rs1.RecordCount <> 0 Then
                       If Val(Rs1!madanabc) < Val(txtRate.text) Then
                    
                            'If MsgBox("Nabavna cena se je spremenila jo zamenjam v šifrantu? " & vbCrLf _
                            '& "HA?   ", vbQuestion + vbYesNo) = vbYes Then
                            '    Dim SelPrice As String
                            '         SelPrice = InputBox("Nabavna cena se je spremenila jo zamenjam v šifrantu:", , Val(txtRate.Text) * 1.1)
                                
                             '   While IsNumeric(SelPrice) = False Or Val(txtRate.Text) > Val(SelPrice)
                              '      SelPrice = InputBox("Nabavna cena se je spremenila jo zamenjam v šifrantu:", , Val(txtRate.Text) * 1.1)
                                          
                               ' Wend
                                  'SQL = "UPDATE mada set madanabc=" & Val(SelPrice) & " where (madasifr=" & Rs1!madasifr & ")"
   
  '                                          DCON.Execute SQL
                            ' Else
                              '  txtRate.SetFocus
                             '
                           'End If
                    End If
'                SQL = "UPDATE mada set madanabc=" & Val(txtrate.text) & " where (madasifr=" & RS1!madasifr & "' AND madanabc <" & Val(txtrate.text) & ")"
               SQL = "UPDATE mada set madanabc=" & Val(txtRate.text) & " where madasifr=" & Rs1!madasifr
              
                
                DCON.Execute SQL
'            Else
                
'                SQL = "Select * from mada where madasifr like '" & Text4 & "%' OR madanazi like'" & Text4 & "%'"
'                Set RS1 = DCON.Execute(SQL)
'
'                      If RS1!madanabc <> txtrate.text Then
'
'                         MsgBox "Cannot update madanabc" & vbTab, vbInformation, "madanabc"
'
'                      End If
'
'                      If RS1.RecordCount <> 0 Then
'                         txtrate.text = RS1!madanabc
'                         txtrate.SetFocus
'                      End If
                   
                    
            End If
         

Set Rs1 = Nothing
Set DCON = Nothing


End Sub


