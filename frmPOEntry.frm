VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{622BF7F8-CCD1-49DF-BF0D-7382B298C9DC}#9.0#0"; "b8Controls4.ocx"
Begin VB.Form frmPOEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kalkulacija"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Shrani"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8580
      TabIndex        =   0
      Top             =   6960
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Preklici"
      Height          =   375
      Left            =   10140
      TabIndex        =   1
      Top             =   6960
      Width           =   1395
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E3F9FB&
      Height          =   6885
      Left            =   0
      ScaleHeight     =   459
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   2
      Top             =   540
      Width           =   11625
      Begin VB.TextBox txtRemarks 
         Height          =   735
         Left            =   8160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   5490
         Width           =   3165
      End
      Begin VB.PictureBox Picture1 
         Height          =   2985
         Left            =   990
         ScaleHeight     =   195
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   50
         Top             =   3180
         Width           =   6915
         Begin b8Controls4.LynxGrid3 listPOProd 
            Height          =   2925
            Left            =   0
            TabIndex        =   51
            Top             =   30
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   5159
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorBkg    =   16056319
            BackColorSel    =   8438015
            ForeColorSel    =   0
            GridColor       =   11136767
            BorderStyle     =   0
            FocusRectColor  =   33023
            AllowUserResizing=   4
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
      End
      Begin VB.CommandButton cmdEditPOProd 
         Caption         =   "&Uredi"
         Enabled         =   0   'False
         Height          =   345
         Left            =   6630
         TabIndex        =   49
         Top             =   2820
         Width           =   645
      End
      Begin VB.TextBox txtRefNum 
         Height          =   315
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   47
         Top             =   120
         Width           =   2205
      End
      Begin VB.CommandButton cmdNewProd 
         Height          =   345
         Left            =   7530
         Picture         =   "frmPOEntry.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2220
         Width           =   375
      End
      Begin VB.TextBox txtPOBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9330
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   4590
         Width           =   1995
      End
      Begin VB.TextBox txtPayAmtOnDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9330
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   3840
         Width           =   2025
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   9300
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   2040
         Width           =   2025
      End
      Begin VB.ComboBox cmbFP 
         Height          =   315
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3450
         Width           =   1635
      End
      Begin VB.ComboBox cmbCA 
         Height          =   315
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3030
         Width           =   1635
      End
      Begin VB.CommandButton cmdAddPOProd 
         Caption         =   "&Dodaj"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5970
         TabIndex        =   37
         Top             =   2820
         Width           =   645
      End
      Begin VB.CommandButton cmdDeletePOProd 
         Caption         =   "&Brisi"
         Enabled         =   0   'False
         Height          =   345
         Left            =   7290
         TabIndex        =   33
         Top             =   2820
         Width           =   615
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2850
         Width           =   975
      End
      Begin VB.TextBox txtQtyPrice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   3450
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F4FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   990
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2850
         Width           =   855
      End
      Begin VB.ComboBox cmbPackTitle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2850
         Width           =   1515
      End
      Begin VB.CommandButton cmdNewSup 
         Height          =   345
         Left            =   7530
         Picture         =   "frmPOEntry.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtAP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EAFDFF&
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   1665
      End
      Begin b8Controls4.b8Line b8Line1 
         Height          =   30
         Left            =   -30
         TabIndex        =   14
         Top             =   480
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComCtl2.DTPicker dtpPODate 
         Height          =   315
         Left            =   9780
         TabIndex        =   12
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM - dd - yyyy"
         Format          =   60489731
         CurrentDate     =   38961
      End
      Begin b8Controls4.b8DataPicker b8DPSup 
         Height          =   360
         Left            =   990
         TabIndex        =   9
         Top             =   900
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropWinWidth    =   6210
      End
      Begin b8Controls4.b8DataPicker b8DPProd 
         Height          =   360
         Left            =   990
         TabIndex        =   8
         Top             =   2220
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropWinWidth    =   9735
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00EAFDFF&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   4545
      End
      Begin VB.TextBox txtPOID 
         BackColor       =   &H00F5F5F5&
         Height          =   285
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   1635
      End
      Begin b8Controls4.b8Line b8Line2 
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8GradLine b8GradLine1 
         Height          =   240
         Left            =   0
         TabIndex        =   10
         Top             =   540
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "     Dobavitelj"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8GradLine b8GradLine2 
         Height          =   240
         Left            =   0
         TabIndex        =   23
         Top             =   1860
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   423
         Color1          =   14737632
         Color2          =   16119285
         Caption         =   "      Naroceno blago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin b8Controls4.b8Line b8Line3 
         Height          =   30
         Left            =   8190
         TabIndex        =   34
         Top             =   4440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line4 
         Height          =   30
         Left            =   0
         TabIndex        =   35
         Top             =   1800
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line5 
         Height          =   30
         Left            =   0
         TabIndex        =   36
         Top             =   6300
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8Line b8Line6 
         Height          =   30
         Left            =   8100
         TabIndex        =   41
         Top             =   2580
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   420
         Top             =   4050
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOEntry.frx":0B20
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.b8Line b8Line7 
         Height          =   30
         Left            =   8160
         TabIndex        =   56
         Top             =   5040
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   53
         BorderColor1    =   15592425
         BorderColor2    =   16777215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Opis:"
         Height          =   195
         Left            =   8190
         TabIndex        =   55
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. #:"
         Height          =   195
         Left            =   420
         TabIndex        =   48
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Balance"
         Height          =   195
         Left            =   8190
         TabIndex        =   45
         Top             =   4620
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount:"
         Height          =   195
         Left            =   8220
         TabIndex        =   43
         Top             =   3870
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Less Payment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8100
         TabIndex        =   40
         Top             =   2700
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Znesek:"
         Height          =   195
         Left            =   8190
         TabIndex        =   39
         Top             =   2070
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form of Payment:"
         Height          =   195
         Left            =   8190
         TabIndex        =   21
         Top             =   3450
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Account:"
         Height          =   195
         Left            =   8190
         TabIndex        =   18
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount"
         Height          =   195
         Left            =   4320
         TabIndex        =   32
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit Price"
         Height          =   195
         Left            =   3450
         TabIndex        =   30
         Top             =   2640
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Kol:"
         Height          =   195
         Left            =   990
         TabIndex        =   28
         Top             =   2640
         Width           =   270
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Left            =   1920
         TabIndex        =   26
         Top             =   2640
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel:"
         Height          =   195
         Left            =   300
         TabIndex        =   25
         Top             =   2250
         Width           =   510
      End
      Begin VB.Label lblAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/P:"
         Height          =   195
         Left            =   5820
         TabIndex        =   17
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datum:"
         Height          =   195
         Left            =   8250
         TabIndex        =   13
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Naslov:"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1380
         Width           =   540
      End
      Begin VB.Label lblRM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9450
         TabIndex        =   5
         Top             =   3030
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   195
         Left            =   6030
         TabIndex        =   4
         Top             =   120
         Width           =   225
      End
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
      TabIndex        =   6
      Top             =   0
      Width           =   10305
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vnos prevzemnega lista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   180
         Left            =   600
         TabIndex        =   54
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREVZEM"
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
         Left            =   600
         TabIndex        =   53
         Top             =   30
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmPOEntry.frx":10BA
         Top             =   30
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmPOEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mFormState As String

Dim ProdPackList() As tProdPack

Dim curPO As tPO
Dim newPO As tPO

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isOn As Boolean


Public Function ShowAdd(Optional ByVal dPODate As Date = 0, Optional ByVal lSupID As Long = 0) As Boolean
    
    'set form state
    mFormState = "add"
    
    'evaluate param
    If dPODate = 0 Then
        newPO.PODate = Now
    Else
        newPO.PODate = dPODate
    End If
    newPO.FK_SupID = lSupID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal lPOID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curPO.POID = lPOID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function



Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select

    
    Unload Me
End Sub


Private Sub cmdEditPOProd_Click()
    Call listPOProd_DblClick
End Sub

Private Sub cmdNewProd_Click()
        
    Dim lProdID As Long
    Dim vProd As tProd
    
    
    If frmProdEntry.ShowAddRetID(lProdID) = False Then
        Exit Sub
    End If
    
    If GetProdByID(lProdID, vProd) = False Then
        Exit Sub
    End If
    
    b8DPProd.DisplayData = vProd.ProdDescription
    b8DPProd.BoundData = lProdID
    
    Call b8DPProd_Change
    
End Sub

Private Sub cmdSave_Click()

    Select Case mFormState
        Case "add"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
    
End Sub



Private Sub dtpPODate_Change()
    Form_UseThisSup CLng(GetTxtVal(b8DPSup.BoundData))
End Sub

Private Sub Form_Activate()
        
    
    If isOn = True Then
        Exit Sub
    End If
    isOn = True

    DoEvents: DoEvents: DoEvents
    
    'make mouse pointer bussy
    Me.MousePointer = vbHourglass
   
    Select Case mFormState
        Case "add"
        
            Me.Caption = "Dodaj novo naroèilo"
            
            'add form of payment list
            Form_RefreshFP
            
            'add charge account list
            Form_RefreshCA
                        
            'set form fields
            txtPOID.Text = modFunction.ComNumZ(modRSPO.GetNewPOID, 10)
            dtpPODate.Value = newPO.PODate
            
            If newPO.FK_SupID > 0 Then
                Form_UseThisSup newPO.FK_SupID
            End If
            
            '
            CAFPChange
            
            
        Case "edit"
        
            Me.Caption = "Uredi narocilo"
       
            If GetPOByID(curPO.POID, curPO) = False Then
                'WriteErrorLog Me.Name, "Form_Activate", "Failed on: 'GetPOByID(curPO.POID, vPO) = False'"
                Unload Me
                GoTo RAE
            End If
            
            txtPOID.Text = modFunction.ComNumZ(curPO.POID, 10)
            txtRefNum.Text = curPO.RefNum
            dtpPODate.Value = curPO.PODate
            'set form ui
            Form_UseThisSup curPO.FK_SupID
            
            'load products
            LoadProducts curPO.POID
                       
            'add form of payment list
            Form_RefreshFP curPO.FP
                       
            'add charge account list
            Form_RefreshCA curPO.CA

            'reasign FP
            Form_RefreshFP curPO.FP
            
            txtRemarks.Text = curPO.Remarks
            
            'calculate
            Call Form_CalTotalAmount
            
            
    End Select
    
    
RAE:
    'restoremouse pointer tonormal
    Me.MousePointer = vbNormal
End Sub


Private Sub Form_Load()
    
    isOn = False
    
    PaintGrad bgHeader, &HEDEBE9, &HFFFFFF, 0

    'set po pord column headers
    With listPOProd
    
        .AddColumn "Qty.", 70, lgAlignRightCenter '0
        .AddColumn "InvQty", 100 '1
        .AddColumn "FK_PackID", 0 '2
        .AddColumn "Unit", 80 '3
        .AddColumn "Product ID", 0 '4
        .AddColumn "Articles", 120 '5
        .AddColumn "Unit Price", 70, lgAlignRightCenter '6
        .AddColumn "Amount", 90, lgAlignRightCenter '7
        
        .RowHeightMin = 21
        .ImageList = ilList
    
    End With
    

    'set supplier list
    With b8DPSup
        Set .DropDBCon = PrimeDB
        .SQLFields = "String(10-Len(Trim([SupID])),'0') & [SupID] AS CSupID, tblSup.SupName"
        .SQLTable = "tblSup"
        .SQLWhereFields = "tblSup.SupID, tblSup.SupName"
        .SQLOrderBy = "tblSup.SupName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "Supplier ID", 100
        .AddColumn "Supplier", 240
        
        
    End With
    
    With b8DPProd
        Set .DropDBCon = DCON
        .SQLFields = "String(10-Len(trim(tblProd.ProdID)),'0') & tblProd.ProdID as CProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackTitle, tblCat.CatTitle, Format$([SupPrice],'Fixed') as SP, Format$([SRPrice],'Fixed') as SRP" ' tblProd.SupPrice, tblProd.SRPrice"
        .SQLTable = "tblPack INNER JOIN (tblCat INNER JOIN tblProd ON tblCat.CatID = tblProd.FK_CatID) ON tblPack.PackID = tblProd.FK_PackID"
        .SQLWhere = "tblProd.Active=True"
        .SQLWhereFields = "tblProd.ProdID, tblProd.ProdCode, tblProd.ProdDescription, tblPack.PackTitle, tblCat.CatTitle, tblProd.SupPrice, tblProd.SRPrice"
        .SQLOrderBy = "tblProd.ProdDescription"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 2
        
        .AddColumn "ID", 100
        .AddColumn "Code", 100
        .AddColumn "Description", 180
        .AddColumn "Unit", 70
        .AddColumn "Category", 80
        .AddColumn "Sup. Price", 60, lgAlignRightCenter
        .AddColumn "SRP", 60, lgAlignRightCenter
        
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    isOn = False
End Sub



Private Sub SaveAdd()
        
    Dim lPTSID As Long
    Dim newPTS As tPTS
    Dim dNewAmount As Double
        
    'default
    
    lPTSID = -1
    
    
    'validate
    'reference number
    If Len(Trim(txtRefNum.Text)) < 1 Then
        MsgBox "Please enter 'Reference Number'.", vbExclamation
        HLTxt txtRefNum
        Exit Sub
    End If
    
    'supplier
    If Len(Trim(b8DPSup.BoundData)) < 1 Then
        MsgBox "Please enter 'Supplier'.", vbExclamation
        b8DPSup.FocusedDropButton
        Exit Sub
    End If
    
    'products
    If Not (GetTxtVal(txtTotalAmount.Text) > 0) Then
        MsgBox "Enter some purchased Product first.", vbExclamation
        b8DPProd.FocusedDropButton
        Exit Sub
    End If

  
    Select Case cmbCA.ListIndex
    
        Case 0 To 1 'full or partial

            Dim sRemarks As String
            sRemarks = "Payment for P.O. with Ref # " & Trim(txtRefNum.Text)

            If frmPTSEntry.ShowAdd(dtpPODate.Value, CLng(GetTxtVal(b8DPSup.BoundData)), cmbFP.Text, GetTxtVal(txtPayAmtOnDate.Text), GetTxtVal(txtTotalAmount.Text), sRemarks, lPTSID) = False Then
                Exit Sub
            End If
            
            If GetPTSByID(lPTSID, newPTS) = False Then
                WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'GetPTSByID(lPTSID, newPTS) = False'"
                Exit Sub
            End If
            
            
            'set new PO
            
            If newPTS.Amount >= FormatNumber(GetTxtVal(txtTotalAmount.Text), 2) Then
                'full
                newPO.CA = "full payment"
            Else
                'partial
                newPO.CA = "partial payment"
            End If

            With newPO
            
                .FP = newPTS.FP
                .OptFK_PTSID = lPTSID
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                
                If newPTS.FP = "check" Then
                    If newPTS.Cleared = True Then
                        dNewAmount = newPTS.Amount
                    Else
                        dNewAmount = 0
                    End If
                Else
                    dNewAmount = newPTS.Amount
                End If
                
                .PayAmtOnDate = dNewAmount
                
                .POBalance = FormatNumber(.TotalAmt - .PayAmtOnDate, 2)
                .Remarks = txtRemarks.Text
                
            End With
            

        Case 2 'not paid
            cmbFP.ListIndex = 2

            With newPO
                .FP = ""
                .CA = "not paid"
            
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                .PayAmtOnDate = 0
                
                .POBalance = .TotalAmt
    
                .Remarks = txtRemarks.Text
            End With
            
    End Select

    'set remaining new PO info
    With newPO
        .POID = CLng(GetTxtVal(txtPOID.Text))
        .RefNum = Trim(txtRefNum.Text)
        .FK_SupID = CLng(GetTxtVal(b8DPSup.BoundData))
        
        ' + 1 second
        .PODate = modFunction.GetRSec(dtpPODate.Value) + (1 / 86400)
                
        .RC = Now
        'RM
        .RCU = CurrentUser.UserID
        'RMU
    End With
    
    'write new PO
    If modRSPO.AddPO(newPO) = True Then
        
        'add PO Items(Products)
        Dim newPOProd As tPOProd
        Dim li As Long
        
        For li = 0 To listPOProd.RowCount - 1
            With newPOProd
                .FK_POID = newPO.POID
                .FK_ProdID = Val(listPOProd.CellText(li, 4))
                .Qty = GetTxtVal(listPOProd.CellText(li, 0))
                .InvQty = GetTxtVal(listPOProd.CellText(li, 1))
                .FK_PackID = GetTxtVal(listPOProd.CellText(li, 2))
                .UnitPrice = GetTxtVal(listPOProd.CellText(li, 6))
                .Amount = GetTxtVal(listPOProd.CellText(li, 7))
            End With
            
            If modRSPOProd.AddPOProd(newPOProd, newPO) = False Then
                WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSPOProd.AddPOProd(newPOProd, NewPO) = False'"
            End If

        Next
                
        'set flag
        mShowAdd = True
        'close this form
        Unload Me
        
    Else
    
        'delete saved pts
        If modRSPTS.DeletePTS(lPTSID) = False Then
            WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSPTS.DeletePTS(lPTSID) = False'"
        End If
        
        WriteErrorLog Me.Name, "SaveAdd", "Failed on: 'modRSPO.AddPO(newPO) = True'"
    End If
        
    
End Sub


Private Sub SaveEdit()
    
    Dim lPTSID As Long
    Dim curPTS As tPTS
    Dim dNewAmount As Double
    Dim tmpPTS As tPTS
    
    
    'default
    
    lPTSID = -1
    
    
    'validate
    'reference number
    If Len(Trim(txtRefNum.Text)) < 1 Then
        MsgBox "Please enter 'Reference Number'.", vbExclamation
        HLTxt txtRefNum
        Exit Sub
    End If
    
    'supplier
    If Len(Trim(b8DPSup.BoundData)) < 1 Then
        MsgBox "Please enter 'Supplier'.", vbExclamation
        b8DPSup.FocusedDropButton
        Exit Sub
    End If
    
    'products
    If Not (GetTxtVal(txtTotalAmount.Text) > 0) Then
        MsgBox "Enter some purchased Product first.", vbExclamation
        b8DPProd.FocusedDropButton
        Exit Sub
    End If

  
    Select Case cmbCA.ListIndex
    
        Case 0 To 1 'full or partial

            Dim sRemarks As String
            sRemarks = "Payment for P.O. with Ref # " & Trim(txtRefNum.Text)

            If GetPTSByID(curPO.OptFK_PTSID, tmpPTS) = True Then
                If frmPTSEntry.ShowEdit(curPO.OptFK_PTSID) = False Then
                    Exit Sub
                End If
            Else
                If frmPTSEntry.ShowAdd(dtpPODate.Value, CLng(GetTxtVal(b8DPSup.BoundData)), cmbFP.Text, GetTxtVal(txtPayAmtOnDate.Text), GetTxtVal(txtTotalAmount.Text), sRemarks, curPO.OptFK_PTSID) = False Then
                    Exit Sub
                End If
            End If
            
            
            If GetPTSByID(curPO.OptFK_PTSID, curPTS) = False Then
                WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'GetPTSByID(curPO.OptFK_PTSID, curPTS) = False'"
                Exit Sub
            End If
            
            
            'set new PO
            
            If curPTS.Amount >= FormatNumber(GetTxtVal(txtTotalAmount.Text), 2) Then
                'full
                curPO.CA = "full payment"
            Else
                'partial
                curPO.CA = "partial payment"
            End If

            With curPO
            
                .FP = curPTS.FP
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                
                If curPTS.FP = "check" Then
                    If curPTS.Cleared = True Then
                        dNewAmount = curPTS.Amount
                    Else
                        dNewAmount = 0
                    End If
                Else
                    dNewAmount = curPTS.Amount
                End If
                
                .PayAmtOnDate = dNewAmount
                
                .POBalance = FormatNumber(.TotalAmt - .PayAmtOnDate, 2)
                .Remarks = txtRemarks.Text
                
            End With
            

        Case 2 'not paid
            cmbFP.ListIndex = 2

            With curPO
                .FP = ""
                .CA = "not paid"
            
                .TotalAmt = GetTxtVal(txtTotalAmount.Text)
                .PayAmtOnDate = 0

                .POBalance = .TotalAmt
    
                .Remarks = txtRemarks.Text
            End With
            
            'Delete pts
            If GetPTSByID(curPO.OptFK_PTSID, tmpPTS) = True Then
                modRSPTS.DeletePTS curPO.OptFK_PTSID
            End If
            
    End Select

    'set remaining new PO info
    With curPO
        .POID = CLng(GetTxtVal(txtPOID.Text))
        .RefNum = Trim(txtRefNum.Text)
        .FK_SupID = CLng(GetTxtVal(b8DPSup.BoundData))
        
        ' + 1 second
        '.PODate = modFunction.GetRSec(dtpPODate.Value) + (1 / 86400)
                
        '.RC = Now
        .RM = Now
        '.RCU
        .RMU = CurrentUser.UserID
    End With
    
    'write new PO
    If modRSPO.EditPO(curPO) = True Then
        
        'add PO Items(Products)
        Dim curPOProd As tPOProd
        Dim li As Long
        
        'delete old po prod
        DeleteAllPOProd curPO.POID, curPO
        
        For li = 0 To listPOProd.RowCount - 1
            With curPOProd
                .FK_POID = curPO.POID
                .FK_ProdID = Val(listPOProd.CellText(li, 4))
                .Qty = GetTxtVal(listPOProd.CellText(li, 0))
                .InvQty = GetTxtVal(listPOProd.CellText(li, 1))
                .FK_PackID = GetTxtVal(listPOProd.CellText(li, 2))
                .UnitPrice = GetTxtVal(listPOProd.CellText(li, 6))
                .Amount = GetTxtVal(listPOProd.CellText(li, 7))
            End With
            
            If modRSPOProd.AddPOProd(curPOProd, curPO) = False Then
                WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSPOProd.AddPOProd(curPOProd, curPO) = False'"
            End If

        Next
                
        'set flag
        mShowEdit = True
        'close this form
        Unload Me
        
    Else
    
        'delete saved pts
        If modRSPTS.DeletePTS(curPO.OptFK_PTSID) = False Then
            WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSPTS.DeletePTS(lPTSID) = False'"
        End If
        
        WriteErrorLog Me.Name, "SaveEdit", "Failed on: 'modRSPO.AddPO(curPO) = True'"
    End If
    
End Sub









'---------------------------------------------------------------
'Supplier Info Procedures
'---------------------------------------------------------------
Private Sub Form_UseThisSup(ByVal lSupID As Long)
    
    Dim vSup As tSup
    
    txtAddress.Text = ""
    txtAP.Text = ""
    
    If modRSSup.GetSupByID(lSupID, vSup) = True Then
    
        b8DPSup.BoundData = vSup.SupID
        b8DPSup.DisplayData = vSup.SupName
        
        txtAddress.Text = vSup.Address
        
        
        'get balance
        txtAP.Text = FormatNumber(modRSAP.GetAPBySup(lSupID, CDate(0), dtpPODate.Value), 2)
  
    End If
    
End Sub


Private Sub b8DPProd_Change()

    If RefeshCurPOProd(CLng(GetTxtVal(b8DPProd.BoundData))) = False Then
        Exit Sub
    End If

    'set focused control
    HLTxt txtQty
    
End Sub

Private Sub b8DPSup_Change()
    
    Form_UseThisSup CLng(GetTxtVal(b8DPSup.BoundData))

End Sub

Private Sub cmdNewSup_Click()
    
    Dim lSupID As Long
    
    lSupID = frmSupEntry.ShowAddRetID
    
    If lSupID >= 0 Then
        Form_UseThisSup lSupID
    End If
End Sub


'---------------------------------------------------------------
'>>> END Supplier Info Procedures
'---------------------------------------------------------------





'---------------------------------------------------------------
'Product Info Procedures
'---------------------------------------------------------------
Private Sub LoadProducts(ByVal lPOID As Long)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long



    sSQL = "SELECT tblPOProd.Qty, tblPOProd.InvQty, tblProd.FK_PackID, tblPack.PackTitle, tblPOProd.FK_ProdID, tblProd.ProdDescription, tblPOProd.UnitPrice, tblPOProd.Amount, tblPOProd.FK_POID" & _
            " FROM tblPack INNER JOIN (tblProd INNER JOIN tblPOProd ON tblProd.ProdID = tblPOProd.FK_ProdID) ON tblPack.PackID = tblProd.FK_PackID" & _
            " WHERE tblPOProd.FK_POID=" & lPOID
    
    If ConnectRS(PrimeDB, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadProducts", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    listPOProd.Redraw = False
    listPOProd.clear

    vRS.MoveFirst
    While vRS.EOF = False
        
        With listPOProd
        li = .AddItem(CStr(ReadField(vRS.Fields("Qty"))))
        .ItemImage(li) = 1
        .CellText(li, 1) = ReadField(vRS.Fields("InvQty"))
        .CellText(li, 2) = ReadField(vRS.Fields("FK_PackID"))
        .CellText(li, 3) = ReadField(vRS.Fields("PackTitle"))
        .CellText(li, 4) = ReadField(vRS.Fields("FK_ProdID"))
        .CellText(li, 5) = ReadField(vRS.Fields("ProdDescription"))
        .CellText(li, 6) = ReadField(vRS.Fields("UnitPrice"))
        .CellText(li, 7) = ReadField(vRS.Fields("Amount"))
        End With
        
        vRS.MoveNext
    Wend
    
RAE:
    listPOProd.Redraw = True
    listPOProd.Refresh
    Set vRS = Nothing

End Sub

Private Sub cmdDeletePOProd_Click()
    If listPOProd.RowCount > 0 Then
        listPOProd.RemoveItem listPOProd.Row
    
        'calculate total amount
        Call Form_CalTotalAmount
    
    End If
End Sub


Private Function RefeshCurPOProd(ByVal lProdID As Long, Optional dQty As Double = 0, Optional lPackID As Long = 0, Optional dPrice As Double = 0) As Boolean
    
    Dim i As Integer
    Dim vProd As tProd
    
    'default
    RefeshCurPOProd = False
    
    'clear & Disable
    cmbPackTitle.clear
    txtQty.Text = dQty
    txtQtyPrice.Text = dPrice
    txtAmount.Text = ""
    
    txtQty.Enabled = False
    cmbPackTitle.Enabled = False
    txtQtyPrice.Enabled = False
    
    cmdAddPOProd.Enabled = False
    
    If GetProdByID(lProdID, vProd) = False Then
        Exit Function
    End If
    
    'fill packages
    If modRSProdPack.FillProdPackToTypeArray(lProdID, ProdPackList) = False Then
        WriteErrorLog Me.Name, "RefeshCurPOProd", "Failed on: 'modRSProdPack.FillProdPackToTypeArray(lProdID, prodpacklis) = False'"
        Exit Function
    End If
    If UBound(ProdPackList) >= 0 Then
        For i = 0 To UBound(ProdPackList)
            cmbPackTitle.AddItem ProdPackList(i).PackTitle
        Next
    Else
        Exit Function
    End If
    
    cmbPackTitle.Enabled = True
    'default package
    cmbPackTitle.ListIndex = 0
    'set current package base on parameter
    For i = 0 To UBound(ProdPackList)
        If ProdPackList(i).FK_PackID = lPackID Then
            cmbPackTitle.ListIndex = i
            Exit For
        End If
    Next
        
    txtQty.Enabled = True
    txtQtyPrice.Enabled = True
    
    'return sucess
    RefeshCurPOProd = True
    
End Function

Private Sub cmbPackTitle_Change()

    cmdAddPOProd.Enabled = False
    
    If cmbPackTitle.ListIndex >= 0 Then
        txtQtyPrice.Text = FormatNumber(ProdPackList(cmbPackTitle.ListIndex).SRPrice, 2)
    Else
        Exit Sub
    End If
    
    'generate Amount
    If GetTxtVal(txtQty.Text) > 0 Then
        txtAmount.Text = FormatNumber(GetTxtVal(txtQty.Text) * GetTxtVal(txtQtyPrice.Text), 2)
    Else
        Exit Sub
    End If
    
    If Not GetTxtVal(txtAmount.Text) > 0 Then
        Exit Sub
    End If
    
    'sucess
    'enable Add
    cmdAddPOProd.Enabled = True
    
End Sub

Private Sub cmbPackTitle_Click()
    Call cmbPackTitle_Change
End Sub

Private Sub listPOProd_DblClick()

    With listPOProd
    
    If .RowCount > 0 Then

        b8DPProd.BoundData = CLng(GetTxtVal(.CellText(.Row, 4)))
        b8DPProd.DisplayData = .CellText(.Row, 5)
        
        RefeshCurPOProd CLng(GetTxtVal(.CellText(.Row, 4))), GetTxtVal(.CellText(.Row, 0)), CLng(GetTxtVal(.CellText(.Row, 2))), _
                        GetTxtVal(.CellText(.Row, 6))
    End If
    
    End With
End Sub

Private Sub listPOProd_ItemCountChanged()

    If listPOProd.RowCount > 0 Then
        cmdDeletePOProd.Enabled = True
        cmdEditPOProd.Enabled = True
    Else
        cmdDeletePOProd.Enabled = False
        cmdEditPOProd.Enabled = False
        txtTotalAmount.Text = "0.00"
    End If
End Sub



Private Sub txtQty_Change()
    'generate amount
    Call cmbPackTitle_Change
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'add item
        Call cmdAddPOProd_Click
    End If
End Sub

Private Sub txtQtyPrice_Change()

    cmdAddPOProd.Enabled = False
    
    If Not (cmbPackTitle.ListIndex >= 0) Then
        Exit Sub
    End If
    
    'generate Amount
    If GetTxtVal(txtQty.Text) > 0 Then
        txtAmount.Text = FormatNumber(GetTxtVal(txtQty.Text) * GetTxtVal(txtQtyPrice.Text), 2)
    Else
        Exit Sub
    End If
    
    If Not GetTxtVal(txtAmount.Text) > 0 Then
        Exit Sub
    End If
    
    'sucess
    'enable Add
    cmdAddPOProd.Enabled = True
    
End Sub

Private Sub cmdAddPOProd_Click()

    Dim li As Long
    Dim lProdID As Long
    Dim dupli As Long
    
    'validate
    If Not (GetTxtVal(txtQty.Text) > 0) Then
        MsgBox "Please enter valid 'Quantity'", vbExclamation
        HLTxt txtQty
        Exit Sub
    End If
    
    If Not (GetTxtVal(txtQtyPrice.Text) > 0) Then
        MsgBox "Please enter valid 'Unit Price'", vbExclamation
        HLTxt txtQtyPrice
        Exit Sub
    End If
    
    'check if the product is already in the list
    lProdID = CLng(GetTxtVal(b8DPProd.BoundData))
    dupli = listPOProd.FindItem(CStr(lProdID), 4, lgSMEqual, False)
    
    If dupli >= 0 Then
        If MsgBox("The Product that you have added is already in the list." & vbNewLine & vbNewLine & _
            "Do you want to replace it?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            'the answer is YES
            'remove old
            listPOProd.RemoveItem dupli
        Else
            'the answer is NO
            Exit Sub
        End If
    End If
            
    
    With listPOProd
        .Redraw = False
        
        li = .AddItem(CStr(GetTxtVal(txtQty.Text)))
        .CellText(li, 1) = ProdPackList(cmbPackTitle.ListIndex).Qty * GetTxtVal(txtQty.Text)
        .CellText(li, 2) = ProdPackList(cmbPackTitle.ListIndex).FK_PackID
        .CellText(li, 3) = cmbPackTitle.Text
        .CellText(li, 4) = lProdID
        .CellText(li, 5) = b8DPProd.DisplayData
        .CellText(li, 6) = FormatNumber(GetTxtVal(txtQtyPrice.Text), 2)
        .CellText(li, 7) = FormatNumber(GetTxtVal(txtAmount.Text), 2)
        
        .Redraw = True
        .Refresh
    End With
    
    
    'calculate total amount
    Call Form_CalTotalAmount
    
    'clear & Disable
    cmbPackTitle.clear
    txtQty.Text = ""
    txtQtyPrice.Text = ""
    txtAmount.Text = ""
    
    b8DPProd.ClearCurData
    
    txtQty.Enabled = False
    cmbPackTitle.Enabled = False
    txtQtyPrice.Enabled = False
    
    cmdAddPOProd.Enabled = False
    
    'set focused on next control
    b8DPProd.FocusedDropButton
    

End Sub

Private Sub Form_CalTotalAmount()

    Dim li As Long
    Dim dTA As Double
    
    'clear
    txtPOBalance.Text = "0.00"
    
    dTA = 0
    For li = 0 To listPOProd.RowCount - 1
        dTA = dTA + GetTxtVal(listPOProd.CellText(li, 7))
    Next
    
    txtTotalAmount.Text = FormatNumber(dTA, 2)
    
    If GetTxtVal(txtPayAmtOnDate.Text) < 0 Then
        Exit Sub
    End If
    
    txtPOBalance.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text) - GetTxtVal(txtPayAmtOnDate.Text), 2)
    
End Sub
'---------------------------------------------------------------
' >>> END Product Info Procedures
'---------------------------------------------------------------



'---------------------------------------------------------------
'Payment Info Procedures
'---------------------------------------------------------------

Private Sub txtPayAmtOnDate_Change()

    txtPOBalance.Text = "0.00"
    
    If GetTxtVal(txtTotalAmount.Text) < 0 Then
        Exit Sub
    End If
    
    If GetTxtVal(txtPayAmtOnDate.Text) < 0 Then
        Exit Sub
    End If
    
    txtPOBalance.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text) - GetTxtVal(txtPayAmtOnDate.Text), 2)
    
End Sub

Private Sub cmbCA_Change()
    
    'disable affected controls
    cmbFP.Enabled = False
    
    Select Case cmbCA.ListIndex
        Case 0 'full
            cmbFP.Enabled = True
            'set FA to cash
            cmbFP.ListIndex = 0
        Case 1 'partial
            cmbFP.Enabled = True
            'set FA to cash
            cmbFP.ListIndex = 0
        Case 2 'no paid
            cmbFP.ListIndex = 2
            
    End Select
    
    
    
End Sub

Private Sub cmbCA_Click()
    
    Call cmbCA_Change
    
    Call CAFPChange
    
End Sub

Private Sub cmbFP_Change()
    Call CAFPChange
End Sub

Private Sub cmbFP_Click()
    Call CAFPChange
End Sub

Private Sub CAFPChange()

    Select Case cmbCA.ListIndex
        Case 0 'full
            Select Case cmbFP.ListIndex
                Case 0 'cash
                    txtPayAmtOnDate.Text = FormatNumber(GetTxtVal(txtTotalAmount.Text), 2)
                    txtPayAmtOnDate.Locked = True
                    
                Case 1 'check
                    txtPayAmtOnDate.Locked = True
                    
                Case 2 'other
                    txtPayAmtOnDate.Locked = False
                    
            End Select
            
            
        Case 1 'partial
            Select Case cmbFP.ListIndex
                Case 0 'cash
                    txtPayAmtOnDate.Locked = False
                Case 1 'check
                    txtPayAmtOnDate.Locked = True
                    
                Case 2 'other
                    txtPayAmtOnDate.Locked = False
            End Select
        Case 2 'not paid
            txtPayAmtOnDate.Locked = True
            txtPayAmtOnDate.Text = "0.00"
    End Select

End Sub

Private Sub txtTotalAmount_Change()
    CAFPChange
End Sub

Private Sub Form_RefreshCA(Optional sCA As String = "Full Payment")
    
    Dim i As Integer
    
    cmbCA.clear
    
    
    cmbCA.AddItem "Full Payment"
    cmbCA.AddItem "Partial Payment"
    cmbCA.AddItem "Not Paid"
    
    For i = 0 To cmbCA.ListCount - 1
        If LCase(Trim(cmbCA.List(i))) = LCase(Trim(sCA)) Then
            cmbCA.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_RefreshFP(Optional sFP As String = "Cash")
    
    cmbFP.clear
    
    cmbFP.AddItem "Cash"
    cmbFP.AddItem "Check"
    cmbFP.AddItem "Other"

    Dim i As Integer
    

    For i = 0 To cmbFP.ListCount - 1
        If LCase(Trim(cmbFP.List(i))) = LCase(Trim(sFP)) Then
            cmbFP.ListIndex = i
            Exit Sub
        End If
    Next

    cmbFP.ListIndex = 2
End Sub

'---------------------------------------------------------------
'Supplier Info Procedures
'---------------------------------------------------------------

