VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOS 
   Caption         =   "POS "
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   600
      Top             =   1800
   End
   Begin VB.TextBox TXT_RATE 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TXT_AMT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TXT_QTY 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4245
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TXT_CODE 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   465
      Left            =   240
      TabIndex        =   16
      Top             =   2220
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   820
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632256
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S.N"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ITEM CODE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ITEM DESCRIPTION"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox PIC_RIGHT 
      Align           =   4  'Align Right
      BackColor       =   &H00400000&
      ForeColor       =   &H00000000&
      Height          =   4380
      Left            =   8265
      ScaleHeight     =   4320
      ScaleWidth      =   3555
      TabIndex        =   9
      Top             =   615
      Width           =   3615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmPOS.frx":0000
         Height          =   2775
         Left            =   0
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483644
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   16
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblerror 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   885
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   3405
      End
      Begin VB.Label lblinvo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1680
         TabIndex        =   37
         Top             =   0
         Width           =   75
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblcat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblunit 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   795
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label lblselling 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C7BDAD&
         Height          =   495
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT IN STOCK:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   390
         Width           =   2370
      End
   End
   Begin VB.PictureBox PIC_TOP 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11880
      TabIndex        =   5
      Top             =   0
      Width           =   11880
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   10560
         Top             =   120
      End
      Begin VB.Label lbl_date 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   240
         Left            =   6600
         TabIndex        =   25
         Top             =   240
         Width           =   2340
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B-MORADA STREET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIPA SOLID AUTO SUPPLY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2670
      End
   End
   Begin VB.PictureBox PIC_BOTTOM 
      Align           =   2  'Align Bottom
      FillColor       =   &H8000000E&
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   11820
      TabIndex        =   4
      Top             =   4995
      Width           =   11880
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame frmLabels 
         Height          =   2055
         Left            =   8160
         TabIndex        =   11
         Top             =   0
         Width           =   5895
         Begin MSMask.MaskEdBox TextAmount 
            Height          =   570
            Left            =   2475
            TabIndex        =   14
            Top             =   180
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1005
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   4194304
            ForeColor       =   1309691
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TextChange 
            Height          =   570
            Left            =   2475
            TabIndex        =   15
            Top             =   1395
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1005
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   4194304
            ForeColor       =   1309691
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TextTend 
            Height          =   570
            Left            =   2475
            TabIndex        =   23
            Top             =   795
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1005
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   4194304
            ForeColor       =   1309691
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Tendered :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   45
            TabIndex        =   24
            Top             =   938
            Width           =   2370
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   540
            TabIndex        =   13
            Top             =   323
            Width           =   1845
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1305
            TabIndex        =   12
            Top             =   1538
            Width           =   1080
         End
      End
      Begin MSForms.CommandButton CMDCANCEL 
         Height          =   375
         Left            =   6720
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
         ForeColor       =   16777215
         BackColor       =   4210752
         Caption         =   "CANCEL"
         Size            =   "1931;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CMDOK 
         Height          =   375
         Left            =   5520
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
         ForeColor       =   16777215
         BackColor       =   4210752
         Caption         =   "OK"
         Size            =   "1931;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label LBLCUST 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   38
         Top             =   600
         Width           =   2415
      End
      Begin MSForms.CommandButton CMDCRED 
         Height          =   1215
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1335
         ForeColor       =   -2147483634
         BackColor       =   4210752
         VariousPropertyBits=   19
         Size            =   "2355;2143"
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMD_CLEAR 
         Height          =   1215
         Left            =   1440
         TabIndex        =   33
         Top             =   120
         Width           =   1335
         ForeColor       =   -2147483634
         BackColor       =   4210752
         VariousPropertyBits=   19
         Size            =   "2355;2143"
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMD_LOG 
         Height          =   1215
         Left            =   2760
         TabIndex        =   32
         Top             =   120
         Width           =   1335
         ForeColor       =   -2147483634
         BackColor       =   4210752
         VariousPropertyBits=   19
         Size            =   "2355;2143"
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMD_EXIT 
         Height          =   1215
         Left            =   4080
         TabIndex        =   31
         Top             =   120
         Width           =   1335
         ForeColor       =   -2147483634
         BackColor       =   4210752
         VariousPropertyBits=   19
         Size            =   "2355;2143"
         FontName        =   "Trebuchet MS"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIPA SOLID AUTO SUPPLY"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0013FBFB&
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   3105
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Credit Customer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "      CLEAR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1440
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   44
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   46
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "          Log Off"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   2760
         TabIndex        =   45
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   48
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "         Exit"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   4080
         TabIndex        =   47
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6840
      TabIndex        =   27
      Top             =   765
      Width           =   795
   End
   Begin VB.Label LBL_DES 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C7BDAD&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5535
      TabIndex        =   18
      Top             =   765
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4260
      TabIndex        =   17
      Top             =   765
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   0
      X2              =   500
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ENCODE PRODUCT NO/NAME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   8
      Top             =   765
      Width           =   2655
   End
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Beep 2000, 55
Dim TXTLEN As Integer
Dim STRT As Integer
Dim PRILEN As Integer


Private Sub CMD_EXIT_Click()
Call BeepEntry

Unload Me
End Sub





Private Sub CMD_CLEAR_Click()
Call BeepEntry





ListView1.ListItems.Clear
TXT_CODE.Text = ""
TXT_AMT.Text = ""
TXT_RATE.Text = ""
TXT_QTY.Text = ""
lblinvo.Caption = ""
TextTend.Text = ""
TextAmount.Text = ""
TextChange.Text = ""
ID = ""
TXT_CODE.SetFocus


End Sub

Private Sub CMD_HOME_Click()

End Sub

Private Sub CMD_LOG_Click()
Call BeepEntry

MsgBox "LOGOFF"
End Sub

Private Sub CMDCANCEL_Click()
Combo1.Visible = False
CMDOK.Visible = False
CMDCANCEL.Visible = False
CMDCRED.Enabled = True
CMD_CLEAR.Enabled = True
End Sub

Private Sub CMDCRED_Click()
Call BeepEntry

'ListView1.ListItems.Clear

Combo1.Visible = True
CMDOK.Visible = True
CMDCANCEL.Visible = True
CMDCRED.Enabled = False
CMD_CLEAR.Enabled = False

Call CMB1("CUSTOMER", "CUST_ID", Combo1)


End Sub

Private Sub CMDOK_Click()
Combo1.Visible = False
CMDOK.Visible = False
CMDCANCEL.Visible = False
CMDCRED.Enabled = True
CMD_CLEAR.Enabled = True
TXT_CODE.Locked = False
TXT_QTY.Locked = False
TXT_CODE.SetFocus
CRED = True

End Sub

Private Sub Combo1_Click()
On Error Resume Next

Call DBCONNECT

Set RS1 = New Recordset

RS1.Open "Select * from CATEGORY where CAT_ID='" _
+ Combo1.Text + "'", DCON, adOpenDynamic, adLockPessimistic


LBLCUST.Caption = RS1!CAT_NAME

Set RS1 = Nothing
Set DCON = Nothing
End Sub

Private Sub CommandButton1_Click()




End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call CMDOK_Click

End If

End Sub

Private Sub Command1_Click()

End Sub

'Private Sub Command1_Click()
'Call DBCONNECT
'
'SQL = "Select * from invoice where LEFT(date,8)='" & Left(Now, 8) & "'"
'
'
'Set RS1 = New Recordset
'Set RS1 = DCON.Execute(SQL)
'
'
'
'MsgBox "NO. OF TRANSACTION/S TODAY IS " & RS1.RecordCount & vbTab, vbInformation
'
'
'
'
'Set DCON = Nothing
'
'End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)


If DataGrid1.Row >= 0 Then

If DataGrid1.Text <> "" Then


    TXT_CODE.Text = DataGrid1.Text
   LBL_DES.Caption = DataGrid1.Text & "," & DataGrid1.Columns(1).Text
   TXT_QTY.SetFocus
   

End If
End If




End Sub

Private Sub Form_Load()
   Call DBCONNECT
  
    DataGrid1.Visible = True
    
    Set RS1 = New Recordset
    
    Set RS1 = DCON.Execute("Select TOP 5 * from Product")
    Set DataGrid1.DataSource = RS1
  
    Set RS1 = Nothing
    Set DCON = Nothing
     
    CRED = False
    
    
End Sub

Private Sub CMD_SAVE_Click()

End Sub


Private Sub Form_Resize()
    
    
    Line1.X2 = Me.Width
    frmLabels.Left = Me.ScaleWidth - frmLabels.Width - 400 '- CMD_SAVE.Width - 200
    ListView1.Move Me.ScaleLeft + 240, ListView1.Top, PIC_RIGHT.Left - 340
    DataGrid1.Move PIC_RIGHT.ScaleLeft, PIC_RIGHT.ScaleTop + 2500, PIC_RIGHT.ScaleWidth, PIC_RIGHT.ScaleHeight
    lbl_date.Move PIC_RIGHT.Left
    lbl_date.BackStyle = 0
    lblerror.Top = DataGrid1.Top + DataGrid1.Height + 50
    
    ListView1.ColumnHeaders(1).Width = ListView1.Width * 0.1
    ListView1.ColumnHeaders(2).Width = ListView1.Width * 0.15
    ListView1.ColumnHeaders(3).Width = ListView1.Width * 0.3
    ListView1.ColumnHeaders(5).Width = ListView1.Width * 0.15
    ListView1.ColumnHeaders(4).Width = ListView1.Width * 0.1
    
End Sub


Sub LISTVIEW_HEIGHT()

If ListView1.Height + ListView1.Top + PIC_TOP.Height >= PIC_BOTTOM.Top Then Exit Sub
ListView1.Height = ListView1.Height + 350

End Sub



Private Sub Form_Unload(Cancel As Integer)

If ListView1.ListItems.Count >= 0 Then
    If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbYes Then
        End
    Else
        Cancel = True
        
    End If
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Call LISTVIEW_HEIGHT
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
EDT = True
If ListView1.ListItems.Count <> 0 Then
      TXT_CODE.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
      
       TXT_QTY.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3)
    TXT_RATE.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(4)
     TXT_AMT.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(5)
End If


End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DD As Integer
With ListView1
If .ListItems.Count <> 0 Then
If KeyCode = vbKeyDelete Then
TextAmount.Text = ""
       
  .ListItems.Remove (.SelectedItem.Index)
  
TXT_CODE.Text = ""

  lblcat.Caption = ""
EDT = False
TXT_RATE.Text = ""
TXT_AMT.Text = ""
TXT_QTY.Text = ""
lblselling.Caption = ""
lblunit.Caption = ""

            For DD = 1 To .ListItems.Count
                  TextAmount.Text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.Text)
            Next
  End If
End If
End With
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TXT_CODE.SetFocus
End If

End Sub

Private Sub Text1_Change()

End Sub

Private Sub TextChange_Change()
    lblMain.Caption = NumToWord(TextChange)
End Sub

Private Sub TextTend_Change()


    TextChange.Text = Val(TextTend.Text) - Val(TextAmount.Text)
 
End Sub

Sub BeepEntry()
Dim i As Integer
For i = 1 To 2
    Beep i * 500, 100
Next i
End Sub
Sub BeepError()
Dim i As Integer
For i = 4 To 1 Step -1
    Beep i * 80, 50
Next i
    Beep 200, 300
End Sub



Private Sub TextTend_KeyPress(KeyAscii As Integer)
Dim CNT1 As Integer

Call Decimals(KeyAscii, TextTend, 2)
 
If KeyAscii = 13 Then

With ListView1
If .ListItems.Count > 0 Then
Call DBCONNECT
   
    If CRED = True Then
        'SalesRegistryHeader
        'SalesRegistryDetail
        
        SQL = "Insert into SalesRegistryDetail('" & ID & "','" _
                                                  & Combo1.Text & "','" _
                                                  & Now & "','0')"
                        
        CRED = False

      
        DCON.Execute SQL
     
    Else
            SQL = "Insert into invoice values('" + ID + "','0','" _
                        & Now & "','0')"
                        
    
        DCON.Execute SQL
      
        
    End If

        For CNT1 = 1 To .ListItems.Count
            
               
        SQL = "Insert into item_invoice values('" + .ListItems(CNT1).SubItems(1) + "','" _
        + ID + "'," & .ListItems(CNT1).SubItems(3) & "," & .ListItems(CNT1).SubItems(4) & ")"
        DCON.Execute SQL

    Set RS1 = New Recordset
              
        SQL = "Select * from Product where ProductID='" & .ListItems(CNT1).SubItems(1) & "'"
       
       Set RS1 = DCON.Execute(SQL)
       
       SQL = "update Product set UnitsInStock=" & Val(Val(RS1!UnitsInStock) - Val(.ListItems(CNT1).SubItems(3))) _
                    & " WHERE ProductID='" & .ListItems(CNT1).SubItems(1) & "'"
                    
       DCON.Execute SQL
    
        
        Next
       
        
Set RS1 = Nothing
Set DCON = Nothing

End If
ListView1.ListItems.Clear
TXT_CODE.Text = ""
TXT_AMT.Text = ""
TXT_RATE.Text = ""
TXT_QTY.Text = ""
lblinvo.Caption = ""
TextTend.Text = ""
TextAmount.Text = ""
TextChange.Text = ""
ID = ""
TXT_CODE.SetFocus

End With

End If


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

With ListView1
If .ListItems.Count > 0 Then



         Select Case Shift & KeyCode
    
    Case vbAltMask & vbKeyF4
                Unload Me

    End Select

    Select Case KeyCode
    
    
    
    Case vbKeyDown
        EDT = True
        
      
        .SetFocus
   
        TXT_CODE.Text = .ListItems(.SelectedItem.Index).SubItems(1)
        TXT_QTY.Text = .ListItems(.SelectedItem.Index).SubItems(3)
        TXT_RATE.Text = .ListItems(.SelectedItem.Index).SubItems(4)
        TXT_AMT.Text = .ListItems(.SelectedItem.Index).SubItems(5)
      
        
    
    Case vbKeyUp
        EDT = True
        
      
       .SetFocus
     
        TXT_CODE.Text = .ListItems(.SelectedItem.Index).SubItems(1)
        TXT_QTY.Text = .ListItems(.SelectedItem.Index).SubItems(3)
        TXT_RATE.Text = .ListItems(.SelectedItem.Index).SubItems(4)
        TXT_AMT.Text = .ListItems(.SelectedItem.Index).SubItems(5)
       
     Case vbKeyHome
        EDT = True
        
      
       .SetFocus
       
        TXT_CODE.Text = .ListItems(.SelectedItem.Index).SubItems(1)
        TXT_QTY.Text = .ListItems(.SelectedItem.Index).SubItems(3)
        TXT_RATE.Text = .ListItems(.SelectedItem.Index).SubItems(4)
        TXT_AMT.Text = .ListItems(.SelectedItem.Index).SubItems(5)
        
     Case vbKeyEnd
        EDT = True
        
      
      .SetFocus
       .ListItems(.ListItems.Count).Selected = True
        
        TXT_CODE.Text = .ListItems(.SelectedItem.Index).SubItems(1)
        TXT_QTY.Text = .ListItems(.SelectedItem.Index).SubItems(3)
        TXT_RATE.Text = .ListItems(.SelectedItem.Index).SubItems(4)
        TXT_AMT.Text = .ListItems(.SelectedItem.Index).SubItems(5)
    End Select
End If

    Select Case KeyCode
    
    Case vbKeyEscape
        Call CMDCANCEL_Click
        
    
    Case vbKeyF2
        Call CMDCRED_Click
    Case vbKeyF3
        Call CMD_HOME_Click
    Case vbKeyF4
       Call CMD_LOG_Click
    Case vbKeyF5
        Call CMD_EXIT_Click
        
        
    End Select
   




End With
End Sub

Private Sub Timer1_Timer()
lbl_date.Caption = "Date:" & Now
End Sub

Private Sub Timer2_Timer()
'Static c As Integer


STRT = STRT + 1



If STRT = 1 Then
Timer2.Interval = 0
  




Dim FVAL As String
Dim DD As Integer
DataGrid1.Visible = True
Call DBCONNECT

Set RS1 = New Recordset

SQL = "Select TOP 5 * from PRODUCT where PRODUCTID like '" & TXT_CODE & "%' OR NAME like'" & TXT_CODE & "%'"
Set RS1 = DCON.Execute(SQL)
Set DataGrid1.DataSource = RS1

If TXT_CODE.Text <> "" Then

    If Not RS1.EOF Then
'        TXT_CODE.SelStart = PRILEN
'        TXT_CODE.Text = RS1!Name
'        TXT_CODE.SelLength = Len(TXT_CODE.Text)
'
      
        FVAL = RS1!ProductID
        TXT_RATE.Text = RS1!UnitSellingPrice
        lblselling.Caption = RS1!UnitSellingPrice
        lblunit.Caption = RS1!UnitsInStock

  With ListView1
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                  
                If InStr(1, .ListItems(DD).SubItems(1), RS1!ProductID) = 1 Then
                  
                 
                            .ListItems(DD).Selected = True
                            lblunit.Caption = Val(lblunit.Caption) - Val(.ListItems(DD).SubItems(3))
                    
                    
                End If
            
               
            Next
         End If
    End With

        ' SQL = "Select Product.*,subcategory.*,category.* from " _
                + "Product,subcategory,category where " _
                + "Product.Product_id='" & FVAL & "' AND " _
                + "Product.sub_cat_id=subcategory.sub_cat_id AND " _
                + "subcategory.cat_id=category.cat_id"
        

       ' Set RS1 = DCON.Execute(SQL)

       ' If Not RS1.EOF Then

      '  lblcat.Caption = RS1!CAT_NAME
      '  End If


    Else
        TXT_RATE.Text = ""
        TXT_AMT.Text = ""
        TXT_QTY.Text = ""
        lblselling.Caption = ""
        lblunit.Caption = ""
        lblcat.Caption = ""
        PRILEN = 0

    End If


    Set RS1 = Nothing
    Set DCON = Nothing


ElseIf TXT_CODE.Text = "" Then
lblcat.Caption = ""
EDT = False
TXT_RATE.Text = ""
TXT_AMT.Text = ""
TXT_QTY.Text = ""
lblselling.Caption = ""
lblunit.Caption = ""
TXT_CODE.SetFocus
End If

End If

End Sub

Private Sub TXT_AMT_GotFocus()
If TXT_QTY.Text <> "" Then
    Call DBCONNECT

    Set RS1 = New Recordset
    SQL = "Select * from PRODUCT where PRODUCTid like '" & TXT_CODE & "%' OR NAME like'" & TXT_CODE & "%'"
    
    Set RS1 = DCON.Execute(SQL)
    If Not RS1.EOF Then

    'SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(TXT_RATE.Text) & " where PRODUCTID='" & RS1!ProductID & "'"
    SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(TXT_RATE.Text) & " where (PRODUCTID='" & RS1!ProductID & "' AND UnitCostPrice <" & Val(TXT_RATE.Text) & ")"
    
    DCON.Execute SQL
    End If
    
    Set RS1 = Nothing
    Set DCON = Nothing
    
  TXT_AMT = Val(TXT_RATE) * Val(TXT_QTY)
End If
End Sub

Private Sub TXT_AMT_KeyPress(KeyAscii As Integer)


Dim CNT As Boolean
Dim DD As Integer

If KeyAscii = 13 Then

If TXT_AMT.Text <> "" Then
   
    CNT = False
  

    Call DBCONNECT
        Set RS1 = New Recordset
      RS1.Open "Select * from Product where ProductID like'" & TXT_CODE & "%' OR Name like'" & TXT_CODE & "%'", DCON, adOpenForwardOnly, adLockReadOnly

If Not RS1.EOF Then
    
       Call LISTVIEW_HEIGHT
    
       LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name & ""
       TXT_RATE = RS1!UnitSellingPrice
            
  With ListView1
        TextAmount.Text = ""
        If .ListItems.Count <> 0 Then
          
            For DD = 1 To .ListItems.Count
                  
                If InStr(1, .ListItems(DD).SubItems(1), RS1!ProductID) = 1 Then
                    If InStr(1, .ListItems(DD).SubItems(2), RS1!Name) = 1 Then
                 
                        If EDT = True Then
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(3) = Val(TXT_QTY.Text)
                            .ListItems(DD).SubItems(4) = Val(TXT_RATE.Text)
                            .ListItems(DD).SubItems(5) = TXT_AMT.Text
                        Else
                            .ListItems(DD).Selected = True
                            .ListItems(DD).SubItems(3) = Val(.ListItems(DD).SubItems(3)) + Val(TXT_QTY.Text)
                            .ListItems(DD).SubItems(4) = Val(TXT_RATE.Text)
                            .ListItems(DD).SubItems(5) = Val(.ListItems(DD).SubItems(3)) * Val(.ListItems(DD).SubItems(4))
                    
                        End If
                    End If
                             
                    CNT = True
                    
                End If
                    
                   
            Next
       
         End If
            
        If CNT = False Then
                        
            .ListItems.Add 1, , .ListItems.Count + 1
            .ListItems(.ListItems.Count).SubItems(1) = RS1!ProductID
            .ListItems(.ListItems.Count).SubItems(2) = RS1!Name
            .ListItems(.ListItems.Count).SubItems(3) = TXT_QTY.Text
            .ListItems(.ListItems.Count).SubItems(4) = TXT_RATE.Text
            .ListItems(.ListItems.Count).SubItems(5) = TXT_AMT.Text
             ' TextAmount.Text = Val(TextAmount.Text) + Val(TXT_AMT.Text)
        
        End If
           
            
             
        
        If .ListItems.Count <= 0 Then
          
            .ListItems.Add 1, , .ListItems.Count + 1
            .ListItems(.ListItems.Count).SubItems(1) = RS1!ProductID
            .ListItems(.ListItems.Count).SubItems(2) = RS1!Name
            .ListItems(.ListItems.Count).SubItems(3) = TXT_QTY.Text
            .ListItems(.ListItems.Count).SubItems(4) = TXT_RATE.Text
            .ListItems(.ListItems.Count).SubItems(5) = TXT_AMT.Text
            
            
        End If


            For DD = 1 To .ListItems.Count
                  TextAmount.Text = Val(.ListItems(DD).SubItems(5)) + Val(TextAmount.Text)
            Next
            
         lblunit.Caption = RS1!UnitsInStock
       
        
       Set RS1 = DCON.Execute("Select * from Product")
        Set DataGrid1.DataSource = RS1
         
        Set RS1 = Nothing
        Set DCON = Nothing
   
  End With
    
End If
    
TXT_CODE.Text = ""
TXT_AMT.Text = ""
TXT_QTY.Text = ""
TXT_RATE.Text = ""
EDT = False
TXT_CODE.SetFocus

Else



TXT_QTY.SetFocus


End If



End If
End Sub

Private Sub TXT_CODE_Change()


Timer2.Interval = 100


TXTLEN = Len(TXT_CODE.Text)
STRT = 0




'Exit Sub
'adder:
'Exit Sub
End Sub

Private Sub TXT_CODE1_Change()


End Sub

Private Sub TXT_CODE_GotFocus()
If TXT_CODE.Text = "" Then
LBL_DES.Caption = ""
End If


End Sub

Private Sub TXT_CODE_KeyPress(KeyAscii As Integer)
 PRILEN = PRILEN + 1
    If KeyAscii = 13 Then
    
    If TXT_CODE.Text <> "" Then
        Call DBCONNECT
       
    Set RS1 = New Recordset
        RS1.Open "Select * from Product where ProductID like'" & TXT_CODE & "%' OR Name like'" & TXT_CODE & "%'", DCON, adOpenDynamic, adLockOptimistic

        If Not RS1.EOF Then
        
        LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name
        TXT_RATE = RS1!UnitSellingPrice
        TXT_QTY.SetFocus
        End If
        
        
        Set RS1 = Nothing
        Set DCON = Nothing
    Else
    
        
      
    End If
    
    End If


End Sub

Private Sub TXT_CODE_LostFocus()
        LBL_DES.Visible = True
        
End Sub

Private Sub TXT_QTY_GotFocus()



 
If TXT_CODE.Text <> "" Then
      Call DBCONNECT
      If ID = "" Then
            Call AUTONUM(DCON, "PRODUCT", "ProductID", "PROD", lblinvo)
       End If
        Set RS1 = New Recordset
        RS1.Open "Select * from Product where ProductID like'" & TXT_CODE & "%' OR name like'" & TXT_CODE & "%'", DCON, adOpenDynamic, adLockOptimistic

        If Not RS1.EOF Then

        LBL_DES.Caption = RS1!ProductID & ", " & RS1!Name
        'TXT_RATE = RS1!UnitSellingPrice
        
        If Val(lblunit.Caption) <= 0 Then
            'MsgBox "Out of Stock  ", vbInformation
            lblerror.Caption = "Out of Stock"
            'TXT_CODE.SetFocus
        
        Else
         If Val(RS1!ReorderLevel) >= Val(lblunit.Caption) And Val(RS1!ReorderQuantity) >= Val(lblunit.Caption) Then
          'MsgBox "You have reached the minimum stock and reorder level of this Product. ", vbInformation
            lblerror.Caption = "You have reached the minimum stock and reorder level of this Product. "
            
         ElseIf Val(RS1!ReorderLevel) >= Val(lblunit.Caption) Then
          'MsgBox "You have reached the reorder level of this item. ", vbInformation
            lblerror.Caption = "You have reached the reorder level of this item. "
      ElseIf Val(RS1!ReorderQuantity) >= Val(lblunit.Caption) Then
          ' MsgBox "You have reached the minimum stock of this item. ", vbInformation
            lblerror.Caption = "You have reached the minimum stock of this item. "
       End If
            End If
            
         End If
End If
  

            
            
     


        Set RS1 = Nothing
        Set DCON = Nothing



End Sub

Private Sub TXT_QTY_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TXT_QTY.Text <> "" Then
        
        If Val(TXT_QTY.Text) > Val(lblunit.Caption) Then
            'MsgBox "The quantity is greater than the unit of stock.  ", vbInformation
            lblerror.Caption = "The quantity is greater than the unit of stock.  "
            
            TXT_QTY.SetFocus
            
     
        Else
        
            TXT_AMT = Val(TXT_RATE) * Val(TXT_QTY)
      

            TXT_AMT.SetFocus
        End If
 
    End If
    
'Call ON_OFF_DEC(KeyAscii, TXT_QTY)
End Sub



Private Sub TXT_RATE_KeyPress(KeyAscii As Integer)
'Call ON_OFF_DEC(KeyAscii, TXT_RATE)

If KeyAscii = 13 Then

Call DBCONNECT

Set RS1 = New Recordset
SQL = "Select TOP 5 * from PRODUCT where PRODUCTID like '" & TXT_CODE & "%' OR NAME like'" & TXT_CODE & "%'"

Set RS1 = DCON.Execute(SQL)




    SQL = "UPDATE PRODUCT set UnitSellingPrice=" & Val(TXT_RATE.Text) & " where (PRODUCTID='" & RS1!ProductID & "' AND UnitCostPrice <" & Val(TXT_RATE.Text) & ")"
    MsgBox SQL
    DCON.Execute SQL
  

Set RS1 = Nothing
Set DCON = Nothing


TXT_AMT.SetFocus
End If

End Sub
