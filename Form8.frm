VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form8"
   ScaleHeight     =   6540
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   5640
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      GridColorUnpopulated=   16711935
      HighLight       =   2
      GridLinesFixed  =   3
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   3
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   
     SQL = "Select top 100 madasifr,madanazi,madampcd,madazalo from mada"
     CatalogueName = "Category"
    

Call GetNewConnection2
Set Rs1 = New Recordset
Set Rs1 = DCON.Execute(SQL)
    Set Form8.MSHFlexGrid1.DataSource = Rs1
Me.MSHFlexGrid1.CellBackColor = Rs1.Fields(2)
Me.MSHFlexGrid1.Refresh
Dim i, a As Integer
i = 3
a = 1

'With Me.MSHFlexGrid1
'For a = 1 To .Rows - 1
'.Col = i
'.Row = a
'.text = Round(.text, 4)
'If .text < 20 Then
'For X = 1 To .Cols - 1
'.Col = X
'.CellBackColor = 255
'Next
'End If

'Next
'End With

Set Rs1 = Nothing
Set DCON = Nothing

End Sub




