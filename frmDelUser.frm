VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDelUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brisi uporabnika"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Dodaj"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Uredi"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Brisi"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
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
         DataField       =   "username1"
         Caption         =   "Users"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
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
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "frmDelUser.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EEECE8&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   3375
      Left            =   120
      Top             =   720
      Width           =   3450
   End
End
Attribute VB_Name = "frmDelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DataGrid1.ApproxCount > 0 Then
If DataGrid1.Columns(0).text <> "" Then

Call GetNewConnection2


DCON.Execute "Delete * from users where username1='" & DataGrid1.Columns(0).text & "'"




Set RS5 = DCON.Execute("Select username1 from users where username1 <> '" & CurUser & "'")


Set DataGrid1.DataSource = RS5

Set RS5 = Nothing
Set DCON = Nothing

End If
End If

End Sub

Private Sub Command2_Click()
frmChange.Show vbModal
End Sub

Private Sub Command3_Click()
frmAddUser.Show vbModal
End Sub

Private Sub Form_Load()
Call GetNewConnection2
Set RS5 = DCON.Execute(SQL)

Set DataGrid1.DataSource = RS5

Set RS5 = Nothing
Set DCON = Nothing



End Sub

Private Sub Text1_Change()
Call GetNewConnection2
Set RS5 = DCON.Execute("Select username1 from users where username1 <> '" & CurUser & "' AND username1 like '" & Text1.text & "%'")

Set DataGrid1.DataSource = RS5

Set RS5 = Nothing
Set DCON = Nothing

End Sub
