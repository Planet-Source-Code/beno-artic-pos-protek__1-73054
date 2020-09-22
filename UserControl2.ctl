VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl UserControl2 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ScaleHeight     =   1020
   ScaleWidth      =   5760
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub opentime()
Timer1.Enabled = True
Label3.Caption = Now
'Format(Now, "hh:mm:ss")
End Sub

Public Sub closetime()
Timer1.Enabled = False

End Sub
Private Sub Timer1_Timer()
On Error GoTo nnn:
Dim XX As Long
XX = 1
Dim yyy As Long
Dim xxx As Long
If Xvs <> Yvs Then
If XX <> Yvs Then
yyy = DateDiff("s", CDate(Label3.Caption), Now)
End If
'Xvs = 31
ProgressBar1.Value = Yvs / Xvs * 100
xxx = ((Yvs / Xvs)) * 100

Label1.Caption = Trim(str(Yvs)) & " od " & Trim(str(Xvs)) & " zapisov"
'Label1.Caption = str(yyy) & " - " & str(xxx)
Label2.Caption = DateAdd("s", 100 * yyy / xxx, CDate(Label3.Caption))
'CDate(Label3.Caption))
XX = Yvs
End If
nnn:
End Sub

