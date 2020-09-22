VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3345
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   6960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3345
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   2760
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   4800
      Top             =   2760
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Benjamin Artic s.p. Žalec SLO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TEK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PRO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents cmdObject As Image
Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long
Private Sub Form_Load()
'Me.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub Timer2_Timer()
Label5.Visible = True
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label6.Visible = True
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Label8.Visible = True
Timer5.Enabled = True
Timer4.Enabled = False
End Sub
Public Function ComputerName() As String
  Dim sBuffer As String
  
  Dim lAns As Long
 
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        err.Raise err.LastDllError, , _
          "A system call returned an error code of " _
           & err.LastDllError
   End If

End Function

Private Sub Timer5_Timer()
'Label9.Visible = Not Label9.Visible
tiskdol = Getnumb("select termi from lokal")
'Pblagajna = UPORABNIK
Pblagajna = stblagg()
If nivo = 1 Then
frmMAIN.Show
Else
If UCase(Left(UPORABNIK, 6)) = "POCKET" Then
'If RTrim(LTrim(UCase(ComputerName))) = "BENCI" Then

pocket.Show
stalnaprij = 1
Else

frmsalesbill.Show
End If
End If
Unload Me
End Sub
