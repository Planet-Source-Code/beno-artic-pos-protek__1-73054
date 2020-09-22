VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form vnoscen 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   10380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   1215
      Left            =   4920
      TabIndex        =   12
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "Zapiši"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "vnoscen.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   1095
      Left            =   7440
      TabIndex        =   13
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "Preklici"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632319
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "vnoscen.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kol:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MPC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Marza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rabat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fakturna vrednost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cena"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "vnoscen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private iden As String
Private xdxd As String
Private xpxp As String
Private akoli As Double
Public Sub cene(dokument As String, pozicija As String)
'frmblag.WindowState = 1
xdxd = Trim(dokument)
xpxp = pozicija

iden = Getnazi("select sifra from trenutna where tip_dok+id_dok='" & dokument & "' and pozicija='" & pozicija & "'")
akoli = Getnumb("select kol from trenutna where tip_dok+id_dok='" & dokument & "' and pozicija='" & pozicija & "'")
Me.Label7.Caption = Getnazi("select madanazi from mada where madasifr='" & iden & "'")
Me.Label8.Caption = FormatNumber(akoli, 2)
Me.Text3.Text = FormatNumber(Getnumb("select pop from trenutna where tip_dok+id_dok='" & dokument & "' and pozicija='" & pozicija & "'"), 2)
Me.Text4.Text = FormatNumber(Getnumb("select x from trenutna where tip_dok+id_dok='" & dokument & "' and pozicija='" & pozicija & "'"), 2)
Me.Text5.Text = FormatNumber(Getnumb("select y from trenutna where tip_dok+id_dok='" & dokument & "' and pozicija='" & pozicija & "'"), 2)
Me.Text1.Text = FormatNumber(Getnumb("select madanabc from mada where madasifr='" & iden & "'"), 2)
Me.Text2.Text = FormatNumber(Getnumb("select madanabc from mada where madasifr='" & iden & "'") * akoli, 2)
Me.Text6.Text = FormatNumber(Getnumb("select madampcd from mada where madasifr='" & iden & "'"), 2)

Me.Show vbModal
End Sub

Private Sub LaVolpeButton1_Click()
'Dim rsse As New ADODB.Recordset
'rsse.Open "select * from trenutna where tip_dok+id_dok='" & xdxd & "' and pozicija='" & xpxp & "'", myConection, adOpenDynamic, adLockOptimistic
myConection.Execute ("update trenutna set cena=" & stevilka(Text1.Text) & ",znes=" & stevilka(Me.Text2.Text) & ",x=" & stevilka(Text4.Text) & ",pop=" & stevilka(Text3.Text) & ",y=" & stevilka(Me.Text6.Text) & " where  tip_dok+id_dok='" & xdxd & "' and pozicija='" & xpxp & "'")
myConection.Execute ("update mada set madampcd=" & stevilka(Text6.Text) & ",madanabc=" & stevilka(Text1.Text) & " where madasifr='" & iden & "'")
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.LaVolpeButton2.SetFocus
End If
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Text2.Text = FormatNumber(Text1.Text * akoli, 2)
Me.Text3.Text = FormatNumber(0, 2)
End Sub

Private Sub Text1_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Text1.SetFocus
End If
If KeyAscii = 13 Then
Me.LaVolpeButton1.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text1.Text = FormatNumber(Text2.Text / akoli, 2)
Me.Text3.Text = FormatNumber(0, 2)
End Sub

Private Sub Text2_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Text1.SetFocus
End If
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
If Val(Text3.Text) <> 0 Then
Me.Text2.Text = FormatNumber(Text1.Text * akoli * (1 - (Text3.Text / 100)), 2)
End If
End Sub

Private Sub Text3_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Text3.SetFocus
End If
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
If Val(Text4.Text) <> 0 Then
Me.Text5.Text = FormatNumber(Text1.Text * (1 + (Text4.Text / 100)), 2)
Me.Text6.Text = FormatNumber(Text1.Text * (1 + (Text4.Text / 100)) * (1 + (Getnumb("select madapd from mada where madasifr='" & iden & "'") / 100)), 2)
End If
End Sub

Private Sub Text4_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub Text5_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Text4.SetFocus
End If
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text5_LostFocus()
If Val(Text5.Text) <> 0 Then
Me.Text4.Text = FormatNumber(((Text5.Text / Text1.Text) - 1) * 100, 2)
Me.Text6.Text = FormatNumber(Text1.Text * (1 + (Text4.Text / 100)) * (1 + (Getnumb("select madapd from mada where madasifr='" & iden & "'") / 100)), 2)
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Text5.SetFocus
End If
If KeyAscii = 13 Then
Me.LaVolpeButton1.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
If Val(Text6.Text) <> 0 Then
Me.Text5.Text = FormatNumber(Text6.Text / (1 + (Getnumb("select madapd from mada where madasifr='" & iden & "'") / 100)), 2)
Text5_LostFocus
End If
End Sub

Private Sub Text6_GotFocus()
Sendkeys "{END}"
Sendkeys "+{HOME}"
End Sub
