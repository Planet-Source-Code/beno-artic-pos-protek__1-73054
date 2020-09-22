VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Postava 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VISINA,DOLZINA"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Postava.frx":0000
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
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "DOLŽINA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "VIŠINA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Postava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_click()
'MsgBox zap
visina = Val(Text1.text)
dolzina = Val(Text2.text)
Unload Me
frmblag.fgtrial.SetFocus
frmblag.fgtrial.TextMatrix(zai, coollx) = Str(visina)
frmblag.fgtrial.TextMatrix(zai, coolly) = Str(dolzina)
'blag.MsfBill.TextMatrix(zap, 1) = sifrt
frmblag.fgtrial.TextMatrix(zai, coollko) = ""


frmblag.fgtrial.Row = zai
frmblag.fgtrial.Col = 3
'blag.ArrangeTextbox blag.txtEnter(blag.MsfBill.Col)
If MsgBox("Spremenim naziv v naziv + X * Y", vbInformation + vbYesNo) = vbYes Then
frmblag.fgtrial.TextMatrix(zai, 2) = Trim(frmblag.fgtrial.TextMatrix(zai, 2)) & " " & Trim(Str(visina)) & " X " & Trim(Str(dolzina))
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")
Select Case KeyCode

 Case vbKeyReturn
Text2.SetFocus
Case Else
    End Select
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox ("5")
Select Case KeyCode

 Case vbKeyReturn
 Me.LaVolpeButton1.SetFocus
Case Else
    End Select
End Sub


