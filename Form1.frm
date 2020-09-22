VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Šifrant"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9840
      Top             =   480
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3360
      Left            =   0
      TabIndex        =   13
      Top             =   6240
      Width           =   14040
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   576
      Left            =   6960
      TabIndex        =   11
      Top             =   1320
      Width           =   2820
   End
   Begin VB.TextBox iis1 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   576
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   6180
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   855
      Left            =   10440
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Izberi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   4210752
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":0000
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   855
      Left            =   12120
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Uredi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   4210752
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":001C
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   855
      Left            =   12120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Briši"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   4210752
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":0038
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   855
      Left            =   10440
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Dodaj"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   4210752
      FCOLO           =   255
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":0054
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
   Begin VB.ListBox List87 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3630
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   14040
   End
   Begin VB.TextBox iis 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   576
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7260
   End
   Begin LVbuttons.LaVolpeButton zalog 
      Height          =   495
      Left            =   8040
      TabIndex        =   15
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Na zalogi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16777215
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":0070
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SERIJSKA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SERIJSKA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label ime1 
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NAZIV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NAZIV 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label ime 
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)




Private Sub Form_Activate()
Dim xxxsql As String
If bremepis = 1 Then
zalog_Click
End If

Me.List87.Refresh
If Len(Trim(atab)) > 15 Then
odg = " and "
Else
odg = " where "
End If
Dim af As String
       iis.SetFocus
       iis.Text = ""
       iis.Text = pritisk
       
        ' xxxsql = "sELECT " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & " WHERE " & ss & " Like '" & iis.text & "%' or " & sx & " Like '" & ime1.Caption & "%'"
         
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%'" & nazalogi
        Me.List87.Visible = False
         Filllist List87, xxxsql
         
         Me.List87.Refresh
        Me.List87.Visible = True
         
         iis.SetFocus
        
       
        
        'af = idar
        'SendKeys af
       ' If idar = "" Then
       ' End If
End Sub



Private Sub Form_Load()
'ReSizeForm Me
'     ime.Caption = ""
Me.Top = opp
Me.Left = oppa
End Sub
Private Sub iis_GotFocus()
Sendkeys "{END}"
End Sub
Private Sub iis_change()
Dim xxxsql As String

         'xxxsql = "sELECT " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & " WHERE " & ss & " Like '" & iis.text & "%' or " & sx & " Like '" & ime1.Caption & "%'"
         
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%'" & nazalogi
       Me.List87.Visible = False
        
         Filllist List87, xxxsql
         Me.List87.Refresh
Me.List87.Visible = True
        
End Sub
Private Sub iis_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Dim xxxsql As String

        Case vbKeyEscape
        
        ime.Caption = ""
        Me.iis = ""
        
       ' Me.List87.Refresh
        Case vbKeyDown
        List87.SetFocus
         Sendkeys "{down}"
        
        Case vbKeyBack
        Dim ax As Integer
        If Me.ime.Caption <> "" Then
        ax = Len(Me.ime.Caption) - 1
        'Me.ime.Caption = Left(Me.ime.Caption, ax)
        Me.ime.Caption = Left(Me.ime.Caption, ax)
        End If
           'xxxsql = "sELECT " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & " WHERE " & ss & " Like '" & iis.text & "%' or " & sx & " Like '" & ime1.Caption & "%'"
         
           
           xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' " & nazalogi
          Me.List87.Visible = False
        
         Filllist List87, xxxsql
'       MsgBox (xxxsql & Me.ime.Caption & "*'));")

        Me.List87.Refresh
        Me.List87.Visible = True
        
        Case vbKey0 To vbKey9
          Me.ime.Caption = Trim(Me.iis.Text) & Trim(Chr(KeyCode))
         'xxxsql = "sELECT " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & " WHERE " & ss & " Like '" & iis.text & "%' or " & sx & " Like '" & ime1.Caption & "%'"
         
         
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' " & nazalogi
       Me.List87.Visible = False
        
         Filllist List87, xxxsql
         Me.List87.Refresh
        Me.List87.Visible = True
        
        Case vbKeyA To vbKeyZ
        'Me.ime.Caption = Me.ime.Caption & Chr(KeyCode)
        Sendkeys "{END}"
       
       Me.ime.Caption = Trim(Me.iis.Text)
         'xxxsql = "sELECT " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & " WHERE " & ss & " Like '" & iis.text & "%' or " & sx & " Like '" & ime1.Caption & "%'"
      
         
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' " & nazalogi
      Me.List87.Visible = False
        
         Filllist List87, xxxsql
         Me.List87.Refresh
         Me.List87.Visible = True
        
         Case vbKeyReturn
       
        Sendkeys "{down}"
'        xxre = List87.Value
 '      DoCmd.Close acForm, "nnn"
'        Me.ime.Caption = ""
 '       Me.List87.Requery
         'List87.Selected (0)
    Case Else
      
    'Me.ime.Caption = Me.ime.Caption & Chr(KeyCode)
       
    
    End Select
    'MsgBox (List87.RowSource)
End Sub
Private Sub iis1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Dim xxxsql As String

        Case vbKeyEscape
        
        ime1.Caption = ""
        Me.iis1 = ""
        
       ' Me.List87.Refresh
        Case vbKeyDown
        List87.SetFocus
         Sendkeys "{down}"
        
        Case vbKeyBack
        Dim ax As Integer
        If Me.ime1.Caption <> "" Then
        ax = Len(Me.ime1.Caption) - 1
        'Me.ime.Caption = Left(Me.ime.Caption, ax)
        Me.ime1.Caption = Left(Me.ime1.Caption, ax)
        End If
           xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' or " & sx & " Like '" & ime1.Caption & "%'" & nazalogi
        Me.List87.Visible = False
         
         Filllist List87, xxxsql
'       MsgBox (xxxsql & Me.ime.Caption & "*'));")

        Me.List87.Refresh
        Me.List87.Visible = True
        
        Case vbKey0 To vbKey9
          Me.ime1.Caption = Trim(Me.iis1.Text) & Trim(Chr(KeyCode))
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' or " & sx & " Like '" & ime1.Caption & "%'" & nazalogi
         Me.List87.Visible = False
        
         Filllist List87, xxxsql
         Me.List87.Refresh
        Me.List87.Visible = True
        
        Case vbKeyA To vbKeyZ
        'Me.ime.Caption = Me.ime.Caption & Chr(KeyCode)
       
         Case vbKeyReturn
       
        Sendkeys "{down}"
'        xxre = List87.Value
 '      DoCmd.Close acForm, "nnn"
'        Me.ime.Caption = ""
 '       Me.List87.Requery
         'List87.Selected (0)
    Case Else
      Me.ime1.Caption = Trim(Me.iis1.Text)
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM " & atab & odg & ss & " Like '" & iis.Text & "%' or " & sx & " Like '" & ime1.Caption & "%'" & nazalogi
        Me.List87.Visible = False
         
         Filllist List87, xxxsql
         Me.List87.Refresh
    'Me.ime.Caption = Me.ime.Caption & Chr(KeyCode)
       Me.List87.Visible = True
        
    
    End Select
    'MsgBox (xxxsql)
End Sub

Private Sub LaVolpeButton1_Click()
frmProdEntry.ShowAdd

'frmSupEntry.ShowAdd
    
End Sub

Private Sub LaVolpeButton3_Click()
MODIFYID = Left(Me.List87, 15)
'If frmControlMain.MSHFlexGrid1.TextMatrix(0, frmControlMain.MSHFlexGrid1.Col) = "MADASIFR" Then
Load frmProdEntry
    frmProdEntry.ShowEdit MODIFYID
End Sub

Private Sub LaVolpeButton4_Click()
List87.SetFocus
Sendkeys "{ENTER}"
End Sub
Private Sub Text1_KeyDown(KeyCode2 As Integer, Shift As Integer)
Dim xxxsql As String
Select Case KeyCode2
Case vbKeyReturn
Dim bg As Integer
bg = Getnazi("select sifbl from dspr where stev='" & Me.Text1.Text & "'")
Me.iis.Text = Trim(bg)
 Me.ime.Caption = Trim(bg)
         xxxsql = "sELECT top 100 " & sl & "," & ss & "," & sx & ",madazalo FROM mada WHERE madasifr=" & ime.Caption & nazalogi
         Me.List87.Visible = False
        
         Filllist List87, xxxsql
         Me.List87.Refresh
         Me.List87.Visible = True
        
Case Else
    End Select

End Sub
Private Sub List87_KeyDown(KeyCode2 As Integer, Shift As Integer)
Me.Timer1.Enabled = True
vrjenniz = ""
Select Case KeyCode2

 Case vbKeyEscape
 
 iis.SetFocus
   Me.ime.Caption = ""
   Me.List87.Refresh
 
 
 Case vbKeyReturn
If IsNull(Me.List87) Then
   iis.SetFocus
   Me.ime.Caption = ""
   Me.List87.Refresh
Else
     vrjenniz = Left(Me.List87, 15)
  xxre = (Left(Me.List87, 15))

'blag.MsfBill.Col = 4

End If
'
'DoSQL = xxre
iskalni = ""
Unload Me


'


'blag.msfbill.SetFocus
'blag.msfbill.Row = zaix
'blag.msfbill.Col = 1
'blag.msfbill.TextMatrix(zaix, 1) = xxre
'blag.ArrangeTextbox blag.cmbItmcode
  
'SendKeys "{enter}"

'blag.MsfBill.Col = 4
'blag.ArrangeTextbox blag.cmbItmcode


 Case Else
    End Select

End Sub
Private Sub list87_change()
Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim bsql As String
If Getnazi("select stev from dspr where sifbl<>' ' and sifbl='" & Left(Me.List87, 15) & "'") = "" Then
Me.Height = Me.Label4.Top + Me.Label4.Height

Else
Me.Height = Me.Label4.Top + Me.List1.Height + (Me.Label4.Height * 2)
bsql = "sELECT sifbl,stev,dim,tip,dot,last FROM dspr WHERE sifbl='" & Left(Me.List87, 15) & "'"
        'Me.List87.Visible = False
         
         Filllist List1, bsql
         Me.List1.Refresh
End If

Me.Timer1.Enabled = False
End Sub

Private Sub zalog_Click()
If zalog.BackColor = &HFF& Then
zalog.BackColor = &HFFFFFF
nazalogi = ""
Else

zalog.BackColor = &HFF&
nazalogi = " and madazalo>0"
End If
iis_change
End Sub
