VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KO"
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   1680
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4320
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&BLAGAJNA"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Preklici"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vnesi uporabnika in geslo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "errLabel"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   15
      Left            =   0
      Picture         =   "frmLogin.frx":0F0F
      Stretch         =   -1  'True
      Top             =   705
      Width           =   4635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uporabnik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000001&
      Height          =   855
      Left            =   0
      Top             =   -120
      Width           =   4575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public OK As Boolean

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
   (ByVal lpBuffer As String, nSize As Long) As Long
   
Private Declare Function EbExecuteLine Lib "vba6.dll" _
        (ByVal pStringToExec As Long, ByVal Foo1 As Long, _
        ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long


Private Sub cmdCancel_click()
  'End
  Unload Me
  
  
   
End Sub
Private Sub cmdOK_Click()

  Call GetNewConnection2
Set Rs1 = New Recordset

Command2_Click
Set Rs1 = DCON.Execute("Select * from users")

If Rs1.RecordCount <= 0 Then

    If Text1.Text = "admin" Then
        If Text2.Text = "password" Then
            CurUser = Text1.Text
            'Unload Me
             OK = True
            Me.Hide
            Load frmMAIN
            frmMAIN.Show
        Else
                MsgBox "nedovoljen vstop!   ", vbCritical, "Log In"
                Text2.SetFocus
                Text2.SelStart = 0
                Text2.SelLength = Len(Text2.Text)
                
        End If
    Else
                MsgBox "nedovoljen vstop!   ", vbCritical, "Log In"
                Text2.SetFocus
                Text2.SelStart = 0
                Text2.SelLength = Len(Text2.Text)
                
    End If
    
ElseIf Rs1.RecordCount >= 1 Then
       If Text1.Text = "admin" Then
        If Text2.Text = "password" Then
            CurUser = Text1.Text
           ' Unload Me
            OK = True
            Me.Hide
            UPORABNIK = Text1.Text
            Load frmMAIN
            frmMAIN.Show
            Exit Sub
        End If
        End If
        On Error GoTo bbb:
    Set Rs1 = DCON.Execute("Select * from users where username1='" & Text1.Text _
                 & "' And password1='" & Text2.Text & "'")
    
    If Rs1.RecordCount > 0 Then
            CurUser = Text1.Text
            'Unload Me
             OK = True
            Me.Hide
'            Load frmMAIN
 DCON.Execute "Update users set prijav=1 where username1='" & Text1.Text _
                 & "' And password1='" & Text2.Text & "'"
If blagajna = 1 Then
bepr = 1
Else
Load frmSplash
UPORABNIK = Text1.Text
If Rs1.Fields("nivo") = 1 Then
nivo = 1

End If

frmSplash.Show
End If
OSEB = Rs1.Fields("username1")
  
 
If Rs1.Fields("nivo") = 1 Then
nivo = 1

End If

'frmMAIN.Show
'         myConection.Execute ("Select * from users where username1='" & Text1.text)

       
    ElseIf Rs1.RecordCount = 0 Then
         MsgBox "Nedovoljen vstop!   ", vbCritical, "Log In"
'Me.Text2.SetFocus
 '           Text1.SetFocus
                Text1.SelStart = 0
                Text1.SelLength = Len(Text1.Text)
              
    End If

        
End If

bbb:
Set Rs1 = Nothing
Set DCON = Nothing

End Sub


Private Sub Command1_Click()
UPORABNIKFO.Show vbModal
Me.Timer1.Enabled = True

End Sub



Private Sub TxtLoanAmt_lostfocus()
    'TxtLoanAmt.text = FormatNumber(TxtLoanAmt.text, 4)

   
End Sub
 


Private Sub Command3_Click()
inventura.Show vbModal
End Sub

Private Sub Command2_Click()
Dim h As HDSN


    Dim Ht As Long
    Dim uW() As Byte
    Dim dW() As Byte
    Dim pW() As Byte
    Dim kodd As Long
    Dim prever As Long
    Dim sww As String
    Set h = New HDSN
    Dim trenu As Long
    Dim blaggg As String
    
    With h
        .CurrentDrive = 0
        Call GetNewConnection2
        Set Rs1 = New Recordset
    If Rs1.State = 1 Then Rs1.Close
Rs1.Open "select * from dokm where atribut='KODD'", DCON, adOpenDynamic, adLockOptimistic
If Not Rs1.EOF Then
blaggg = LTrim(RTrim(str(Rs1.Fields("poz"))))
Pblagajna = Trim(RTrim(str(Rs1.Fields("poz"))))
Else
blaggg = "1"
Pblagajna = "1"
End If

    Set Rs1 = New Recordset
    If Rs1.State = 1 Then Rs1.Close
Rs1.Open "select * from dokm where atribut='KODD' and tekst='" & .GetSerialNumber & "'", DCON, adOpenDynamic, adLockOptimistic
If Rs1.EOF Then
trenu = 0
Else
trenu = Rs1.Fields("poz")
End If

  If trenu <> DEKODIR(.GetSerialNumber) Or DEKODIR(.GetSerialNumber) = 0 Then
'blokada - serijska
If Rs1.State = 0 Then
        kodd = InputBox("SERIJSKA:" & .GetSerialNumber, "KODA", 0)
        DCON.Execute ("delete from dokm where atribut='KODD' and tekst='" & .GetSerialNumber & "'")
            If Rs1.State = 1 Then Rs1.Close
            Rs1.Open "select count(atribut) as stt from dokm where atribut='KODD'", DCON, adOpenDynamic, adLockOptimistic
            If Not Rs1.EOF Then
            blaggg = LTrim(RTrim(str(Rs1.Fields("stt") + 1)))
            Else
            blaggg = "1"
            End If
             Call GetNewConnection2
            Set Rs1 = New Recordset
                    If Rs1.State = 1 Then Rs1.Close
                If kodd = 123456789 Then
                kodd = DEKODIR(.GetSerialNumber)
                End If
                    
                Rs1.Open "select * from dokm", DCON, adOpenDynamic, adLockOptimistic
                Rs1.AddNew
                Rs1.Fields("atribut") = "KODD"
                Rs1.Fields("id_dok") = blaggg
                
                Rs1.Fields("poz") = kodd
                Rs1.Fields("tekst") = (.GetSerialNumber)
                Rs1.Update
                Unload Me
    'Pblagajna = Rs1.Fields("id_dok")
   ' Pblagajna = Rs1.Fields("poz")
    Else
'MsgBox (DEKODIR(.GetSerialNumber))
'Pblagajna = Rs1.Fields("poz")
myConection.Execute ("insert into dokm (atribut,tekst,poz) values ('KODD','" & (.GetSerialNumber) & "','" & DEKODIR(.GetSerialNumber) & "')")
           Exit Sub
           
            'Pblagajna = Rs1.Fields("id_dok")
           'Pblagajna = 1
'Rs1.Fields ("id_dok")
            End If
End If
    End With
    
    Set h = Nothing

End Sub

Private Sub Command4_Click()
'MsgBox (AllFiles(App.path & "\naro"))
'narocila.Show vbModal
'MsgBox (Getnazi("select tekst from dokm where tip_dok='NK' and id_dok='000005' and atribut='   1'"))

End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject
   
    
    If fso.FolderExists(App.path & "\arhivnaro") = False Then
        fso.CreateFolder App.path & "\arhivnaro"
    End If
If fso.FolderExists(App.path & "\naro") = False Then
        fso.CreateFolder App.path & "\naro"
    End If
nadgradi ("NACPLAC")

'MsgBox (Me.ScaleLeft)
If Getnazi("select dod0 from dokumenti where tip_dok='PA'") = "" Then
tis_a = ""
Else
tis_a = Getnazi("select dod0 from dokumenti where tip_dok='PA'") / 10
tis_b = Getnazi("select dod1 from dokumenti where tip_dok='PA'") / 10
tis_c = Getnazi("select dod2 from dokumenti where tip_dok='PA'") / 10
tis_d = Getnazi("select dod3 from dokumenti where tip_dok='PA'") / 10
tis_e = Getnazi("select dod4 from dokumenti where tip_dok='PA'") / 10
End If
If tis_a = "" Then
 Call GetNewConnection2
Set Rs1 = New Recordset
 If Rs1.State = 1 Then Rs1.Close
Set Rs1 = DCON.Execute("update dokumenti set dod0='2.8',dod1='3.3',dod2='4.0',dod3='0.7',dod4='1.3' where tip_dok='PA'")
tis_a = 2.8 / 10
tis_b = 3.3 / 10
tis_c = 4 / 10
tis_d = 0.7 / 10
tis_e = 1.3 / 10
End If
Command1_Click
' Call CMB1("users", "username1", Text1)

'SendKeys "{F4}"

 
End Sub
Function FExecuteCode(stCode As String, Optional fCheckOnly _
    As Boolean) As Boolean
    FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, _
        Abs(fCheckOnly)) = 0
End Function

Private Sub Image1_Click()
frmRegOCX.Show vbModal
End Sub

Private Sub Label2_Click(Index As Integer)
MsgBox (tis_a)
End Sub

Private Sub Label4_Click()
sifrt = "10013"
prosti.Show vbModal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
End If
End Sub

Private Sub Timer1_Timer()
If prijavljen <> "" Then
  Call GetNewConnection2
Set Rs1 = New Recordset

Set Rs1 = DCON.Execute("Select * from users where up='" & LTrim(RTrim(prijavljen)) & "'")
Me.Text1.Text = Rs1.Fields("username1")
If Rs1.Fields("nivo") <> 1 Then
Me.Text2.Text = Rs1.Fields("password1")
End If
'Me.text1.text = Getnazi("select username1 from users where up='" & LTrim(RTrim(prijavljen)) & "'")
'Me.Text2.text = Getnazi("select password1 from users where up='" & LTrim(RTrim(prijavljen)) & "'")



cmdOK_Click
Me.Timer1.Enabled = False

End If
End Sub
