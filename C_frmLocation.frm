VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form C_frmLocation 
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kategorije"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3900
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   100
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   3300
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   2000
      TabIndex        =   15
      Top             =   2700
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   2000
      TabIndex        =   13
      Top             =   2330
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   8
      Top             =   2650
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   2340
      Width           =   255
   End
   Begin VB.TextBox txtremarks 
      Height          =   400
      Left            =   120
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   3300
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   120
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Preklici"
      Height          =   375
      Left            =   2565
      TabIndex        =   4
      Top             =   3500
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1650
      TabIndex        =   3
      Top             =   3500
      Width           =   855
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   375
      Left            =   100
      TabIndex        =   18
      Top             =   3500
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "B"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "C_frmLocation.frx":0000
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
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VSTOPNICE"
      Height          =   255
      Left            =   2300
      TabIndex        =   16
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STORITVE"
      Height          =   255
      Left            =   2300
      TabIndex        =   14
      Top             =   2330
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CIGARETI"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PIJACA"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   2390
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HRANA"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2700
      Width           =   735
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Naziv"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   390
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sifra grupe"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Podatki o grupah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "C_frmLocation.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "C_frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Check1_Click(Index As Integer)
aavr = Index
Dim i, x
For i = 0 To 3
If i = aavr Then
'Me.Check1(i).Value = True
Else
Me.Check1(i).Value = False
End If
Next
End Sub

Private Sub cmdCancel_click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim aa As Integer
        
         aa = aavr
myConection.Execute ("delete from  dokm where atribut='BARV' and id_dok='" & Trim(Me.txtremarks.Text) & "'")
myConection.Execute ("insert into dokm (atribut,id_dok,poz) values ('BARV','" & Trim(txtremarks.Text) & "'," & Me.Picture1.BackColor & ")")
myConection.Execute ("delete from grupa where sifra=" & Trim(txtname.Text))
myConection.Execute ("insert into grupa (sifra,grupa,vr) values(" & (Trim(txtname.Text)) & ",'" & txtremarks & "'," & aa & ")")

frmMAIN.beno_os (16)
Unload Me

End Sub
Private Sub LocUpdate()
 Call GetNewConnection2
    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from grupa where grupa='" & txtname.Text & "'")
        Dim aa As Integer
        
        If Rs1.RecordCount = 0 Then
       
         aa = aavr
            DCON.Execute "Update grupa set sifra=" & Trim(txtname.Text) & ",grupa='" & txtremarks & "',vr=" & aa & " where sifra=" & Val(MODIFYID)
            MsgBox "Zapisano", vbInformation
            Unload Me
          Else
          
            If txtname.Text <> MODIFYID Then
            
               MsgBox "Zapis že obstaja", vbInformation
             
             Else
            MsgBox "Zapis je bil nadgrajen", vbInformation
             Unload Me
            End If
        
        End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
        
End Sub
Private Sub LocSave()
Dim CDb As CDbase
Dim CIns As New CInsert
Dim CustID As String


Call GetNewConnection(CIns)
Set CDb = CIns



CDb.TableName = "grupa"


CIns.FieldVal txtname, CText
CIns.FieldVal txtremarks, CText
'CIns.FieldVal Check1, CText
CIns.Insert
    MsgBox "Shranjeno", vbInformation
    txtname.Text = ""
    txtremarks.Text = ""
'For Each Control In C_frmCustomer
'    If TypeOf Control Is TextBox Then
'        Control.text = ""
'    End If
'Next

Set CIns = Nothing
End Sub
Private Sub Form_Load()
 Call GetNewConnection2

    Set Rs1 = New Recordset
    Set Rs1 = DCON.Execute("Select * from grupa where sifra=" & Val(MODIFYID))

    If Rs1.RecordCount <> 0 Then
       a = Rs1!vr
        txtname.Text = Rs1!sifra & " "
        txtremarks.Text = Rs1!grupa & " "
        Me.Check1(a).Value = 1
  
    End If
     Set Rs1 = Nothing
    Set DCON = Nothing
    If Val(Getnazi("select poz from dokm where atribut='BARV' and id_dok='" & Trim(txtremarks.Text) & "'")) <> 0 Then
        Me.Picture1.BackColor = Getnazi("select poz from dokm where atribut='BARV' and id_dok='" & Trim(txtremarks.Text) & "'")
    End If
'Else
If Me.txtname.Text = "" Then
'MsgBox Getnazi("SELECT Max(Val([SIFRA])) AS Izr1 from grupa")
Me.txtname.Text = Val(Getnazi("SELECT Max(Val([SIFRA])) AS Izr1 from grupa") + 1)
End If
'Me.SetFocus
End Sub
Private Sub Form_Activate()
On Error Resume Next
Dim a As Integer
If ADDING = False Then
   
   
End If
 Me.Picture1.BackColor = Getnazi("select poz from dokm where atribut='BARV' and id_dok'" & Trim(txtremarks.Text) & "'")
 
 
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub LaVolpeButton3_Click()
 On Error GoTo NoColorChosen
   With CommonDialog1
      .CancelError = True
      ' Entire dialog box is displayed, including the Define Custom Colors section
      .flags = cdlCCFullOpen
      .ShowColor  ' Launch the Color Dialog
      Me.Picture1.BackColor = .Color  ' Assign selected color to background of Picture1
      Exit Sub
   End With
NoColorChosen:
   ' Get here if user clicks the Cancel button
   MsgBox "NISI SI IZBRAL BARVE!", vbInformation, "Cancelled"
   Exit Sub
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
