VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Begin VB.Form C_frmCustomer 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Podatki o partnerju"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9030
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
   ScaleHeight     =   6825
   ScaleMode       =   0  'User
   ScaleWidth      =   5283.449
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3760
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4640
      Width           =   4964
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   2400
      TabIndex        =   28
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Preveri v spletu"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomers.frx":0000
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
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Sinhroniziraj"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   16315377
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomers.frx":001C
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
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   6200
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   222
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2600
      Width           =   1033
   End
   Begin VB.TextBox txtcontact 
      Height          =   375
      Left            =   222
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3200
      Width           =   3255
   End
   Begin VB.TextBox txttitle 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1368
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2600
      Width           =   3255
   End
   Begin VB.TextBox txtdisc 
      Height          =   375
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   9
      Top             =   5360
      Width           =   2565
   End
   Begin VB.TextBox txtdays 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   5360
      Width           =   735
   End
   Begin VB.TextBox txtlimit 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtremarks 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3215
      Left            =   5429
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtphone 
      Height          =   375
      Left            =   240
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4640
      Width           =   3255
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaxLength       =   200
      TabIndex        =   1
      Top             =   1920
      Width           =   4964
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaxLength       =   230
      TabIndex        =   0
      Top             =   1200
      Width           =   4964
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Preklici"
      Height          =   375
      Left            =   7994
      TabIndex        =   13
      Top             =   6200
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   7034
      TabIndex        =   12
      Top             =   6200
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TRR"
      Height          =   175
      Left            =   3840
      TabIndex        =   29
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Dob/Kup"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5960
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pošta"
      Height          =   255
      Left            =   222
      TabIndex        =   25
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   24
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Opomba"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   5429
      TabIndex        =   23
      Top             =   1000
      Width           =   600
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Država:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   22
      Top             =   5120
      Width           =   570
   End
   Begin VB.Label lbl_crlimit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Max kredo:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   5120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Placilo dni:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   1440
      TabIndex        =   20
      Top             =   5120
      Width           =   750
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   4400
      Width           =   270
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Davcna"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   222
      TabIndex        =   18
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Naslov:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label lbl_cust 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mesto"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   1367
      TabIndex        =   16
      Top             =   2340
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Naziv"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partnerji"
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
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmCustomers.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9119
   End
End
Attribute VB_Name = "C_frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sSQL As String
Dim NOVA As Integer
Dim vrst As String
If Me.txtdays.Text = "" Then
Me.txtdays.Text = "0"
End If
If Me.Text1.Text = "" Then
Me.Text1.Text = "0"
End If

vrst = Left(CatalogueName, 1)
 sSQL = "SELECT * " & _
            " From partner"
            If rs.State = 1 Then rs.Close
         rs.Open sSQL, myConection, adOpenStatic, adLockOptimistic
            NOVA = LTrim(Me.Label1.Caption)
            myConection.Execute ("delete from partner where sifra=" & NOVA)
       rs.AddNew
       rs.Fields("sifra") = NOVA
       rs.Fields("naziv") = Me.txtname.Text
       rs.Fields("ulica") = Me.txtaddress.Text
       rs.Fields("mesto") = Me.txttitle.Text
       rs.Fields("davcna") = Trim(Me.txtcontact.Text)
       rs.Fields("telefon") = Me.txtphone.Text
       rs.Fields("vrsta") = Text2.Text
       rs.Fields("maxlimit") = Val(Me.txtdays.Text)
       rs.Fields("posta") = Val(Me.Text1.Text)
        rs.Fields("oseba") = Trim(Me.txtdisc.Text)
        rs.Fields("ziro") = Trim(Me.Text3.Text)
       rs.Update
           '  SQL = "Insert into partner (sifra,naziv,ulica,mesto,davcna,telefon,vrsta,maxlimit,posta) values (" & NOVA & ",'" &
           'Me.txtname.text & "','" & Me.txtaddress.text & "','" & Me.txttitle.text & "','" & Me.txtcontact.text
           '& "','" & Me.txtphone.text & "','" & vrst & "'," & Val(Me.txtdays.text) & ",'" & Me.Text1.text & "')"
      'myConection.Execute SQL
frmMAIN.beno_os (28)
Unload Me
 frmMAIN.Label6_Click


End Sub

Private Sub CustSave()
Dim sSQL As String
Dim NOVA As Integer
Dim vrst As String
vrst = Left(CatalogueName, 1)
 sSQL = "SELECT Max(partner.sifra)+1 AS novast" & _
            " From partner"
            If rs.State = 1 Then rs.Close
         rs.Open sSQL, myConection, adOpenStatic, adLockOptimistic
            NOVA = rs.Fields("novast")
             SQL = "Insert into partner (sifra,naziv,ulica,mesto,davcna,telefon,vrsta) values (" & NOVA & ",'" & Me.txtname.Text & "','" & Me.txtaddress.Text & "','" & Me.txttitle.Text & "','" & Me.txtcontact.Text & "','" & Me.txtphone.Text & "','" & vrst & "')"
      myConection.Execute SQL
Unload Me
 frmMAIN.Label6_Click


End Sub
Private Sub CustUpdate()
Dim CDb As CDbase
Dim CUpd As New CUpdate
Dim CustID As String


Call GetNewConnection2

Set Rs1 = New Recordset
Set Rs1 = DCON.Execute("Select * from partner where naziv='" & txtname.Text & "'")

If Rs1.RecordCount = 0 Then

Call GetNewConnection(CUpd)
Set CDb = CUpd

CDb.TableName = "partner"
CDb.ClauseStatement = "Where sifra=" & Val(MODIFYID)

Call CUpd.FieldVal(Trim(txtname), Trim(txtaddress), Trim(txtphone), Trim(txtcontact), Trim(txttitle), Trim(txtremarks))

Call CUpd.Update("naziv", "ulica", "telefon", "davcna", "mesto", "ziro")

MsgBox "Zapis shranjen!", vbInformation

Unload Me

Set CUpd = Nothing


    Else
          
            If Rs1!sifra <> Val(MODIFYID) Then
            
                MsgBox "The Category was already exist", vbInformation
            Else
            
            Call GetNewConnection(CUpd)
            Set CDb = CUpd

            CDb.TableName = "partner"
                CDb.ClauseStatement = "Where sifra=" & Val(MODIFYID)

           Call CUpd.FieldVal(Trim(txtname), Trim(txtaddress), Trim(txtphone), Trim(txtcontact), Trim(txttitle), Trim(txtremarks))

             Call CUpd.Update("naziv", "ulica", "telefon", "davcna", "mesto", "ziro")

            MsgBox "Zapis shranjen!", vbInformation

                Unload Me

            Set CUpd = Nothing

          
            End If
End If

Set Rs1 = Nothing
Set DCON = Nothing


End Sub
Private Sub Form_Activate()
On Error Resume Next
If Getnazi("select max(sifra) as xx from partner") = "" Then
Me.Label1.Caption = "1"
Else
Me.Label1.Caption = LTrim(str(Val(Getnazi("select max(sifra) as xx from partner") + 1)))
End If
Me.Text2.Text = Left(CatalogueName, 1)
If ADDING = False Then
    Call GetNewConnection2
    Set Rs1 = New Recordset
   ' MsgBox (MODIFYID)
    Set Rs1 = DCON.Execute("Select * from partner where sifra=" & Val(MODIFYID))

    If Rs1.RecordCount <> 0 Then
        txtname.Text = Rs1!naziv & " "
        txtaddress.Text = Rs1!ulica & " "
        txtphone.Text = Rs1!telefon & " "
        txtcontact.Text = Rs1!davcna & " "
        txttitle.Text = Rs1!mesto & " "
        txtremarks.Text = Rs1!ziro & " "
        Me.Label1.Caption = Rs1!sifra & " "
        Me.txtdays.Text = Rs1!maxlimit & " "
        Me.Text1.Text = Rs1!posta & " "
         Me.txtdisc.Text = Rs1!oseba & " "
         Me.Text2.Text = Rs1!vrsta & ""
          Me.Text3.Text = Rs1!ziro & ""
    End If
    Set Rs1 = Nothing
    Set DCON = Nothing
    
End If
End Sub

Private Sub Form_Load()
'    lblError.BackColor = vbWhite
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
KEYSCII = Asc(UCase(Chr(KeyAscii)))


End Sub


Private Sub txtphone_Validate(KEEPFOCUS As Boolean)

''str=Replace(txtphone, "-", "", Len(txtphone), 1, , vbTextCompare))
''If Len(txtphone.text) <> 0 And IsNumeric(= False Then
'    KeepFocus = True
'    txtphone.BackColor = vbRed
'    lblError = "** Invalid Entry "
'Else
'    txtphone.BackColor = vbWhite
'    lblError = ""
'End If
End Sub

