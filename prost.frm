VERSION 5.00
Begin VB.Form prosti 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Izbor proste zaloge BREMEPIS"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10575
   LinkTopic       =   "prosti"
   ScaleHeight     =   5490
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "POTRDI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
      Caption         =   "Vzeta zal"
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cena"
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Zaloga"
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Naziv"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sifra"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Prosta zaloga:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Zeljena kolicina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   840
      Width           =   10215
   End
End
Attribute VB_Name = "prosti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rsta As New ADODB.Recordset
Dim tii, idd As String
Dim ys As String
Dim yc, yp As Double

tii = Left(frmblag.dok.Caption, 2)
idd = Mid(frmblag.dok.Caption, 3)
Dim xpo, lirow As Integer
xpo = 1
If rsta.State = 1 Then rsta.Close
rsta.Open "select * from trenutna where tip_dok='" & tii & "' and id_dok='" & idd & "' and kol=0 order by pozicija desc", myConection, adOpenDynamic, adLockOptimistic
If Not rsta.EOF Then
If Val(rsta.Fields("pozicija")) > 1 Then
xpo = Val(rsta.Fields("pozicija")) - 1
Else
xpo = Val(rsta.Fields("pozicija"))
End If
End If
 For lirow = 0 To List1.ListCount - 1
       
          Me.List1.ListIndex = lirow
    'ys = Left(List1.Column(0, lirow), 9)
    'yc = Mid(List1.Column(0, lirow), 52, 8)
    'yp = Right(RTrim(List1.Column(0, lirow)), 7)
    ys = Left(List1.Text, 9)
    yc = Replace(Mid(List1.Text, 54, 8), "|", "")
    yp = Right(RTrim(List1.Text), 7)
'If rsta.State = 1 Then rsta.Close
If yp > 0 Then
If lirow > 0 Then
rsta.AddNew
End If
rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("sifra") = ys
rsta.Fields("naziv") = Getnazi("select madanazi from mada where madasifr='" & ys & "'")
rsta.Fields("cena") = yc
rsta.Fields("kol") = -yp
rsta.Fields("znes") = -yp * yc
rsta.Fields("datum") = frmblag.DTPicker1.Value
rsta.Fields("pozicija") = levi_pres(LTrim(str(xpo)), 4)
 rsta.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")
   rsta.Fields("skl") = frmblag.sklad.BoundDatax
   rsta.Fields("DOZA") = Getnazi("select madaDOZA from mada where madasifr='" & ys & "'")
   
  
'rsta.Fields("doza") = RSS.Fields("madadoza")
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("ins*ert into trenutna (tip_dok,id_dok,sifra,naziv,kolicina,cena,znes,faktor,doza) values ('" & tii & "','" & idd & "','" & RSS.Fields("madasifr") & "','" & RSS.Fields("madanazi") & "'," & RSS.Fields("madazalo") & "," & RSS.Fields("madanabc") & "," & Round(RSS.Fields("madazalo") * RSS.Fields("madanabc"), 2) & ",1," & RSS.Fields("madadoza"))

xpo = xpo + 1
End If
Next
rsta.AddNew

rsta.Fields("tip_dok") = tii
rsta.Fields("id_dok") = idd
rsta.Fields("datum") = frmblag.DTPicker1.Value
rsta.Fields("pozicija") = levi_pres(LTrim(str(xpo + 1)), 4)
 rsta.Fields("uporabnik") = Getnazi("select up from users where username1='" & UPORABNIK & "'")
   rsta.Fields("skl") = frmblag.sklad.BoundDatax
   
rsta.Fields("faktor") = 1
rsta.Update
'myConection.Execute ("delete from trenutna where kol=0")
'Call frmblag.cmdAdd_Click
Unload Me
'Call frmblag.doddaa
frmblag.fgtrial.Redraw = True

frmblag.refre
dell = 1
   ' fgtrial.DataSource = RS
'    MsgBox ""
    lstRow = frmblag.fgtrial.Rows - 1
    For lirow = 1 To lstRow
    frmblag.fgtrial.TextMatrix(lirow, 0) = lirow
    Next
    frmblag.fgtrial.Row = lstRow
    frmblag.fgtrial.Col = coollsi
     frmblag.fgtrial.TextMatrix(lstRow, 0) = lstRow
     frmblag.fgtrial.TextMatrix(lstRow, 1) = ""
     frmblag.fgtrial.SetFocus
frmblag.refre

'     frmblag.doddaa
End Sub

Private Sub Form_Load()
'myConection.Execute ("delete from trenutna where sifra=''")

Filpro List1, "select min(sifra) as sif,min(naziv) as naz,sum(prosta) as pros,cena from zaloga where prosta>0 and sifra='" & LTrim(RTrim(sifrt)) & "' group by cena"
Label2.Caption = Getnazi("select sum(prosta) as x from zaloga where prosta>0 and sifra='" & LTrim(RTrim(sifrt)) & "' group by sifra")
End Sub
Private Function Filpro(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long
Dim pobr, pobr1 As Double
If Not rs.EOF Then

    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    If Me.Text1.Text = "" Then
    Me.Text1.Text = 0
    
    End If
    pobr = Round((Me.Text1.Text), 3)
    '
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    If pobr > 0 Then
    If rs.Fields("pros") > pobr Then
    pobr1 = pobr
    Else
    pobr1 = rs.Fields("pros")
    End If
    pobr = pobr - rs.Fields("pros")
    Else
    pobr1 = 0
    End If
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
        dolg = presled(Trim((rs.Fields("sif"))), 10)
       ' zalo = Str(RS.Fields(3))
        With rs
        
            Cmbl.AddItem dolg & " | " & presled(Trim(.Fields("naz")), 25) & " | " & levi_pres(Trim(.Fields(2)), 8) & " | " & levi_pres(Trim(.Fields(3)), 8) & " | " & levi_pres(str(Round(pobr1, 3)), 8)
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function

Private Sub List1_GotFocus()
Me.Text1.SelStart = 0
Me.Text1.SelLength = 5
Me.Text1.SetFocus
End Sub

Private Sub Text1_Change()
Text1.Text = Trim(str(Val(Text1.Text)))
If Val(Text1.Text) > Val(Label2.Caption) Then
 Text1.Text = Label2.Caption
 End If
Filpro List1, "select min(sifra) as sif,min(naziv) as naz,sum(prosta) as pros,cena from zaloga where prosta>0 and sifra='" & LTrim(RTrim(sifrt)) & "' group by cena"
End Sub
