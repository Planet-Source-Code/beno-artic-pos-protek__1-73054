VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Xpl 
   Caption         =   "Xpl"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   LinkTopic       =   "Form7"
   ScaleHeight     =   10890
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton Prekini 
      Height          =   975
      Left            =   13320
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "PREKINI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
      MICON           =   "Xpl.frx":0000
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
   Begin LVbuttons.LaVolpeButton Zapis 
      Height          =   975
      Left            =   11640
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "ZAPIŠI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
      MICON           =   "Xpl.frx":001C
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5760
      Top             =   1440
   End
   Begin ProsVent.UserControl1 sifr 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Locked          =   0   'False
      polje           =   "tekst"
      ssql            =   "select poz,tekst from dokm where atribut='PLLE' order by poz"
      TextLocked      =   0   'False
   End
   Begin VB.TextBox txtNewData 
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
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgtrial 
      DragIcon        =   "Xpl.frx":0038
      Height          =   7320
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   12912
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      ForeColorSel    =   16777215
      BackColorUnpopulated=   16777152
      GridColor       =   12632256
      GridColorFixed  =   16777215
      GridColorUnpopulated=   14737632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLines       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin ProsVent.UserControl1 Zapo 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   5175
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "sifra"
      ssql            =   "select * from zaposleni"
      TextLocked      =   0   'False
   End
   Begin MSForms.CommandButton COO 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   480
      Width           =   615
      Caption         =   "==>"
      Size            =   "1085;661"
      FontHeight      =   165
      FontCharSet     =   238
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Caption         =   "Zaposlen:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Xpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub COO_Click()
Dim yrst As New ADODB.Recordset
Dim iiy As Long
iiy = 1
If yrst.State = 1 Then xrst.Close

yrst.Open "select * from delo where sifd='" & Zapo.BoundDatax & "'", myConection, adOpenDynamic, adLockOptimistic
If Not yrst.EOF Then
For i = 1 To yrst.Fields.Count - 1

If Not IsNull(yrst.Fields(i)) Then
If Not Trim(yrst.Fields(i)) = "" Then
If IsNumber(yrst.Fields(i)) Then
If Not Val(yrst.Fields(i)) = 0 Then
fgtrial.Rows = iiy + 1
'MsgBox yrst.Fields(i).Name
fgtrial.TextMatrix(iiy, 0) = iiy
fgtrial.TextMatrix(iiy, 1) = Getnazi("select poz from dokm where atribut='PLLE' and id_dok='" & yrst.Fields(i).Name & "'")
fgtrial.TextMatrix(iiy, 2) = Getnazi("select tekst from dokm where atribut='PLLE' and id_dok='" & yrst.Fields(i).Name & "'")
fgtrial.TextMatrix(iiy, 3) = yrst.Fields(i)
iiy = iiy + 1
End If
End If
End If
End If
Next
End If

End Sub

Private Sub fgTrial_KeyPress(KeyAscii As Integer)

If fgtrial.Col = 1 Then
      sifr.Visible = True
      sifr.Move fgtrial.CellLeft + fgtrial.Left - 10, fgtrial.CellTop + _
                        fgtrial.Top - 10, fgtrial.CellWidth + 10, fgtrial.CellHeight

        sifr.SetFocus
Else
        txtNewData.Move fgtrial.CellLeft + fgtrial.Left - 10, fgtrial.CellTop + _
                        fgtrial.Top - 10, fgtrial.CellWidth + 10, fgtrial.CellHeight - 40
        txtNewData.Visible = True
        txtNewData.SetFocus
 End If
End Sub


Private Sub Form_Load()
Msr
xizb = ""
fgtrial.TextMatrix(1, 0) = "1"
fgtrial.Col = 1
fgtrial.Row = 1


End Sub

Private Sub Msr()
Dim sngVertFactor As Single
    sngVertFactor = getFactor(True)
With fgtrial
      .Cols = 4
      .Rows = 2
      .FormatString = "^No | SIFRA | NAZIV    | Kolicina | Vrednost  "
       gSlno = 0
       gItemCode = 1
       gItemname = 2
       gQty = 3
       gRate = 4
       gTotal = 5
       gpop = 6
       .Row = 0
       kjje = 1
       For Inti = 0 To .Cols - 1
          .Col = Inti
          .CellFontBold = True
       Next
       .ColWidth(0) = 6 * 100 * sngVertFactor
       .ColWidth(1) = 20 * 100 * sngVertFactor
       .ColWidth(2) = 60 * 100 * sngVertFactor
       .ColWidth(3) = 20 * 100 * sngVertFactor
       .ColWidth(4) = 20 * 100 * sngVertFactor
       .RowHeight(0) = 350 * sngVertFactor
       .RowHeightMin = 350 * sngVertFactor
End With
End Sub

Private Sub Prekini_Click()
Unload Me
End Sub

Private Sub sifr_GotFocus()
kjje = fgtrial.Row
SendKeys "{HOME} +{END}"
End Sub

Private Sub Timer1_Timer()

If xizb <> "" Then

fgtrial.TextMatrix(kjje, 2) = xizb
fgtrial.TextMatrix(kjje, 1) = Getnazi("select poz from dokm where atribut='PLLE' and tekst='" & xizb & "'")
xizb = ""
sifr.Visible = False
fgtrial.Row = kjje
fgtrial.Col = 3
'Me.txtNewData.Visible = True
'Me.txtNewData.SetFocus
fgtrial.SetFocus
SendKeys "{ENTER}"

End If
End Sub

Private Sub txtNewData_Change()
fgtrial.text = txtNewData.text
End Sub


Private Sub txtNewData_gotfocus()
If Me.txtNewData.text <> "" Then

SendKeys "{HOME} +{END}"
End If
End Sub

Private Sub txtNewData_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
 Case vbKey0 To vbKey9
 If kolik = 1 Then
SendKeys "{BS}"
SendKeys Chr(KeyCode)
kolik = 0
End If
 Case vbKeyA To vbKeyZ
 SendKeys "{BACKSPACE}"

Case Else
    End Select
End Sub
Private Sub txtNewData_KeyPress(KeyAscii As Integer)
kjje = fgtrial.Row
If KeyAscii >= 48 And KeyAscii <= 57 Then

End If
If KeyAscii = 13 Then
  If fgtrial.Col = 1 Then
     fgtrial.Col = 2
     Arran txtNewData
  ElseIf fgtrial.Col = 2 Then
     fgtrial.Col = 3
     Arran txtNewData
  ElseIf fgtrial.Col = 3 Then
     
  If fgtrial.TextMatrix(kjje, 3) = "" Then Exit Sub
  
  
           fgtrial.Rows = fgtrial.Rows + 1
           kjje = fgtrial.Rows - 1
           fgtrial.Col = 1
           fgtrial.Row = fgtrial.Rows - 1
           fgtrial.TextMatrix(kjje, 0) = kjje
           txtNewData.Visible = False
           Arran sifr
 
End If
End If
End Sub
Public Sub Arran(ctrl As Control)
  ctrl.Left = fgtrial.Left + fgtrial.CellLeft
  ctrl.Top = fgtrial.Top + fgtrial.CellTop
  ctrl.BoundDatax = fgtrial.text
  ctrl.Width = fgtrial.ColWidth(fgtrial.Col) - 10
  If TypeOf ctrl Is TextBox Then
  ctrl.Height = fgtrial.RowHeight(fgtrial.Row) - 10
  End If
  ctrl.Visible = True
  ctrl.BoundDatax = ""
  ctrl.SetFocus
  kjje = fgtrial.Row
'  ctrl.SelStart = 0
'  ctrl.SelLength = Len(ctrl.BoundDatax)
End Sub

Private Sub Zapis_Click()
Dim xrst As New ADODB.Recordset
If xrst.State = 1 Then xrst.Close
If Getnazi("select sifd from delo where sifd='" & Zapo.BoundDatax & "'") = "" Then
myConection.Execute "insert into delo (sifd) values ('" & Zapo.BoundDatax & "')"
End If
xrst.Open "select * from delo where sifd='" & Zapo.BoundDatax & "'", myConection, adOpenDynamic, adLockOptimistic
Dim kirfi As String
kirfi = ""
If Not xrst.EOF Then
For i = 1 To fgtrial.Rows - 1
kirfi = Trim(Getnazi("select id_dok from dokm where atribut='PLLE' and poz=" & fgtrial.TextMatrix(i, 1)))
If kirfi <> "" Then
xrst.Fields(kirfi) = fgtrial.TextMatrix(i, 3)
xrst.Update
End If
Next
End If

End Sub

Private Sub Zapo_Validate(Cancel As Boolean)
Dim yrst As New ADODB.Recordset

If yrst.State = 1 Then xrst.Close

yrst.Open "select * from delo where sifd='" & Zapo.BoundDatax & "'", myConection, adOpenDynamic, adLockOptimistic
If Not yrst.EOF Then
For i = 1 To yrst.Fields.Count - 1
fgtrial.Rows = i + 1
fgtrial.TextMatrix(i, 1) = Getnazi("select poz from dokm where atribut='PLLE' and id_dok='" & yrst.Fields(i).Name & "'")
Next
End If

End Sub
