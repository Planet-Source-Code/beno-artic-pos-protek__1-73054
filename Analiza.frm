VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Analiza 
   Caption         =   "Analiza"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Analiza.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   8325
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton Pregle 
      Height          =   615
      Left            =   10680
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   4
      TX              =   "Pregled"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Analiza.frx":0CCE
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
   Begin LVbuttons.LaVolpeButton Exce 
      Height          =   615
      Left            =   12000
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Excel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Analiza.frx":0CEA
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
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   300
      Left            =   11760
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin LVbuttons.LaVolpeButton izv 
      Height          =   735
      Left            =   11880
      TabIndex        =   12
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Analiza prodaje"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Analiza.frx":0D06
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   600
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DATOD 
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51249153
      CurrentDate     =   39507
   End
   Begin MSComCtl2.DTPicker DATDO 
      Height          =   375
      Left            =   12360
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51249153
      CurrentDate     =   39507
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   4815
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin ProsVent.UserControl1 UserControl11 
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   6800
      _ExtentY        =   661
      Locked          =   0   'False
      polje           =   "madasifr"
      ssql            =   "select madasifr,madanazi from mada order by madanazi"
      TextLocked      =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "Analiza.frx":0D22
      Height          =   3825
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   6747
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
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   735
      Left            =   9840
      TabIndex        =   17
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Analiza Zaloge"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15790320
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Analiza.frx":102C
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
End
Attribute VB_Name = "Analiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public skly As Integer
Public naslo As String

Private Sub Check1_Click()
If Me.Check1 Then
Me.DATDO.Enabled = True
Me.DATOD.Enabled = True

Else
Me.DATDO.Enabled = False
Me.DATOD.Enabled = False


End If
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
If UCase(Trim(Me.Combo1(Index).Text)) = "PARTNER" Then

Me.UserControl11(Index).sSQL = "select sifra,naziv from partner order by naziv"
Me.UserControl11(Index).polje = "naziv"
Me.UserControl11(Index).opentime
'Me.UserControl11(Index).UserControl_InitProperties
'MsgBox UserControl11(Index).sSQL
End If
If UCase(Trim(Me.Combo1(Index).Text)) = "SIFRA" Then

Me.UserControl11(Index).sSQL = "select madasifr,madanazi from mada order by madanazi"
Me.UserControl11(Index).polje = "madasifr"
Me.UserControl11(Index).opentime
'Me.UserControl11(Index).UserControl_InitProperties
'MsgBox UserControl11(Index).sSQL
End If
skly = Index
For xxx = Index To 3
Me.UserControl11(xxx + 1).Visible = False
Me.Combo1(xxx + 1).Text = ""
Me.Combo1(xxx + 1).Visible = False
Next xxx

If Index < 4 Then
Me.UserControl11(Index + 1).Visible = True
Me.Combo1(Index + 1).Visible = True
FillCom_fields Combo1(Index + 1), "nabasif", Trim(Combo1(0).Text), Trim(Combo1(1).Text), Trim(Combo1(2).Text), Trim(Combo1(3).Text), Trim(Combo1(4).Text)
End If

End Sub

Private Sub DTPicker1_CallbackKeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Exce_Click()
'MSHFlexGrid1.DataSource
If Me.MSHFlexGrid1.Visible = True Then
Dim naslv  As String
naslv = Getnazi("select glava1 from oblikar") & vbCrLf & naslo
Dim slik As String
slik = App.path & "\gaber.jpg"

FlexGrd_SaveToExcel MSHFlexGrid1, naslv, Date, vbWhite, 255, "", , , , , True, True
End If
End Sub

Private Sub Form_Load()
FillCom_fields Combo1(0), "nabasif", "", "", "", "", ""
Me.DATDO.Value = Date
Me.DATOD.Value = Date

End Sub

Private Sub izv_Click()
Dim dodno, dodn, sqwh As String
Dim commd, ccc As String
naslo = " Analiza po :"
commd = ""
ccc = ""
dodn = ""

dodno = ""
sqwh = ""
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
For bb = 0 To skly
commd = "nabasif." & Trim(Me.Combo1(bb).Text)
ccc = Trim(Me.Combo1(bb).Text)
naslo = naslo & ccc
If Me.UserControl11(bb).BoundDatax <> "" Then
naslo = naslo & "(vsebuje " & Me.UserControl11(bb).BoundDatax & "),"
Else
naslo = naslo & ","
End If
If ccc <> "" Then

If UCase(ccc) = "ID_DOK" Then
Dim part As String

commd = "nabasif.id_dok, min(glavna.dod0) as partner"

End If
If UCase(ccc) = "SIFRA" Then
commd = "nabasif.sifra,min(nabasif.naziv) as naziv"
End If


If dodno = "" Then
dodno = commd & ","
Else
dodno = dodno & commd & ","
End If
If dodn = "" Then
dodn = "nabasif." & Trim(Me.Combo1(bb).Text) & ","
Else
dodn = dodn & "nabasif." & Trim(Me.Combo1(bb).Text) & ","
End If

If Me.UserControl11(bb).BoundDatax <> "" Then
If sqwh = "" Then
sqwh = " where " & "nabasif." & Trim(Me.Combo1(bb).Text) & " in ('" & Replace(Trim(Me.UserControl11(bb).BoundDatax), ",", "','") & "')"
Else
sqwh = sqwh & " and " & "nabasif." & Trim(Me.Combo1(bb).Text) & " in ('" & Replace(Trim(Me.UserControl11(bb).BoundDatax), ",", "','") & "')"

End If
End If
End If
'MsgBox sqwh
Next bb
If Me.Check1 Then
naslo = naslo & vbCrLf & "Za obdobje od: " & Format(Me.DATOD.Value, "dd.mm.yyyy") & " do: " & Format(Me.DATDO.Value, "dd.mm.yyyy")
End If
If rs.State = 1 Then rs.Close

sxqll = "select space(1) as q," & dodno & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl,sum(format(nabasif.kol,'#,##0.00')) as kol,sum(format((nabasif.cena-(nabasif.cena*(1-(iif(nabasif.pop=0,100,nabasif.pop)/100)))),'#,##0.00')) as cena,sum(format((nabasif.cena-(nabasif.cena*(1-(iif(nabasif.pop=0,100,nabasif.pop)/100))))*nabasif.kol,'#,##0.00')) as znesek from nabasif LEFT JOIN glavna ON (nabasif.id_dok = glavna.id_dok) AND (nabasif.tip_dok = glavna.tip_dok) " & sqwh & " group by " & dodn & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl order by  " & dodn & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl"
If Me.DATDO.Enabled = True Then
If sqwh = "" Then
sxqll = Replace(sxqll, "group", "where nabasif.DATUM between #" & dod & "# And  #" & ddo & "# group")
Else
sxqll = Replace(sxqll, "where", "where nabasif.DATUM between #" & dod & "# And  #" & ddo & "# and")
End If

End If
rs.Open sxqll, myConection, adOpenDynamic, adLockOptimistic


'koli = koli + RS.Fields("kol")
'popu = popu + ((RS.Fields("cena") - (RS.Fields("cena") * (1 - (RS.Fields("pop") / 100)))) * RS.Fields("kol"))

If rs.EOF Then
   Me.MSHFlexGrid1.Visible = False
    
   
Else
   Me.MSHFlexGrid1.Visible = True
    
    Set Me.MSHFlexGrid1.DataSource = rs
     
End If

osve
   

End Sub


Private Sub osve()
Dim skupko, skupce, skupzn, zall, szza As Double
Dim cooszza As Integer
cooszza = 0
skupko = 0
skupce = 0
skupzn = 0
zall = 0
szza = 0
If frmControlMain.MSHFlexGrid1.Visible = True Then
For i = MSHFlexGrid1.Col To MSHFlexGrid1.Cols - 1
        Dim asx As String
        
        asx = Trim(MSHFlexGrid1.TextMatrix(0, i))
        If UCase(asx) = "CENA" Then
        COOZALO = i
        End If
         If UCase(asx) = "KOL" Then
        cooskkol = i
        End If
        
         If UCase(asx) = "ZALOGA" Then
        cooszza = i
        End If
        
        If UCase(asx) = "ZNESEK" Or UCase(asx) = "MADAMPCD" Or UCase(asx) = "MADANABC" Then
        cooznes = i
        End If
        If UCase(asx) = "SIFRA" Or UCase(asx) = "MADASIFR" Then
        UREJAJ = i
        End If
        If UCase(asx) = "ID_DOK" Then
        UR_id = i
        End If
Next
'If cooznes <> 0 Then
  Dim cenn As Double
  Dim ZAL, skkk As Double
        With MSHFlexGrid1
       ' MsgBox fgtrial.TextMatrix(lCount, coollce)
        .Redraw = False ' makes it about 10x faster !
       
        For lcount = .FixedRows To .Rows - 1
           'cena
          If MSHFlexGrid1.Rows > 1 Then
          ' MsgBox MSHFlexGrid1.TextMatrix(lCount, cooznes)
          If cooznes <> 0 Then
          
          If MSHFlexGrid1.TextMatrix(lcount, cooznes) = "" Then
          cenn = 0
          Else
         'cenn = MSHFlexGrid1.TextMatrix(lCount, cooznes)
       '  MsgBox cooznes
         '  cenn = Replace(IIf(MSHFlexGrid1.TextMatrix(lCount, cooznes) = "", "0", MSHFlexGrid1.TextMatrix(lCount, cooznes)), ".", ",")
          ' MsgBox MSHFlexGrid1.TextMatrix(lCount, cooznes)
           cenn = Replace(Replace(MSHFlexGrid1.TextMatrix(lcount, cooznes), ",", ""), ".", ",")
         
           End If
             MSHFlexGrid1.TextMatrix(lcount, cooznes) = FormatNumber(cenn, 2)
          skupzn = skupzn + FormatNumber(cenn, 2)
             End If
             If cooskkol <> 0 Then
             
             If MSHFlexGrid1.TextMatrix(lcount, cooskkol) <> "" Then
             skkk = Replace(Replace(MSHFlexGrid1.TextMatrix(lcount, cooskkol), ",", ""), ".", ",")
             MSHFlexGrid1.TextMatrix(lcount, cooskkol) = FormatNumber(skkk, 2)
             skupko = skupko + FormatNumber(skkk, 2)
             End If
             End If
           
           If COOZALO <> 0 Then

ZAL = Replace(IIf(MSHFlexGrid1.TextMatrix(lcount, COOZALO) = "", "0", MSHFlexGrid1.TextMatrix(lcount, COOZALO)), ".", ",")
          ZAL = Replace(MSHFlexGrid1.TextMatrix(lcount, COOZALO), ".", ",")
          zall = zall + Replace(MSHFlexGrid1.TextMatrix(lcount, cooskkol), ".", ",")
If ZAL <> "" Then
             MSHFlexGrid1.TextMatrix(lcount, COOZALO) = FormatNumber(ZAL, 2)
             skupce = skupce + FormatNumber(ZAL, 2)
            End If
If zall <> "" Then
             MSHFlexGrid1.TextMatrix(lcount, cooszza) = FormatNumber(zall, 2)
             szza = szza + FormatNumber(zall, 2)
            End If
            
            End If
         If coolldat <> 0 Then
         '    If MSHFlexGrid1.TextMatrix(lcount, coolldat) <> "" Then
              'MSHFlexGrid1.TextMatrix(lCount, coolldat) = Format(MSHFlexGrid1.TextMatrix(lCount, coolldat), "long date")
             
          '   End If
             End If
              
            End If
        Next
       
        .ColAlignment(cooznes) = flexAlignRightCenter
        .ColAlignment(COOZALO) = flexAlignRightCenter
            .ColAlignment(cooszza) = flexAlignRightCenter
    
        .Redraw = True ' dont forget to do this !
        
        End With
End If

 MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
 MSHFlexGrid1.Row = MSHFlexGrid1.Rows - 1
 MSHFlexGrid1.Col = 1
    MSHFlexGrid1.ColSel = MSHFlexGrid1.Cols() - MSHFlexGrid1.FixedCols - 1
         
 MSHFlexGrid1.CellBackColor = 255
 MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1) = "SKUPAJ"
 MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, cooskkol) = skupko
 MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, COOZALO) = skupce
 MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, cooznes) = skupzn
 
 
 Call FG_AutosizeCols(MSHFlexGrid1, Me, , , False)

End Sub

Private Sub LaVolpeButton1_Click()
Dim dodno, dodn, sqwh As String
Dim commd, ccc As String
naslo = " Analiza po :"
commd = ""
ccc = ""
dodn = ""

dodno = ""
sqwh = ""
Dim das, des
das = Format(Me.DATOD.Value, "dd.mm.yyyy")
des = Format(Me.DATDO.Value, "dd.mm.yyyy")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
ddo = Mid(des, 4, 2) & "/" & Left(des, 2) & "/" & Mid(des, 7, 4)
For bb = 0 To skly
commd = "nabasif." & Trim(Me.Combo1(bb).Text)
ccc = Trim(Me.Combo1(bb).Text)
naslo = naslo & ccc
If Me.UserControl11(bb).BoundDatax <> "" Then
naslo = naslo & "(vsebuje " & Me.UserControl11(bb).BoundDatax & "),"
Else
naslo = naslo & ","
End If
If ccc <> "" Then

If UCase(ccc) = "ID_DOK" Then
Dim part As String

commd = "nabasif.id_dok, min(glavna.dod0) as partner"

End If
If UCase(ccc) = "SIFRA" Then
commd = "nabasif.sifra,min(nabasif.naziv) as naziv"
End If


If dodno = "" Then
dodno = commd & ","
Else
dodno = dodno & commd & ","
End If
If dodn = "" Then
dodn = "nabasif." & Trim(Me.Combo1(bb).Text) & ","
Else
dodn = dodn & "nabasif." & Trim(Me.Combo1(bb).Text) & ","
End If

If Me.UserControl11(bb).BoundDatax <> "" Then
If sqwh = "" Then
sqwh = " where " & "nabasif." & Trim(Me.Combo1(bb).Text) & " in ('" & Replace(Trim(Me.UserControl11(bb).BoundDatax), ",", "','") & "')"
Else
sqwh = sqwh & " and " & "nabasif." & Trim(Me.Combo1(bb).Text) & " in ('" & Replace(Trim(Me.UserControl11(bb).BoundDatax), ",", "','") & "')"

End If
End If
End If
'MsgBox sqwh
Next bb
If Me.Check1 Then
naslo = naslo & vbCrLf & "Za obdobje od: " & Format(Me.DATOD.Value, "dd.mm.yyyy") & " do: " & Format(Me.DATDO.Value, "dd.mm.yyyy")
End If
If rs.State = 1 Then rs.Close

sxqll = "select space(1) as q," & dodno & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl,sum(format(nabasif.kol*nabasif.faktor,'#,##0.00')) as kol,sum(format(nabasif.kol*0,'#,##0.00')) as ZALOGA,sum(format((nabasif.cena-(nabasif.cena*(1-(iif(nabasif.pop=0,100,nabasif.pop)/100)))),'#,##0.00')) as cena,sum(format((nabasif.cena-(nabasif.cena*(1-(iif(nabasif.pop=0,100,nabasif.pop)/100))))*nabasif.kol,'#,##0.00')) as znesek from nabasif LEFT JOIN glavna ON (nabasif.id_dok = glavna.id_dok) AND (nabasif.tip_dok = glavna.tip_dok) " & sqwh & " group by " & dodn & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl order by  " & dodn & "nabasif.DATUM,nabasif.tip_dok,nabasif.skl"
If Me.DATDO.Enabled = True Then
If sqwh = "" Then
sxqll = Replace(sxqll, "group", "where nabasif.DATUM between #" & dod & "# And  #" & ddo & "# group")
Else
sxqll = Replace(sxqll, "where", "where nabasif.DATUM between #" & dod & "# And  #" & ddo & "# and")
End If

End If
sxqll = Replace(sxqll, "where", "where nabasif.faktor<>0 and")

rs.Open sxqll, myConection, adOpenDynamic, adLockOptimistic


'koli = koli + RS.Fields("kol")
'popu = popu + ((RS.Fields("cena") - (RS.Fields("cena") * (1 - (RS.Fields("pop") / 100)))) * RS.Fields("kol"))

If rs.EOF Then
   Me.MSHFlexGrid1.Visible = False
    
   
Else
   Me.MSHFlexGrid1.Visible = True
    
    Set Me.MSHFlexGrid1.DataSource = rs
     
End If

osve
   


End Sub

Private Sub Pregle_Click()
If Me.MSHFlexGrid1.Visible = True Then
Call Print_osn(naslo, Me.MSHFlexGrid1)
End If
End Sub
