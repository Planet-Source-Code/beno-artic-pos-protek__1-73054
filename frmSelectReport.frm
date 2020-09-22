VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTT~1.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form izpisi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IZPISI"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   HelpContextID   =   1014
   Icon            =   "frmSelectReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelectReport.frx":038A
   ScaleHeight     =   5805
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "DODAJ"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":3389B
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
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "BRIŠI"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":338B7
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
   Begin LVbuttons.LaVolpeButton cmdRun 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "PRIKAŽI"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":338D3
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
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "KOPIRAJ"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":338EF
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
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UREDI"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":3390B
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      FocusRect       =   0
      AllowUserResizing=   1
      FormatString    =   "        |NAZIV IZPISA                                                                                  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   275
      Left            =   75
      TabIndex        =   6
      Top             =   0
      Width           =   11200
      _ExtentX        =   19764
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   375
      Left            =   3500
      TabIndex        =   7
      Top             =   5280
      Width           =   500
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Def"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSelectReport.frx":33927
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
Attribute VB_Name = "izpisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDelete_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        boolConfirm = MsgBox("Are you sure you want to delete this Reports Factory Template : " & MSFlexGrid1.TextMatrix(lngX, 1) & " ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
            fso.DeleteFile App.path & "\Publish\" & MSFlexGrid1.TextMatrix(lngX, 1), True
            If lngX > 1 Then
                MSFlexGrid1.RemoveItem lngX
            Else
                MSFlexGrid1.TextMatrix(1, 0) = ""
                MSFlexGrid1.TextMatrix(1, 1) = ""
                MSFlexGrid1.TextMatrix(1, 2) = ""
                MSFlexGrid1.TextMatrix(1, 3) = ""
                
            End If
            MsgBox "Report template deleted", vbInformation
        Else
            MSFlexGrid1.Redraw = True
            Exit Sub
        End If
        MSFlexGrid1.Redraw = True
        Exit Sub
    End If
    lngX = lngX + 1
Wend

Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "Please select a report to delete", vbExclamation

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
Call ShowAppHelp(1014)
End Sub

Private Sub cmdPrevious_Click()

End Sub

Private Sub cmdRun_Click()
Dim lngX As Long
Dim reporx As String
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
     reporx = MSFlexGrid1.TextMatrix(lngX, 1)
    If Left(MSFlexGrid1.TextMatrix(lngX, 1), 1) = "*" Then
    repor = Mid(MSFlexGrid1.TextMatrix(lngX, 1), 3)
    Else
        repor = MSFlexGrid1.TextMatrix(lngX, 1)
    End If
       'ured_izp.Show
      ' Call PrintFlexix(repor)
     ' MsgBox repor
     If Val(Getnazi("select pozicija from izpisi where naziv='" & reporx & "'")) = 1 Then
      Call Print_dob(reporx)
     End If
     If Val(Getnazi("select pozicija from izpisi where naziv='" & reporx & "'")) = 2 Then
      Call Print_preg(repor)
     End If
    If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 3 Then
      zalo.Show
     End If
      If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 4 Then
      Call Print_osn(repor, frmControlMain.MSHFlexGrid1)
     End If
       If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 5 Then
      Call Print_zal_fifo(repor)
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 7 Then
      Call PrintFlexix(repor)
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 10 Then
     If CatalogueName = "Category" Then
frmControlMain.MSHFlexGrid1.Col = UREJAJ
MODIFYID = frmControlMain.MSHFlexGrid1.Text
End If
If MODIFYID <> "" Then
myConection.Execute ("delete from bbe")

myConection.Execute ("Insert into bbe SELECT * from zaloga where sifra='" & MODIFYID & "' order by datum,id_dok,poz")
If rs.State = 1 Then rs.Close
rs.Open "select * from bbe", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF() Then
rs.MoveFirst
Dim aha As Double
Dim vre As Double
aha = 0
vre = 0
Do While Not rs.EOF
'rst.Edit
If rs.Fields("tip_dok") = "NA" Then
rs.Fields("kon") = rs.Fields("kol")

rs.Fields("vri") = 0
rs.Fields("vrn") = rs.Fields("vrednost")

rs.Fields("koi") = 0

End If
If rs.Fields("tip_dok") = "IZ" Then

rs.Fields("vri") = rs.Fields("vrednost")
rs.Fields("vrn") = 0

rs.Fields("koi") = rs.Fields("kol")
rs.Fields("kon") = 0
End If
aha = aha + Round(rs.Fields("kon") + rs.Fields("koi"), 3)
vre = vre + Round(rs.Fields("vrn") + rs.Fields("vri"), 3)

rs.Fields("prosta") = aha
rs.Fields("prostav") = vre
rs.Update
rs.MoveNext
Loop
End If

     PRINTSNAP repor, ""
     'sgBox (MODIFYID)
End If
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 8 Then
     PRINTSNAP repor, "tip_dok='" & tip_dok & "' and id_dok='" & xid_dok & "'"
     End If
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 9 Then
     
     'MsgBox "datum>=#" & Replace(frmControlMain.datod.Value, ".", "-") & "# and datum<=#" & Replace(frmControlMain.datdo.Value, ".", "-") & "#"
     'MsgBox "datum>=#" & Format(frmControlMain.datod.Value, "mm/dd/yyyy") & "# and datum<=#" & Format(frmControlMain.datdo.Value, "mm/dd/yyyy") & "#"
     If repor = "DNEVNIK" Then
     PRINTSNAP repor, "datum>=#" & Replace(Format(frmControlMain.DATOD.Value, "mm/dd/yyyy"), ".", "/") & "# and datum<=#" & Replace(Format(frmControlMain.DATDO.Value, "mm/dd/yyyy"), ".", "/") & "#"
     Else
     PRINTSNAP repor, "tip_dok='" & tip_dok & "' and datum>=#" & Replace(Format(frmControlMain.DATOD.Value, "mm/dd/yyyy"), ".", "/") & "# and datum<=#" & Replace(Format(frmControlMain.DATDO.Value, "mm/dd/yyyy"), ".", "/") & "#"
     End If
     End If
     
     If Val(Getnazi("select pozicija from izpisi where  naziv='" & reporx & "'")) = 6 Then
      Call Print_dob_les(repor)
     End If
     Unload Me
       Exit Sub
        Else
            
        End If
    
    lngX = lngX + 1
Wend
'Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "PROSIM IZBERI SI IZPIS LEVO KLIK DA SE PRIKAŽE KLUKICA !", vbExclamation


End Sub

Private Sub Form_Activate()
Dim oFolder, oFiles, oFile
Dim myFile
Dim strRptName As String, strRptDesc As String
Dim strFileRead As String
MSFlexGrid1.Rows = 1
MSFlexGrid1.Rows = 2
    If rs.State = 1 Then rs.Close
    rs.Open "Select * from izpisi where tip_dok='" & tip_dok & "' or tip_dok='VS'", myConection, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
             MSFlexGrid1.AddItem "" & vbTab & rs.Fields("naziv"), 1
             
    rs.MoveNext
    Loop
    End If
On Error Resume Next
MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1

Dim lngX As Long
lngX = 1
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    MSFlexGrid1.Row = lngX
    MSFlexGrid1.CellFontName = "Wingdings"
    lngX = lngX + 1
Wend

lngX = 1
While lngX + 1 <= MSFlexGrid1.Cols
    MSFlexGrid1.ColAlignment(lngX) = 1
    lngX = lngX + 1
Wend

Call AltFlexiColors(MSFlexGrid1, 1, 1)
End Sub
Private Sub refr()
Dim oFolder, oFiles, oFile
Dim myFile
Dim strRptName As String, strRptDesc As String
Dim strFileRead As String
MSFlexGrid1.Rows = 1
MSFlexGrid1.Rows = 2
    If rs.State = 1 Then rs.Close
    rs.Open "Select * from izpisi where tip_dok='" & tip_dok & "'", myConection, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
             MSFlexGrid1.AddItem "" & vbTab & rs.Fields("naziv"), 1
             
    rs.MoveNext
    Loop
    End If
On Error Resume Next
MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1

Dim lngX As Long
lngX = 1
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    MSFlexGrid1.Row = lngX
    MSFlexGrid1.CellFontName = "Wingdings"
    lngX = lngX + 1
Wend

lngX = 1
While lngX + 1 <= MSFlexGrid1.Cols
    MSFlexGrid1.ColAlignment(lngX) = 1
    lngX = lngX + 1
Wend

Call AltFlexiColors(MSFlexGrid1, 1, 1)

End Sub


Private Sub LaVolpeButton1_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
Do While Not lngX = MSFlexGrid1.Rows

    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        repor = MSFlexGrid1.TextMatrix(lngX, 1)
       ured_izp.Show
       Exit Sub
        
        End If
    
    lngX = lngX + 1
Loop
'Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "PROSIM IZBERI SI IZPIS LEVO KLIK DA SE PRIKAŽE KLUKICA !", vbExclamation


End Sub

Private Sub LaVolpeButton2_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        strTemplateFileName = MSFlexGrid1.TextMatrix(lngX, 1)
         boolConfirm = MsgBox("Ali želiš kopirati ta Izpis?? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
        myConection.Execute ("insert into izpisi_r select * from izpisi where  naziv='" & LTrim(strTemplateFileName) & "'")
        Dim novo As String
        Dim novapoz As Integer
        novapoz = 1 + Val(Getnazi("select pozicija from izpisi_r where naziv='" & LTrim(strTemplateFileName) & "'"))
        novo = InputBox("Report morate na novo poimenovati", "Vnesi nov naziv reporta")
        myConection.Execute ("update izpisi_r set pozicija=" & novapoz & " where naziv='" & LTrim(strTemplateFileName) & "'")
        If Getnazi("select naziv from izpisi_r where  naziv='" & LTrim(novo) & "'") = "" Then
        Else
        novo = novo & "1"
        End If
        myConection.Execute ("update izpisi_r set naziv='" & novo & "' where naziv='" & LTrim(strTemplateFileName) & "'")
             myConection.Execute ("insert into izpisi select * from izpisi_r where naziv='" & novo & "'")
   
        myConection.Execute ("delete from izpisi_r where  naziv='" & novo & "'")
        
        End If
        refr
           Exit Sub
        
        End If
    
    lngX = lngX + 1
Wend
'Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "PROSIM IZBERI SI IZPIS LEVO KLIK DA SE PRIKAŽE KLUKICA !", vbExclamation

End Sub

Private Sub LaVolpeButton3_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        strTemplateFileName = MSFlexGrid1.TextMatrix(lngX, 1)
         boolConfirm = MsgBox("Ali želiš izbrisati ta Izpis?? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
        myConection.Execute ("delete from izpisi where  naziv='" & LTrim(strTemplateFileName) & "'")
        End If
        refr
          Exit Sub
       
        End If
    
    lngX = lngX + 1
Wend
'Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "PROSIM IZBERI SI IZPIS LEVO KLIK DA SE PRIKAŽE KLUKICA !", vbExclamation

End Sub

Private Sub LaVolpeButton4_Click()
'templati.Show
End Sub

Private Sub LaVolpeButton5_Click()
Dim lngX As Long
lngX = 1
MSFlexGrid1.Redraw = False
MSFlexGrid1.Col = 0
While lngX + 1 <= MSFlexGrid1.Rows
    If MSFlexGrid1.TextMatrix(lngX, 0) <> "" Then
        strTemplateFileName = MSFlexGrid1.TextMatrix(lngX, 1)
         boolConfirm = MsgBox("Ali želiš nastaviti ta izpis za DEFAULT IZPIS?? ?", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
        If boolConfirm = vbYes Then
        myConection.Execute ("update izpisi set naziv=mid(naziv,3) where  naziv like '* %' and tip_dok='" & tip_dok & "'")
        myConection.Execute ("update izpisi set naziv='* " & LTrim(strTemplateFileName) & "' where  naziv='" & LTrim(strTemplateFileName) & "' and tip_dok='" & tip_dok & "'")
        End If
        refr
          Exit Sub
       
        End If
    
    lngX = lngX + 1
Wend
'Call AltFlexiColors(MSFlexGrid1, 1, 1)
MSFlexGrid1.Redraw = True

MsgBox "PROSIM IZBERI SI IZPIS LEVO KLIK DA SE PRIKAŽE KLUKICA !", vbExclamation

End Sub

Private Sub MSFlexGrid1_Click()
With MSFlexGrid1
.Redraw = False

If .Row = 0 Or .MouseCol <> 0 Then
    'nothing
Else
    Dim lngTemp As Long
    Dim boolTicked As Boolean
    lngTemp = MSFlexGrid1.Row

    If .TextMatrix(lngTemp, 0) = "ü" Then
        boolTicked = True
        .TextMatrix(lngTemp, 0) = ""
    Else
        boolTicked = False
        .TextMatrix(lngTemp, 0) = "ü"
    End If
    
    Dim lngX As Long
    lngX = 1
    MSFlexGrid1.Col = 0
    While lngX + 1 <= MSFlexGrid1.Rows
        MSFlexGrid1.TextMatrix(lngX, 0) = ""
        lngX = lngX + 1
    Wend
    
    If boolTicked = True Then
        .TextMatrix(lngTemp, 0) = ""
    Else
        .TextMatrix(lngTemp, 0) = "ü"
    End If
End If
.Redraw = True

End With
End Sub

Private Sub MSFlexGrid1_DblClick()
If MSFlexGrid1.Col = 0 Then
    Exit Sub
End If

Call SortFlexiArrows(MSFlexGrid1, False, False)
Call AltFlexiColors(MSFlexGrid1, 1, 1)
End Sub

Function checkPassword() As Boolean
checkPassword = True
Dim strPassword As String, strEnterPwd As String
Dim myFile, strFileRead
strPassword = "": strEnterPwd = ""

Set myFile = fso.OpenTextFile(App.path & "\Publish\" & strTemplateFileName, 1, -2)
strFileRead = Trim(myFile.ReadLine)
While myFile.AtEndOfStream <> True
    If Left(UCase(strFileRead), Len("Password:=")) = UCase("Password:=") Then
        strPassword = Mid(strFileRead, Len("Password:=") + 1)
        If Len(strPassword) > 0 Then
            strEnterPwd = InputBox("This report is password protected. You have to enter the password to run this report", "Enter Password")
            If strEnterPwd <> strPassword Then
                MsgBox "The password entered by you is invalid", vbExclamation
                checkPassword = False
                Exit Function
            End If
        End If
    End If
    strFileRead = Trim(myFile.ReadLine)
Wend
Set myFile = Nothing
End Function
