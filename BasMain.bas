Attribute VB_Name = "BasMain"
Option Explicit
Public dodatni_ar As String
Public DOD_AR As String
Public Pblagajna As String
Public trenslika As String
Public resis, ohonac, idtipk As Integer
Public bremepis As Integer
Public nazalogi As String
Public tiskdol As Double
Public vrjenniz, prijavljen, id_inv As String
Public pritisk, idpo, stevnaro As String
Public UREJAJ, UR_id, zapore, izp_fifx, prvaa, dell As Integer
Public repor As String
Public jedobavnica As String
Public xpozi As String
Public UPORABNIK As String
Public sxqll As String
Public imedn, sqt, xizb As String
Public intCtrlDown, kjje As Integer
Public dtip_dok, SQLREP As String
Public prvay As Integer
Public ma_ko, xpox, xzago As Integer
Public fr As String
Public xEM, xskladd As String
Public tresi, xrep As String
Public kater As Integer
Public Xvs, Yvs As Integer
Public vmessql As String
Public PathFileName, tipa As String
Public myConection As New ADODB.Connection
Public rs As New ADODB.Recordset
Public cPrint As clsMultiPgPreview
Public sl, sx, erro As String
Public plax, odg As String
Public atab As String
Public odpr As String
Public iskalni As String
Public normati As String
Public xid_dok
Public xopis
Public ve_ro As Integer
Public zaposle As String
Public fso As New FileSystemObject
Public dodajwh As String
Public avtomob As String
Public relacij, idfx, idko, printsql, PRINTREP As String
Public coollna, coollce, coollko, coollmarz, coollmpc, coolldat_k, coollur, coollzn, coollzal, coollsi, coollem, coollpro, coolldat, collchk, cooznes, coollx, coolly, COOZALO, coollpop, coollles, coollstek, cooskkol As Integer
Public zai As Long
Public zaix As Long
Public ssqq As String
Public izja As Integer
Public kosovni As Integer
Public ma_ured As Integer
Public trenu As Integer
Public visina As Long
Public webhw As Long
Public dolzina As Long
Public sifrt As String
Public knjiz
Public osve As Integer
Public odprt As Integer
Public podatek As String
Public sqlb As String
Public ss As String
Public aavr As Integer
Public vvvv As Integer
Public tip_dok
Public uda As String
Public kolik As Integer
Public idar As String
Public xxre As String
Public bepr As Integer
Public odp As Integer
Public dara As Integer
Public skumi, xpopu As Double
Public uredira As Integer
Public zavrnit As Double
Public zap As Integer
Public opp As Integer
Public oppa As Integer
Public idstran As Integer
Public jestran As Integer
Public refr As Integer
Public tSup As String
Public odprta As Integer
Public nivo As Integer
Public fora As Integer
Public OSEB As String
Public blagajna As Integer
Public OSE As String
Public deln As Integer
Public ber As Integer
Public izbrko As Integer
Public siff As Long
Public uredi As Integer
Public dod As String
Public ddo As String
Public edi As Integer
Public std As String
Public dob As Integer
Public stm1 As Integer
Public dejpre As Integer
Public counter As Double
Public tis_a, tis_b, tis_c, tis_d, tis_e As Double
Public stalnaprij As Integer

Public QuitCommand As Boolean

Public Declare Function GetTickCount Lib "kernel32" () As Long

Option Compare Text


Private Declare Function EbExecuteLine Lib "vba6.dll" _
        (ByVal pStringToExec As Long, ByVal Foo1 As Long, _
        ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long

  'resize.
Public ind As Integer
Dim RsNewNo As New ADODB.Recordset

Public DBPathFileName As String

Dim Inti, i

Public Function LoadFileToTB(TxtBox As Object, FilePath As _
   String, Optional Append As Boolean = False) As Boolean
 

Dim iFile As Integer
Dim s As String

If Dir(FilePath) = "" Then Exit Function

On Error GoTo ErrorHandler:
s = TxtBox.Text

iFile = FreeFile
Open FilePath For Input As #iFile
s = Input(LOF(iFile), #iFile)
If Append Then
    TxtBox.Text = TxtBox.Text & s
Else
    TxtBox.Text = s
End If

LoadFileToTB = True

ErrorHandler:
If iFile > 0 Then Close #iFile

End Function

Public Function SaveFileFromTB(TxtBox As Object, _
   FilePath As String, Optional Append As Boolean = False) _
   As Boolean
  


Dim iFile As Integer

iFile = FreeFile
If Append Then
    Open FilePath For Append As #iFile
Else
    Open FilePath For Output As #iFile
End If

Print #iFile, TxtBox.Text
SaveFileFromTB = True

ErrorHandler:

Close #iFile
End Function

Function FExecuteCode(stCode As String, Optional fCheckOnly _
    As Boolean) As Boolean
    On Error GoTo bnn:
    FExecuteCode = EbExecuteLine(StrPtr(stCode), 0&, 0&, _
        Abs(fCheckOnly)) = 0
bnn:
End Function
Public Function Repl(sString As Variant _
          , sFind As String _
          , sReplace As String _
          , Optional iCompare As Long = vbBinaryCompare) As Variant

    Dim iStart As Integer, iLength As Integer
    If IsNull(sString) Then
        Repl = Null
    Else
        iStart = InStr(1, sString, sFind, iCompare)
        Do While iStart > 0
            sString = Left(sString, iStart - 1) _
                    & sReplace & Mid(sString, iStart + Len(sFind), Len(sString) - iStart - Len(sFind) + 1)
            iStart = InStr(iStart, sString, sFind)
        Loop
        Repl = sString
    End If

End Function
Public Function obstaja(tabela As String) As Boolean

  On Error GoTo neobstaja
  If rs.State = 1 Then rs.Close
  rs.Open "select * from " & tabela, myConection, adOpenStatic, adLockOptimistic
rs.Close
  obstaja = True
  Exit Function
  
neobstaja:
  obstaja = False
  
End Function
Public Function jefield(tabela As String, Field As String) As Boolean

  On Error GoTo neobstaja
  If rs.State = 1 Then rs.Close
  rs.Open "select " & Field & " from " & tabela, myConection, adOpenStatic, adLockOptimistic
rs.Close
  jefield = True
  Exit Function
  
neobstaja:
  jefield = False
  
End Function
Public Function tx(ind As Integer) As Double
Dim ree As Boolean
On Error GoTo bnn:
tx = 0
Dim dfs As String
dfs = "placa.label22.caption = Placa.Text1(" & ind & ").text"
'MsgBox dfs
ree = FExecuteCode(dfs)


bnn:
End Function
Public Function fixx(cx As Double) As Double
fixx = 0
fixx = cx

End Function
Public Function Prevod(cmd As String) As Variant
    Dim i As Long
    Dim J As Long
    Dim str As String
    Dim valVar As Variant

    If (InStr(1, cmd, "Format(")) Then
        i = InStr(1, cmd, "Format(") + 7
        J = InStr(i + 1, cmd, ",")
        valVar = Prevod(Mid(cmd, i + 1, J - i - 1))
        i = InStr(J, cmd, """")
        J = InStr(i + 1, cmd, """")
        str = Mid(cmd, i + 1, J - i - 1)
        Prevod = Format(valVar, str)
    Else
        Prevod = cmd
    End If
End Function

Public Function tx2(ind As Integer) As Double
Dim ree As Boolean
On Error GoTo bnn:
tx2 = 0
Dim dfs As String
dfs = "placa.label22.caption = Placa.Text2(" & ind & ").text"
'MsgBox dfs
ree = FExecuteCode(dfs)


bnn:
End Function

Public Sub ime_form()
If Getnazi("select min(poz) from dokm where atribut='STAP'") = "" Then
Dim ads As Integer
ads = 1
For ads = 1 To 5
myConection.Execute ("insert into dokm (atribut,poz) values ('STAP'," & ads & ")")
Next
End If
If Getnazi("select min(sifra) from skla") = "" Then
If rs.State = 1 Then rs.Close
rs.Open "select * from skla", myConection, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveFirst
Dim aax As Integer
aax = 1
Do While Not rs.EOF
rs.Fields("sifra") = Trim(str(aax))
aax = aax + 1
rs.MoveNext
Loop

End If
End If
If Getnazi("select id_dok from dokm where atribut='FORM'") = "" Then
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','partner','c_frmcustomer')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','grupe','C_frmLocation')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','merske','eme')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','tip_art','tip_art')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','skladisce','skladisce')")

myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','relacije','relacije')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','vrsta poste','vr_po')")
myConection.Execute ("insert into dokm (atribut,id_dok,tekst) values ('FORM','zaposleni','zaposleni')")
End If
End Sub
Private Sub Main1()

PathFileName = App.path + "\DATABASE\Thesis.mdb"
myConection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="
If myConection.State = adStateOpen Then
   frmsalesbill.Show
Else
   MsgBox "Error in Connecting Database please check Connection", vbCritical
End If
End Sub
Public Function FillC_(cmb As ComboBox, strSQl As String)
If rs.State = 1 Then rs.Close

rs.Open strSQl, myConection
If Not rs.EOF Then
    cmb.clear
    rs.MoveFirst
    
    Do While Not rs.EOF
        With rs
        If Not IsNull(.Fields(0)) Then
        For i = 1 To Len(Trim(.Fields(0))) Step 3
            cmb.AddItem Mid(Trim(.Fields(0)), i, 2)
        Next i
        End If
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function FillCombo(cmb As ComboBox, strSQl As String)
If rs.State = 1 Then rs.Close

rs.Open strSQl, myConection
If Not rs.EOF Then
    cmb.clear
    rs.MoveFirst
    Do While Not rs.EOF
        With rs
            cmb.AddItem .Fields(0)
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Ficombo(cmb As ComboBox, strSQl As String)
If rs.State = 1 Then rs.Close

rs.Open strSQl, myConection
If Not rs.EOF Then
    cmb.clear
    rs.MoveFirst
    Do While Not rs.EOF
        With rs
            cmb.AddItem presled(.Fields(0), 10) & .Fields(1)
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function FillCom_fields(cmb As ComboBox, ytable As String, n1 As String, n2 As String, n3 As String, n4 As String, n5 As String)
If rs.State = 1 Then rs.Close
Dim fld As ADODB.Field
rs.Open ytable, myConection

    cmb.clear
    For Each fld In rs.Fields
    'MsgBox fld.Name
    If fld.Type = 202 Then
    'MsgBox Trim(UCase(n1))
    If Trim(UCase(fld.Name)) <> Trim(UCase(n1)) Then
    If Trim(UCase(fld.Name)) <> Trim(UCase(n2)) Then
    If Trim(UCase(fld.Name)) <> Trim(UCase(n3)) Then
    If Trim(UCase(fld.Name)) <> Trim(UCase(n4)) Then
    If Trim(UCase(fld.Name)) <> Trim(UCase(n5)) Then
            cmb.AddItem presled(fld.Name, 20)
        End If
        End If
        End If
        End If
        End If
        
        End If
        Next fld
    cmb.AddItem presled("Partner", 20)

End Function
Public Function GetNewNo(SQL As String) As Long
    Dim id As Long
    Set RsNewNo = myConection.Execute(SQL)
    If Not IsNull(RsNewNo(0)) Then
        id = RsNewNo(0).Value
    Else
        id = 10000
    End If
    GetNewNo = id
    RsNewNo.Close
    Set RsNewNo = Nothing
End Function

Public Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint As Boolean = False, Optional AllowNegative As Boolean = False) As Boolean
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
        If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
        If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
            KeyAscii = False
            IsNumber = False
        Else
            IsNumber = True
        End If
    Else
        IsNumber = IsNumeric(CheckString)
    End If
End Function

Public Function GetN(SQL As String) As Long
    Dim id As Long
    Set RsNewNo = myConection.Execute(SQL)
    
    GetN = id
    RsNewNo.Close
    Set RsNewNo = Nothing
End Function
Public Function Getdo(SQL As String) As String
    Dim naz As String
   ' MsgBox SQL
    Set RsNewNo = myConection.Execute(SQL)
    If Not RsNewNo.EOF Then
        If IsNull(RsNewNo(0).Value) Then
            naz = ""
        Else
           RsNewNo.MoveFirst
            Do While Not RsNewNo.EOF
              If naz = "" Then
                naz = RsNewNo(0).Value
              Else
                naz = naz & "," & RTrim(RsNewNo(0).Value) & RTrim(RsNewNo(1).Value)
              End If
            RsNewNo.MoveNext
            Loop
            End If
    Else
        naz = ""
    End If
    Getdo = naz
    RsNewNo.Close
    Set RsNewNo = Nothing
End Function


Public Function DEKODIR(sifra As String) As Integer
Dim s As Integer
DEKODIR = 0
On Error GoTo addr:
s = 1
Do While s < Len(sifra)
If Mid(sifra, s, 1) = "0" Then
DEKODIR = DEKODIR + 999
End If
If Mid(sifra, s, 1) = "1" Then
DEKODIR = DEKODIR - 1
End If
If Mid(sifra, s, 1) = "2" Then
DEKODIR = DEKODIR + 22
End If
If Mid(sifra, s, 1) = "3" Then
DEKODIR = DEKODIR + 13
End If
If Mid(sifra, s, 1) = "4" Then
DEKODIR = DEKODIR + 11
End If
If Mid(sifra, s, 1) = "5" Then
DEKODIR = DEKODIR - 181
End If
If Mid(sifra, s, 1) = "6" Then
DEKODIR = DEKODIR + 131
End If
If Mid(sifra, s, 1) = "7" Then
DEKODIR = DEKODIR + 791
End If
If Mid(sifra, s, 1) = "8" Then
DEKODIR = DEKODIR + 4
End If
If Mid(sifra, s, 1) = "9" Then
DEKODIR = DEKODIR + 1
End If


s = s + 1
Loop
addr:
    
'    DEKODIRAJ = Getnazi("select " & kaj & " from tiskal")

End Function
Public Function stevilka(st As String) As String
  
  On Error GoTo addr:
  
  stevilka = Trim(Replace(CDbl(st), ",", "."))
addr:
 
End Function
Public Function Getnazi(SQL As String) As String
    Dim naz As String
   'MsgBox SQL
  On Error GoTo addr:
    Set RsNewNo = myConection.Execute(SQL)
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    naz = ""
    Else
        naz = RsNewNo(0).Value
        End If
    Else
        naz = ""
    End If
    Getnazi = naz
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function
Public Function Getnumb(SQL As String) As Double
    Dim naz As Double
   'MsgBox SQL
  Getnumb = 0
  On Error GoTo evb:
    Set RsNewNo = myConection.Execute(SQL)
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    naz = 0
    Else
        naz = RsNewNo(0).Value
        End If
    Else
        naz = 0
    End If
    Getnumb = naz
    RsNewNo.Close
    Set RsNewNo = Nothing
evb:
End Function
Public Function Getcena(yident As String, datum As Date) As Double
    Dim nazy, XDOO, xdoz As Double
    Dim sifra As String
    nazy = 0
    ' MsgBox SQL
 Getcena = 0
  Dim das As String
  das = Format(datum, "dd.mm.yyyy")
  dod = novast(LTrim(LTrim(str(Month(das)))), 2) & "/" & novast(LTrim(LTrim(str(Day(das)))), 2) & "/" & LTrim(LTrim(str(Year(das))))

  'On Error GoTo evb:
  Dim rsta As New ADODB.Recordset
  Dim rstb As New ADODB.Recordset
 rsta.Open "select sifras,kol from sestavi where sifra=" & yident, myConection, adOpenDynamic, adLockOptimistic
  If rsta.EOF Then
    'If rstb.State = 1 Then rstb.Close
    
    'rstb.Open "select top 1 cena as ycennxx from nabasif where sifra='" & yident & "' and tip_dok='NA' and datum<=#" & dod & "# order by datum desc", myConection, adOpenDynamic, adLockOptimistic
    
    
  
    'If Not rstb.EOF Then
    '    If IsNull(rstb.Fields("ycennxx").Value) Then
    '    nazy = 0
    '     Else
            XDOO = Getnumb("select madadoza from mada where madasifr='" & yident & "'")
    
            If XDOO > 3 Then
               ' nazy = rstb.Fields("ycennxx").Value / XDOO
               nazy = Getnumb("select  cena as ycennxx from nabasif where sifra='" & yident & "' and tip_dok='NA' and datum<=#" & dod & "# order by datum desc") / XDOO
            Else
                'nazy = rstb.Fields("ycennxx").Value * XDOO
                nazy = Getnumb("select  cena as ycennxx from nabasif where sifra='" & yident & "' and tip_dok='NA' and datum<=#" & dod & "# order by datum desc") * XDOO
            End If
     '   End If
    'Else
    '    nazy = 0
   ' End If
    
  Else
  nazy = 0
  rsta.MoveFirst
  Do While Not rsta.EOF
  xdoz = rsta.Fields("kol")
  sifra = rsta.Fields("sifras")
  If sifra <> "" Then
  'If xdoz > 3 Then
  'naz = naz + (Getnumb("select top 1 cena from nabasif where sifra='" & sifra & "' and tip_dok='NA' and datum<=#" & dod & "# order by datum desc") / xdoz)
  'Else
  nazy = nazy + (Getnumb("select top 1 cena from nabasif where sifra='" & sifra & "' and tip_dok='NA' and datum<=#" & dod & "# order by datum desc") * xdoz)
  'End If
  End If
  rsta.MoveNext
   Set rstb = Nothing
  Loop
  
  End If
'MsgBox (naz)
    Getcena = nazy
'    rstb.Close
   
evb:
End Function

Public Function getdate(SQL As String) As Date
    Dim naz As Date
  ' MsgBox SQL
  On Error GoTo evb:
    Set RsNewNo = myConection.Execute(SQL)
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    naz = ctod("31.12.1899")
    Else
        naz = RsNewNo(0).Value
        End If
    Else
        naz = ctod("31.12.1899")
    End If
    getdate = naz
    RsNewNo.Close
    Set RsNewNo = Nothing
evb:
End Function


Public Sub clear(Frm As Form)
   Dim c As Control
      For Each c In Frm.Controls
          If TypeOf c Is TextBox Or TypeOf c Is ComboBox Then
             c.Text = ""
          End If
      Next c
End Sub

Public Function Update1(Table1 As Variant, ParamArray arr() As Variant)
If rs.State = 1 Then rs.Close
rs.Open "select * from " & Table1, myConection, 3, 3
rs.AddNew
   For i = 0 To UBound(arr())
       rs.Fields(i) = arr(i)
       
   Next
 rs.Update
 rs.Close
End Function

Public Function GetTxtVal(ByVal sTxt As String) As Double

    Dim sNew As String
    Dim sC As String
    Dim i As Integer
    
    'default
    GetTxtVal = 0
        
    sTxt = Trim(sTxt)
    
    If Len(sTxt) > 0 Then
        For i = 1 To Len(sTxt)
            sC = Mid(sTxt, i, 1)
            If sC = "-" Or sC = "," Or sC = "1" Or sC = "2" Or sC = "3" Or sC = "4" Or sC = "5" Or sC = "6" Or sC = "7" Or sC = "8" Or sC = "9" Or sC = "0" Then
                sNew = sNew & sC
            End If
        Next
    
        If Len(sNew) > 0 Then
            GetTxtVal = (sNew)
        End If
    End If
    
    
End Function
Public Function ctod(ByVal sTxt As String) As Date
On Error GoTo evb:
Dim prva, druga As Integer
If IsNumber(Mid(LTrim(sTxt), 2, 1)) Then
prva = 2
Else
prva = 1
End If
If IsNumber(Mid(LTrim(sTxt), 2 + prva + 1, 1)) Then
druga = 2
Else
druga = 1
End If
If prva = 1 Then
sTxt = "0" & LTrim(sTxt)
End If
If druga = 1 Then
sTxt = Left(sTxt, 3) & "0" & Mid(sTxt, 4)
End If
'MsgBox (sTxt)
ctod = Mid(LTrim(sTxt), 4, 2) & "/" & Left(LTrim(sTxt), 2) & "/" & Right(RTrim(sTxt), 4)
evb:
End Function
Public Function dtoc(ByVal sTxt As Date) As String
On Error GoTo evb:
Dim das, des
das = Format(sTxt, "dd.mm.yyyy")
sTxt = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
dtoc = IIf(Len(LTrim(Day(sTxt))) = 1, "0" & Trim(Day(sTxt)), Trim(Day(sTxt))) & "." & IIf(Len(LTrim(Month(sTxt))) = 1, "0" & Trim(Month(sTxt)), Trim(Month(sTxt))) & "." & Trim(Year(sTxt))
'dtoc = Format(dtoc, "dd.mm.yyyy.")
evb:
End Function
Public Function strVal(ByVal sTxt As String) As String

    Dim kolikode As Integer
    Dim kjede As Integer
    Dim dodat As String
    kolikode = 4
    'Getnazi ("select max(dol_ce) as xd from dokumenti ")
    sTxt = Replace(Trim(sTxt), ".", ",")
    kjede = 0
    dodat = ""
   For i = 1 To Len(sTxt)
  
   If Mid(sTxt, i, 1) = "," Then
   kjede = i
   End If
   Next i
   'MsgBox kjede
   If kjede = 0 Then
       strVal = sTxt & ",0000"
   Else
   'MsgBox kolikode - (Len(sTxt) - kjede)
   If Len(sTxt) - kjede < kolikode Then
   For i = 1 To kolikode - (Len(sTxt) - kjede)
   dodat = dodat & "0"
   'MsgBox dodat
   Next i
   strVal = sTxt & Trim(dodat)
  ' MsgBox strval
   Else
   strVal = sTxt
   
   End If
   End If
    
End Function
Public Function strVal2(ByVal sTxt As String) As String

    Dim kolikode As Integer
    Dim kjede As Integer
    Dim dodat As String
    kolikode = 2
    'Getnazi ("select max(dol_ce) as xd from dokumenti ")
    sTxt = Replace(Trim(sTxt), ".", ",")
    kjede = 0
    dodat = ""
   For i = 1 To Len(sTxt)
  
   If Mid(sTxt, i, 1) = "," Then
   kjede = i
   End If
   Next i
   'MsgBox kjede
   If kjede = 0 Then
       strVal2 = sTxt & ",00"
   Else
   'MsgBox kolikode - (Len(sTxt) - kjede)
   If Len(sTxt) - kjede < kolikode Then
   For i = 1 To kolikode - (Len(sTxt) - kjede)
   dodat = dodat & "0"
   'MsgBox dodat
   Next i
   strVal2 = sTxt & Trim(dodat)
  ' MsgBox strval
   Else
   strVal2 = sTxt
   
   End If
   End If
    
End Function
Public Sub PaintGrad(ByRef Obj As Object, lColor1 As Long, lColor2 As Long, iAngle As Integer)
   ' Dim cGrad As New clsGrad
    On Error Resume Next
    Obj.AutoRedraw = True
    'cGrad.Color1 = lColor1
    'cGrad.Color2 = lColor2
    'cGrad.Angle = iAngle
    'cGrad.Draw Obj
    Obj.Refresh
    'Set cGrad = Nothing
    err.clear
End Sub
Public Function Flistvel(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long
Cmbl.clear
 Cmbl.AddItem "VSE"
If Not rs.EOF Then
    
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
   
    dolg = ""
     If Not IsNull(rs.Fields(0)) Then
        dolg = Left(Trim(str(rs.Fields(0))), 4)
     End If
        With rs
         
            Cmbl.AddItem dolg
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Filllist(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
'MsgBox (strSQl)
strSQl = Replace(strSQl, "'mada", "' and mada")
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg, xnazz As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Double
Dim cen As Double
cen = 0
If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    cen = Getnumb("select madampcd from mada where madasifr='" & rs.Fields(0) & "'")
    zalo = Getnumb("select madazalo from mada where madasifr='" & rs.Fields(0) & "'")
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
     If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='OPISPO'") = "D" Then
     xnazz = Trim(Getnazi("SELECT dokm.tekst FROM nabasif INNER JOIN dokm ON (nabasif.pozicija = dokm.atribut) AND (nabasif.id_dok = dokm.id_dok) AND (nabasif.tip_dok = dokm.tip_dok) where nabasif.sifra='" & rs.Fields(0) & "' order by nabasif.datum desc"))
     End If
        dolg = presled(Trim((rs.Fields(0))), 14)
     '   zalo = Str(RS.Fields(3))
        With rs
         If sx <> "" Then
            Cmbl.AddItem dolg & " | " & presled(Left(Trim(.Fields(1)) & " " & xnazz, 50), 50) & " | " & presled(Trim(.Fields(2)), 3) & " | " & presled(FormatNumber(zalo, 2), 10) & " | " & presled(FormatNumber(cen, 4), 10)
            Else
            Cmbl.AddItem dolg & " | " & presled(Left(Trim(.Fields(1)) & " " & xnazz, 50), 50) & " | " & presled(FormatNumber(zalo, 2), 10) & " | " & presled(FormatNumber(cen, 4), 10)
            End If
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Filll(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
        dolg = presled(Trim(str(rs.Fields(0))), 10)
       ' zalo = Str(RS.Fields(3))
        With rs
        
            Cmbl.AddItem dolg & " | " & presled(Trim(.Fields(1)), 30)
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Fiil(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    
    dolg = ""
   
        dolg = presled(Trim(str(rs.Fields(0))), 6)
      
        With rs
        
            Cmbl.AddItem dolg & " | " & presled(Trim(.Fields(2)) & " " & .Fields(1), 30)
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Filipotne(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dd As Integer
Dim AAS As Integer
Dim zalo As Long

If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
        dolg = presled(Trim((rs.Fields(0))), 10)
       ' zalo = Str(RS.Fields(3))
        With rs
        
            Cmbl.AddItem dolg & " | " & presled(Trim(.Fields(1)), 30) & " | " & LTrim(str(.Fields(2)))
            
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function Dodat_a(asss As String) As String
dodatni_ar = asss
 
dodatni.Show vbModal
   Dodat_a = dodatni_ar
   dodatni_ar = ""
End Function
Public Function DoSQLbe(atable As String, first As String, secc As String, doda As String) As String

    ss = ""
    sx = ""
    If secc <> "" Then
    ss = secc
    End If
    If doda <> "" Then
    sx = doda
    End If
    'xxre=""
   sl = first
atab = atable
    'xxxsql = "sELECT " & sl & ss & " FROM " & atable & " WHERE (((" & sl & " or " & atable & "." & secc & ") Like '"
Form1.Show vbModal
'MsgBox xxre
   DoSQLbe = vrjenniz
   vrjenniz = ""
End Function

Public Function DoSQL(atable As String, first As String, secc As String, doda As String) As String

    ss = ""
    sx = ""
    If secc <> "" Then
    ss = secc
    End If
    If doda <> "" Then
    sx = doda
    End If
    'xxre=""
   sl = first
atab = atable
    'xxxsql = "sELECT " & sl & ss & " FROM " & atable & " WHERE (((" & sl & " or " & atable & "." & secc & ") Like '"
Form1.Show vbModal
'MsgBox xxre
   DoSQL = xxre
   xxre = ""
End Function
Public Sub DoSQL3(atable As String, first As String, secc As String)

    ss = ""
    If secc <> "" Then
    ss = secc
    End If
   sl = first
atab = atable
    'xxxsql = "sELECT " & sl & ss & " FROM " & atable & " WHERE (((" & sl & " or " & atable & "." & secc & ") Like '"
Form1.Show
   
End Sub
Public Sub touch()
   Dim RSt1 As New ADODB.Recordset
      Dim rst As New ADODB.Recordset
         Dim RSt2 As New ADODB.Recordset


myConection.Execute ("delete * from swit")
If ConnectRS(myConection, RSt2, "select * from swit") = True Then
End If


RSt2.AddNew
RSt2.Fields("switchboar") = 1
RSt2.Fields("itemnumber") = 0
RSt2.Fields("itemtext") = "TOUCHSCREEN"
RSt2.Fields("command") = 0
RSt2.Fields("argument") = "Default"
RSt2.Update
If ConnectRS(myConection, rst, "select * from grupa") = True Then
End If

'Set rst = myConection.OpenRecordset("grupa")
rst.MoveFirst
Do While Not rst.EOF
If rst.EOF Then
Exit Do
End If
RSt2.AddNew
RSt2.Fields("switchboar") = 1
RSt2.Fields("itemnumber") = rst.Fields("sifra")
RSt2.Fields("itemtext") = rst.Fields("grupa")
RSt2.Fields("command") = 1

RSt2.Fields("argument") = rst.Fields("sifra") + 2
RSt2.Update
rst.MoveNext
Loop
rst.MoveFirst
Dim a As Integer
a = 2
Do While Not rst.EOF
If rst.EOF Then
Exit Do
End If
If ConnectRS(myConection, RSt1, "select * from mada where madagrup=" & rst.Fields("sifra") & " order by madanazi") = True Then
End If

'Set rst1 = myConection.OpenRecordset("select * from mada where madagrup=" & rst.Fields("sifra"))
RSt2.AddNew
RSt2.Fields("switchboar") = rst.Fields("sifra") + 2
RSt2.Fields("itemnumber") = 0
RSt2.Fields("itemtext") = rst.Fields("grupa")
RSt2.Fields("command") = 0
RSt2.Update
RSt2.AddNew
RSt2.Fields("switchboar") = rst.Fields("sifra") + 2
RSt2.Fields("itemnumber") = 1
RSt2.Fields("itemtext") = "IZHOD"
RSt2.Fields("command") = 1
RSt2.Fields("argument") = 1
RSt2.Update
Do While Not RSt1.EOF
If RSt1.EOF Then
a = 2
Exit Do
End If

RSt2.AddNew
RSt2.Fields("switchboar") = rst.Fields("sifra") + 2
RSt2.Fields("itemnumber") = a
If Getnazi("select tekst from dokm where tip_dok='XX' and id_dok='CENAPA'") = "D" Then
RSt2.Fields("itemtext") = RTrim(LTrim(RSt1.Fields("madanazi"))) & " - " & RTrim(LTrim(str(RSt1.Fields("madampcd"))))
Else
RSt2.Fields("itemtext") = RTrim(RSt1.Fields("madanazi"))
End If
RSt2.Fields("command") = 8
RSt2.Fields("argument") = RSt1.Fields("madasifr")

RSt2.Fields("dim") = RSt1.Fields("dimm")
RSt2.Update

RSt1.MoveNext
a = a + 1
Loop
a = 2
rst.MoveNext
Loop
End Sub
Public Function Filllist1(Cmbl As ListBox, strSQl As String)
If rs.State = 1 Then rs.Close
rs.Open strSQl, myConection, adOpenKeyset, adLockOptimistic
Dim dolg As String
Dim dolg1 As String
Dim dolg2 As String
Dim dd As Integer
Dim AAS As Integer
If Not rs.EOF Then
    Cmbl.clear
    rs.MoveFirst
    'MsgBox ("")
    Do While Not rs.EOF
    dd = Len(rs.Fields(0))
    AAS = 15 - dd
    
    dolg = ""
    'For i = 0 To AAS
     'dolg = dolg & " "
     'Next i
      dolg2 = Getnazi("select madaenme from mada where madasifr='" & rs.Fields(0) & "'")
        dolg1 = Getnazi("select madanazi from mada where madasifr='" & rs.Fields(0) & "'")
        dolg = str(rs.Fields(0))
        With rs
         
            Cmbl.AddItem presled(dolg, 13) & "  " & presled(dolg1, 25) & "  " & FormatNumber(.Fields(1), 2) & " " & dolg2
            
        End With
    rs.MoveNext
    Loop
End If
End Function
Public Function levi_pres(ByVal vVal As Variant, ByVal iWidth As Integer) As String
    If Len(Trim(vVal)) > iWidth Then
        levi_pres = CStr(vVal)
    Else
    If IsNull(vVal) Then
    Else
        levi_pres = String$(iWidth - Len(Trim(vVal)), " ") & Trim(vVal)
    End If
    End If
End Function


Public Function presled(ByVal vVal As Variant, ByVal iWidth As Integer) As String
    If Len(Trim(vVal)) > iWidth Then
        presled = CStr(vVal)
    Else
    If IsNull(vVal) Then
    Else
        presled = Trim(vVal) & String$(iWidth - Len(Trim(vVal)), " ")
    End If
    End If
End Function
Public Function novast(ByVal vVal As Variant, ByVal iWidth As Integer) As String
    If Len(Trim(vVal)) > iWidth Then
        novast = CStr(vVal)
    Else
        novast = String$(iWidth - Len(Trim(vVal)), "0") & Trim(vVal)
    End If
End Function
Public Function fref(Index As Integer)
    frmControlMain.Wbrow.Visible = False
    frmControlMain.MSHFlexGrid1.Visible = True
    Select Case Index
    
    Case 0
         SQL = "Select * from partner "
        CatalogueName = "Customer"
    Case 1
       ' Sql = "Select * from Supplier WHERE SuppliersID<>'CASH'"
        SQL = "Select * from partner "
        CatalogueName = "Supplier"
    Case 2
     SQL = "Select madasifr,madanazi,madampcd,madazalo from mada"
     CatalogueName = "Category"
    Case 3
        SQL = "Select * from grupa"
        CatalogueName = "Location"
    Case 4
'        SQL = "Select * from PurchaseOrderHeader"
        SQL = "Select stdok,min(datum) as datum,sum(nabcena) as nabcena, min(sifrapart) as sifrapart from nabasif group by stdok"
        CatalogueName = "Purchase Order"
    Case 5
'        SQL = "Select * from PurchaseReturnHeader"
        'Sql = "Select * from TotalReturn"
        'CatalogueName = "Purchase Return"
    Case 6
'        SQL = "Select * from PurchaseRegistryHeader"
       ' Sql = "Select * from TotalPurchase"
        'CatalogueName = "Purchase Registry"
    Case 7
        SQL = "Select st,min(datum) as datum,  sum(znesek) as znesek,min(oseba) as oseba from racusif where [datum] between #" & dod & "# AND #" & ddo & "# group by st order by st"
        CatalogueName = "Sales Return"
    Case 8
       ' Sql = "Select * from TotalSales"
        'CatalogueName = "Sales Registry"
    End Select
    

Call GetNewConnection2
Set Rs1 = New Recordset
If CatalogueName <> "" Then

Set Rs1 = DCON.Execute(SQL)
If Rs1.RecordCount <= 0 Then
    frmControlMain.MSHFlexGrid1.Visible = False
Else
    Set frmControlMain.MSHFlexGrid1.DataSource = Rs1
End If
End If
Set Rs1 = Nothing
Set DCON = Nothing

End Function
Public Function TABLEExist(TableName As String) As Boolean

  On Error GoTo FileDoesNotExist
  'rs.Open "select * from " & TableName, myConection, adOpenDynamic, adLockOptimistic
  myConection.Execute ("select * from " & TableName)
  TABLEExist = True
  Exit Function
  
FileDoesNotExist:
  TABLEExist = False
  
End Function

Public Function nadgradi(xime As String)
If TABLEExist(xime) Then
Else
myConection.Execute ("select znes as sifra,znes as znesek,sifra as dokument,datum into " & xime & "  from trenutna where stdok='XXXX'")
End If
If TABLEExist("zaloga") Then
Else

myConection.Execute ("SELECT sifra,naziv,skl, datum, kol,kol*0 as vrednost,tip_dok, id_dok, cena,cena*0 as prosta,space(10) as veza_td,space(20) as veza_id  into zaloga  from trenutna where stdok='XXXX'")
End If
If TABLEExist("dobavn") Then
Else
On Error GoTo bbb
myConection.Execute ("select kol as st,sifra,kol,znes,cena,naziv,mpc,sifra as stranka ,sifra as faktura,datum into dobavn from trenutna where stdok='XXXX'")
myConection.Execute ("insert into dobavn (st,sifra)values(0,'1')")
bbb:
End If
End Function

Public Function Getnacin(ndok As String) As String
    Dim naz As String
    Dim vreddok As Double
    Dim vreddel As Double
  ' MsgBox SQL
  On Error GoTo addr:
    Set RsNewNo = myConection.Execute("select sum(znes) as znes from nabasif where tip_dok+id_dok='" & ndok & "'")
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    vreddok = 0
    Else
        vreddok = RsNewNo(0).Value
        End If
    Else
        vreddok = 0
    End If
    
    If rs.State = 1 Then rs.Close
    rs.Open ("select * from nacplac where dokument='" & ndok & "'")
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    Getnacin = Getnacin & "       " & presled(Trim(Getnazi("select tekst from dokm where atribut='NACP' and poz=" & rs.Fields("sifra"))), 20) & " : " & levi_pres(FormatNumber(rs.Fields("znesek"), 2), 10) & vbNewLine
    vreddel = vreddel + rs.Fields("znesek")
    rs.MoveNext
    
    Loop
    End If
    
    Getnacin = Getnacin & "       " & presled("GOTOVINA   ", 20) & " : " & levi_pres(FormatNumber(vreddok - vreddel, 2), 10)
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function
Public Function Getdoba(ndok As String) As String
    Dim naz As String
    Dim vreddok As Double
    Dim vreddel As Double
  ' MsgBox SQL
  On Error GoTo addr:
    
  
    
    If rs.State = 1 Then rs.Close
    rs.Open ("select st from dobavn where faktura='" & ndok & "' group by st order by st")
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    Getdoba = Getdoba & Trim(rs.Fields("st")) & ","
    
    rs.MoveNext
    
    Loop
    End If
    
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function

Public Function Getnacin1(ndok As String, zness As Double) As String
    Dim naz As String
    Dim vreddok As Double
    Dim vreddel As Double
  ' MsgBox SQL
  On Error GoTo addr:
    
    vreddok = zness
    
    If rs.State = 1 Then rs.Close
    rs.Open ("select * from nacplac where dokument='" & ndok & "'")
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    Getnacin1 = Getnacin1 & "       " & presled(Trim(Getnazi("select tekst from dokm where atribut='NACP' and poz=" & rs.Fields("sifra"))), 20) & " : " & levi_pres(FormatNumber(rs.Fields("znesek"), 2), 10) & vbNewLine
    vreddel = vreddel + rs.Fields("znesek")
    rs.MoveNext
    
    Loop
    End If
    
    Getnacin1 = Getnacin1 & "       " & presled("GOTOVINA   ", 20) & " : " & levi_pres(FormatNumber(vreddok - vreddel, 2), 10)
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function

Public Function Getnacindan(datum As String, zapso As String) As String
    Dim naz As String
    Dim vreddok As Double
    Dim vreddel As Double
     Dim das, des
das = datum

dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
Dim qas As String
 'MsgBox zapso
  On Error GoTo addr:
  If zapso = "" Then
    qas = "select sum(znes) as znes from nabasif where tip_dok='PA' and datum=#" & dod & "#"
    Else
   qas = "select sum(znes) as znes from nabasif where tip_dok='PA' and uporabnik='" & zapso & "' and datum=#" & dod & "#"
    End If
    '
    Set RsNewNo = myConection.Execute(qas)
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    vreddok = 0
    Else
        vreddok = RsNewNo(0).Value
        End If
    Else
        vreddok = 0
    End If
    
    If rs.State = 1 Then rs.Close
    If zapso = "" Then
    qas = "select sifra,sum(znesek) as znesek from nacplac where datum=#" & dod & "# group by sifra"
   
    Else
     qas = "select sifra,sum(znesek) as znesek from nacplac where datum=#" & dod & "# and uporabnik='" & zapso & "' group by sifra"
    End If
    'MsgBox (qas)
    rs.Open qas, myConection, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    Getnacindan = Getnacindan & " " & presled(Trim(Getnazi("select tekst from dokm where atribut='NACP' and poz=" & rs.Fields("sifra"))), 20) & " : " & levi_pres(FormatNumber(rs.Fields("znesek"), 2), 10) & vbNewLine
    vreddel = vreddel + rs.Fields("znesek")
    rs.MoveNext
    
    Loop
    End If
    
    Getnacindan = Getnacindan & " " & presled("GOTOVINA   ", 20) & " : " & levi_pres(FormatNumber(vreddok - vreddel, 2), 10)
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function
Public Function Getnacindancig(datum As String, zapso As String) As String
    Dim naz As String
    Dim vreddok As Double
    Dim vreddel As Double
     Dim das, des
     Getnacindancig = ""
das = datum
Dim grupp As String
If rs.State = 1 Then rs.Close
rs.Open "select sifra from grupa where vr=2", myConection, adOpenDynamic, adLockOptimistic
grupp = ""
If Not rs.EOF Then
rs.MoveFirst
Do While Not rs.EOF
If grupp = "" Then
grupp = Trim(str(rs.Fields("sifra")))
Else
grupp = grupp & "," & Trim(str(rs.Fields("sifra")))
End If
rs.MoveNext
Loop
End If

myConection.Execute ("update mada set storkol=1 where madagrup in (" & grupp & ")")
dod = Mid(das, 4, 2) & "/" & Left(das, 2) & "/" & Mid(das, 7, 4)
Dim qas As String
 'MsgBox zapso
  
  On Error GoTo addr:
  '(select madasifr from mada where madagrup in (select sifra from grupa where vr=2)")
  If rs.State = 1 Then rs.Close
    qas = "select a.sifra,sum(a.kol) as koli from nabasif a,mada b where a.tip_dok='PA' and a.uporabnik='" & zapso & "' and a.datum=#" & dod & "# and b.storkol=1 and a.sifra=b.madasifr group by a.sifra"
    '
    
    rs.Open qas, myConection, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
    Getnacindancig = Getnacindancig & " " & presled(Trim(Getnazi("select madanazi from mada where madasifr='" & rs.Fields("sifra") & "'")), 15) & " : " & levi_pres(FormatNumber(rs.Fields("koli"), 2), 10) & vbNewLine
    
    rs.MoveNext
    
    Loop
    End If
    
    
    RsNewNo.Close
addr:
    Set RsNewNo = Nothing
End Function

Public Function AllFiles(ByVal FullPath As String) As String

'************************************************

Dim oFs As New FileSystemObject
Dim sAns() As String
Dim oFolder As Folder
Dim oFile As File
Dim lElement As Long

ReDim sAns(0) As String
If oFs.FolderExists(FullPath) Then
Set oFolder = oFs.GetFolder(FullPath)
AllFiles = ""
For Each oFile In oFolder.Files
'MsgBox (oFile.Name)
lElement = IIf(sAns(0) = "", 0, lElement + 1)
ReDim Preserve sAns(lElement) As String
AllFiles = oFile.Name
Next
End If

 
'MsgBox AllFiles

ErrHandler:
Set oFs = Nothing
Set oFolder = Nothing
Set oFile = Nothing
End Function
Public Function JUSTFileName(WithPath As String)

Dim sWithoutPath As String
Dim iLen As Integer
Dim iWhere As Integer

sWithoutPath = WithPath
Do Until InStr(sWithoutPath, "\") = 0
iLen = Len(sWithoutPath)
iWhere = InStr(sWithoutPath, "\")
sWithoutPath = Right(sWithoutPath, iLen - iWhere)
Loop
JUSTFileName = sWithoutPath

End Function
Public Function ocisti(besed As String) As String
   ocisti = Replace(besed, "Å¾", "ž")
   ocisti = Replace(ocisti, "Å¡", "š")
   ocisti = Replace(ocisti, "Ä", "è")
   ocisti = Replace(ocisti, "Ä‘", "ð")
   
   ocisti = Replace(ocisti, "Ä‡", "æ")
   ocisti = Replace(ocisti, "Å½", "Ž")
   ocisti = Replace(ocisti, "Ä†", "Æ")
   
   ocisti = Replace(ocisti, "ÄŒ", "È")
   ocisti = Replace(ocisti, "Ä", "Ð")
   ocisti = Replace(ocisti, "Å ", "Š")
   ocisti = Replace(ocisti, "â‚¬ ", "€")
   
   
End Function
Public Function savekirablg() As String
Dim stbl As String
stbl = InputBox("Vnesi številko blagajne", "Vnos številke blagajne", "1")
savekirablg = stbl
SaveSetting "bll", "sifrablg", "blg", stbl
            
End Function
Public Function stblagg() As String
 stblagg = GetSetting("bll", "sifrablg", "blg", "")
            If stblagg = "" Then
            stblagg = savekirablg()
            End If
End Function
Public Function velifont() As Integer
 velifont = GetSetting("bll", "velfont", "blg", 0)
            
   
End Function

Public Function saveodmiz() As String
Dim stmid As String
stmid = InputBox("Vnesi od katere mize naprej prikazuje", "Vnos številke mize", "1")
saveodmiz = stmid
SaveSetting "bll", "sifrablg", "odmize", stmid
            
End Function
Public Function getodmiz() As String
 getodmiz = GetSetting("bll", "sifrablg", "odmize", "")
            If getodmiz = "" Then
            getodmiz = saveodmiz()
            End If
End Function

