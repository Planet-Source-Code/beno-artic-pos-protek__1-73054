Attribute VB_Name = "tesx"
Option Explicit
Public Pblagajna As String
Public tiskdol As Double
Public vrjenniz, prijavljen, id_inv As String
Public pritisk, idpo As String
Public UREJAJ, UR_id, zapore, izp_fifx, prvaa, dell As Integer
Public repor As String
Public UPORABNIK As String
Public sxqll As String
Public imedn, sqt, xizb As String
Public intCtrlDown, kjje As Integer
Public dtip_dok As String
Public ma_ko, xpox, xzago As Integer
Public FR As String
Public xEM, xskladd As String
Public tresi As String
Public kater As Integer
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

Public dodajwh As String
Public avtomob As String
Public relacij, idfx, idko As String
Public coollna, coollce, coollko, coolldat_k, coollur, coollzn, coollzal, coollsi, coollem, coollpro, coolldat, collchk, cooznes, coollx, coolly, COOZALO, coollpop, coollles, coollstek, cooskkol As Integer
Public zai As Long
Public zaix As Long
Public ssqq As String
Public izja As Integer
Public kosovni As Integer
Public ma_ured As Integer
Public trenu As Integer
Public visina As Long
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
Public skumi As Double
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

Private Sub Main()

Dim ids As String
Dim idda As Long
ids = InputBox("raèun št:", "Raèun", "")
idda = InputBox("Davèna št:", "Davèna", "")
PathFileName = App.Path + "\DATABASE\Thesis.mdb"
myConection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password="
izpr ids, idda
End Sub
Public Function izpr(strac As String, davc As Long)
Dim tString  As String
Dim davcnas As Long
Dim tiskdol As Double
tiskdol = Getnumb("select termi from lokal")
  Dim cPrint As clsMultiPgPreview
  davcnas = davc
    Set cPrint = New clsMultiPgPreview
    

    
SendToPrinter:
    'picPrinting.Visible = True
    
    cPrint.pStartDoc
    cPrint.FontSize = 8
    cPrint.FontName = "Courier new"
    cPrint.CurrentY = 0
    cPrint.pPrint Getnazi("select glava1 from oblikar")
    cPrint.pPrint Getnazi("select glava2 from oblikar")
    cPrint.pPrint Getnazi("select glava3 from oblikar")
    cPrint.pPrint Getnazi("select glava4 from oblikar")
    cPrint.pPrint Getnazi("select glava5 from oblikar")
    
    cPrint.pPrint
    davcnas = Getnumb("select org from nabasif where tip_dok='PA' and id_dok='" & strac & "'")
    'cPrint.pPrint " Prodajalec: " & Me.Label3.Caption
    If davcnas <> 0 Then
    Dim idstr, nall As String
    idstr = Getnazi("select ime from po where dav='" & davcnas & "'")
    nall = Getnazi("select nasl from po where dav='" & davcnas & "'")
    If idstr = "" Then
    nall = Getnazi("select nasl from fozD where dav='" & davcnas & "'")
    idstr = Getnazi("select ime from fozD where dav='" & davcnas & "'")
    End If

    cPrint.pPrint
    cPrint.pPrint "Stranka:"
    cPrint.pPrint Left(idstr, 40)
cPrint.pPrint Mid(idstr, 40, 40)
cPrint.pPrint Left(nall, 40)
cPrint.pPrint Mid(nall, 40, 40)
cPrint.pPrint "ID.ST.: SI" & str(davcnas)

    
    End If
    
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0.1, True
    cPrint.pPrint strac, 1, True
    cPrint.pPrint " z dne " & Format(Getnazi("select datum from nabasif where tip_dok='PA' and id_dok='" & strac & "'"), "dd/mm/yyyy")
    '& " " & Format(Time(), "hh:mm")
    
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    cPrint.pPrint "Naziv                  kol  pop  znesek", 0.1, False
    cPrint.pPrint "=======================================", 0.1, False
    Dim i, ass
    Dim popu As Double
    Dim sku As Double
    Dim stri, stri1
    Dim ddv1 As Double
    Dim znep As Double
    Dim ddv2 As Double
    Dim xplacc As Integer
    ddv1 = 0
    ddv2 = 0
    popu = 0
    znep = 0
    sku = 0
    Dim rst As New ADODB.Recordset
    rst.Open "select * from nabasif where tip_dok='PA' and id_dok='" & strac & "'", myConection, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
    End
    End If
    rst.MoveFirst
    Do While Not rst.EOF
    xplacc = rst.Fields("placilo")
        
   If Getnazi("select madapd from mada where madasifr='" & (rst.Fields("sifra")) & "'") = "20" Then
   ddv1 = ddv1 + FormatNumber(rst.Fields("znes"), 2)
   End If
    If Replace(Getnazi("select madapd from mada where madasifr='" & (rst.Fields("sifra")) & "'"), ",", ".") = "8.5" Then
   ddv2 = ddv2 + FormatNumber(rst.Fields("znes"), 2)
   End If
    stri = Format(rst.Fields("kol"), "standard")
    stri1 = Format(rst.Fields("znes"), "standard")
    If rst.Fields("znes") <> "" Then
    sku = sku + FormatNumber(rst.Fields("znes"), 2)
    End If
     If stri1 <> "" Then
     If (Getnumb("select madampcd from mada where madasifr='" & rst.Fields("sifra") & "'") - FormatNumber(rst.Fields("znes"), 2)) > 0 Then
     znep = znep + (Getnumb("select madampcd from mada where madasifr='" & rst.Fields("sifra") & "'") - FormatNumber(rst.Fields("znes"), 2))
     End If
    'MsgBox (Val(Getnazi("select madampcd from mada where madasifr=" & Val(MSHFlexGrid1.TextMatrix(i, 1)))) - (Val(MSHFlexGrid1.TextMatrix(i, 5)) / Val(rst.fields("znes"))))
    If rst.Fields("pop") <> 0 Then
    popu = popu + (Getnumb("select madampcd from mada where madasifr='" & (rst.Fields("sifra")) & "'")) - (FormatNumber(rst.Fields("znes") / rst.Fields("kol"), 2))
    End If
    End If
    'popu = 0
    'popu = FormatNumber(popu, 2)
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint Left(rst.Fields("naziv"), 18), 0.1, True
    cPrint.pRightJust stri, 2.7 * tiskdol, True
    
    cPrint.pRightJust rst.Fields("pop"), 3.1 * tiskdol, True
    cPrint.pRightJust stri1, 4 * tiskdol, True
    rst.MoveNext
    Loop
   
    cPrint.pPrint ""
    'cPrint.pPrint ""
    cPrint.pPrint "=======================================", 0.1, False
    'cPrint.pPrint ""
    If popu <> 0 Then
    cPrint.pPrint "Popust vracunan v ceni", 0.1, True
    cPrint.pRightJust Format(popu, "standard"), 4, True
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    End If
    cPrint.pPrint "ZA PLACILO EUR ", 0.1, True
    cPrint.pRightJust Format(sku, "standard"), 2.5, True
    cPrint.pPrint "", 0.1, False
    zavrnit = sku
    
      cPrint.pPrint
    
      If ddv1 <> 0 Or ddv2 <> 0 Then
    cPrint.pPrint "---------------------------------------", 0.1, False
    cPrint.pPrint "Osnova DDV-a   DDV Znesek DDV  Vrednost", 0.1, False
    cPrint.pPrint "---------------------------------------", 0.1, False
    If ddv1 <> 0 Then
    'cPrint.pPrint
    cPrint.pRightJust Format(ddv1 / 1.2, "standard"), 0.7, True
    cPrint.pRightJust "20 %", 2.3 * tiskdol, True
    cPrint.pRightJust Format(ddv1 - (ddv1 / 1.2), "standard"), 3.3 * tiskdol, True
    cPrint.pRightJust Format(ddv1, "standard"), 4 * tiskdol, True
    End If
     If ddv2 <> 0 Then
    cPrint.pPrint
    cPrint.pRightJust Format(ddv2 / 1.085, "standard"), 0.7, True
    cPrint.pRightJust "8.5 %", 2.3, True
    cPrint.pRightJust Format(ddv2 - (ddv2 / 1.085), "standard"), 3.3, True
    cPrint.pRightJust Format(ddv2, "standard"), 4, True
    
    End If
    End If
    Dim pl As String
    
    If xplacc = 9999 Then
    pl = "KARTICA"
    Else
    pl = "GOTOVINA"
    End If
     If xplacc = 1 Then
    pl = "INTERNA     Podpis ______________________"
    Else
    pl = "GOTOVINA"
    End If
    'cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint " Placilo: " & plax, 0.1, False
    If Getnazi("select konec1 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec1 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec2 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec2 from oblikar"), 0.1, False
    End If
    If Getnazi("select konec3 from oblikar") <> "" Then
    cPrint.pPrint Getnazi("select konec3 from oblikar"), 0.1, False
    End If
    cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
   cPrint.pPrint "", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
    cPrint.pPrint " ", 0.1, False
        cPrint.pPrint "", 0.1, False
   cPrint.pPrint " ", 0.1, False
 cPrint.pPrint ""


'If FileExist("c:\be.txt") Then
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
'End If
   ' picPrinting.Visible = False
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
 

End Function
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
Public Function Getnazi(SQL As String) As String
    Dim naz As String
  ' MsgBox SQL
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
  ' MsgBox SQL
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
Public Function Getcena(yident As String) As Double
    Dim naz As Double
  ' MsgBox SQL
  Getcena = 0
  On Error GoTo evb:
    Set RsNewNo = myConection.Execute("select cena from cenik where sifra='" & yident & "' order by datum desc")
    If Not RsNewNo.EOF Then
    If IsNull(RsNewNo(0).Value) Then
    naz = 0
    Else
        naz = RsNewNo(0).Value
        End If
    Else
        naz = 0
    End If
    Getcena = naz
    RsNewNo.Close
    Set RsNewNo = Nothing
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
    Err.clear
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
    cen = Getnazi("select madampcd from mada where madasifr='" & rs.Fields(0) & "'")
    zalo = Getnazi("select madazalo from mada where madasifr='" & rs.Fields(0) & "'")
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





