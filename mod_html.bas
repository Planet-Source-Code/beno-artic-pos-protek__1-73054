Attribute VB_Name = "mod_html"
Option Explicit
Public Function MakeShortReport(sqlstring As String, header As String) As Boolean
'On Error GoTo adder:
Dim tempRs As New ADODB.Recordset
Dim i As Integer
Dim data2 As String
 Call GetNewConnection2
    Set Rs1 = New Recordset
       
    Set Rs1 = DCON.Execute(sqlstring)
    
        frmView.Wbrow.Navigate "about:blank"
        
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
   frmView.Wbrow.Document.Write ("<body oncontextmenu='return false;'>")

While Rs1.EOF <> True
    frmView.Wbrow.Document.Write ("<font face=arial>" & header & Rs1.Collect(0) & "<table width=100%>")
    For i = 0 To Rs1.Fields.Count - 1
     
            If IsNull(Rs1.Collect(i)) = True Then
                data2 = " "
            Else
                data2 = Rs1.Collect(i)
            End If
           
        frmView.Wbrow.Document.writeln ("<TR><Td bgcolor=#cccccc><font size=2>" & Rs1.Fields(i).Name & "</td><td bgcolor=#CBC7B6> <font size=2>" & data2 & "</td></Tr>")
    Next i

Rs1.MoveNext
    frmView.Wbrow.Document.Write ("</font></table><BR><BR>")
Wend

   Set Rs1 = Nothing
    Set DCON = Nothing
 
  
    frmView.Show
    
   
'Exit Function
'adder:
'    MakeShortReport = False
End Function

Public Function CreateH_Page1(sqlstring As String, header As String) As Boolean
'On Error GoTo adder:
Dim tempRs As New ADODB.Recordset
Dim i As Integer
Dim data2 As String
 Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim LPath As String
Dim File As String
Dim SQL As String
    
LPath = "g:\tv\sp\"
File = "dspr.dbf"

conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & LPath & ";Extended Properties=dBASE IV;"
          
rs.CursorLocation = adUseClient
rs.CursorType = adOpenKeyset

SQL = "select * from dspr"
rs.Open sqlstring, conn
    'Set Rs1 = New Recordset
       
    'Set Rs1 = DCON.Execute(sqlstring)
  If rs.EOF Then
  MsgBox ("Ni podatkov, Vnesi pravo serijsko!!")
  Else
        frmView.Wbrow.Navigate "about:blank"
        
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
   frmView.Wbrow.Document.Write ("<body oncontextmenu='return false;'>")

While rs.EOF <> True
    frmView.Wbrow.Document.Write ("<font face=arial>" & header & rs.Collect(0) & "<table width=100%>")
    For i = 0 To rs.Fields.Count - 1
     
            If IsNull(rs.Collect(i)) = True Then
                data2 = " "
            Else
                data2 = rs.Collect(i)
            End If
           
        frmView.Wbrow.Document.writeln ("<TR><Td bgcolor=#cccccc><font size=2>" & rs.Fields(i).Name & "</td><td bgcolor=#CBC7B6> <font size=2>" & data2 & "</td></Tr>")
    Next i

rs.MoveNext
    frmView.Wbrow.Document.Write ("</font></table><BR><BR>")
Wend

   Set rs = Nothing
    Set conn = Nothing
 
  
    frmView.Show vbModal
   End If
   
'Exit Function
'adder:
'    MakeShortReport = False
End Function



Public Sub CreateH_Page(strSqry As String, titiles As String)
Dim fld As ADODB.Field
Dim i As Integer
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Dim RS2 As ADODB.Recordset
Set RS2 = myConection.Execute("select glava1,glava2,glava3,glava4,glava5 from oblikar")
   'MsgBox (strSqry)
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    frmView.Wbrow.Navigate2 "about:blank"
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        
    With frmView.Wbrow.Document
    
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
           .Write ("<td>" & RS2.Fields("glava1") & "</td>")
        .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava2") & "</td>")
           .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava3") & "</td>")
           .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava4") & "</td>")
          .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava5") & "</td>")
          .Write ("</tr>")
         .Write ("<td>" & titiles & "</td>")
          .Write ("</tr>")
            .Write ("<tr>")
           .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
        .Write ("<tr>")
        '.Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        For Each fld In Rs1.Fields
            .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
          '  .Write ("<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>")
        Next fld
        'First row
        Dim ii As Double
         Dim iix As Double
        Dim too As Integer
        too = 0
        
         Dim iii As Double
             .Write ("<tr>")
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
           ' If Left(titiles, 1) = "X" Then
             
             If fld.Name = "MADAGRUP" Then
             If too <> fld.Value Then
             .Write ("</td>")
             .Write ("<td>" & Getnazi("select grupa from grupa where sifra=" & fld.Value) & "</td>")
           
             .Write ("</td>")
           .Write ("<tr>")
            .Write ("</tr>")
            .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
            .Write ("<tr>")
           .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
        .Write ("<tr>")
        '.Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
            .Write ("<tr>")
             too = fld.Value
             End If
             End If
            'End If
            If i Mod 2 <> 0 Then
                If fld.Name = "znesek" Or fld.Name = "madampcd" Or fld.Name = "nabv" Then
                   .Write ("<Align=Right VAlign=Middle>")
                .Write ("<td>" & Format(fld.Value, "standard") & "</td>")
                
                Else
            
                .Write ("<td>" & fld.Value & "</td>")
                End If
                If fld.Name = "znesek" Then
                 
                ii = ii + fld.Value
                End If
                 If fld.Name = "nabv" Then
                 If Not IsNull(fld.Value) Then
                  iix = iix + fld.Value
                  End If
                End If
            Else
            If fld.Name = "znesek" Or fld.Name = "madampcd" Or fld.Name = "nabv" Then
             .Write ("<Align=Right VAlign=Middle>")
                .Write ("<td bgcolor=#CCCCC2>" & Format(fld.Value, "standard") & "</td>")
                Else
            
              .Write ("<td bgcolor=#CCCCC2>" & fld.Value & "</td>")
                End If
                
                If fld.Name = "znesek" Then
                ii = ii + fld.Value
                End If
                If fld.Name = "nabv" Then
                 If Not IsNull(fld.Value) Then
                iix = iix + fld.Value
                End If
                End If
                 
            End If
    Next fld
            .Write ("</tr>")
            Rs1.MoveNext
        Wend
        .Write ("</tr>")
         .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
         .Write ("<tr>")
                .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
                 .Write ("<tr>")
                   .Write ("</tr>")
                     .Write ("<tr>")
                     .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
                     .Write ("<tr>")
                     If ii <> 0 Then
                     
                    .Write ("<tr>" & "SKUPAJ VREDNOST Z DDV:                 " & Format(ii, "standard") & "</tr>")
                    .Write ("<tr>" & "SKUPAJ VREDNOST BREZ DDV:                 " & Format(ii / 1.2, "standard") & "</tr>")
                    End If
                    .Write ("<tr>")
                     If iix <> 0 Then
                     
                    .Write ("<tr>" & "SKUPAJ NAB. VREDNOST :                 " & Format(iix, "standard") & "</tr>")
                    .Write ("<tr>" & "RVC :                 " & Format((ii / 1.2) - iix, "standard") & "</tr>")
                    
                    End If
                     If iii <> 0 Then
                     
                    .Write ("<tr>" & "SKUPAJ PO PRODAJNI CENI:                 " & Format(iii, "standard") & "</tr>")
                    End If
                     
                     .Write ("<tr>")
        
            .Write ("</td></tr></table></BODY></HTML>")

        frmView.Wbrow.Document.Script.Document.clear
        frmView.Wbrow.Document.Script.Document.Close
End With


Set DCON = Nothing
frmView.Show vbModal

End Sub
Public Sub CreateH_PageZAL(strSqry As String, titiles As String)
Dim fld As ADODB.Field
Dim i As Integer
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Dim RS2 As ADODB.Recordset
Set RS2 = myConection.Execute("select glava1,glava2,glava3,glava4,glava5 from oblikar")
   'MsgBox (strSqry)
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    Open App.path & "\tempx.htm" For Output As #1

'Print #1, Chr(27) & Chr(112) & Chr(0) & Chr(50) & Chr(100)


    
        Print #1, "<HTML><head></head><style type='text/css'> body,td{font-family: Arial;} body,td{font-size:11px;}</style>" 'Style
        Print #1, "<BODY Scroll=Yes oncontextmenu='return false';>" '
        Print #1, "<table Width=100% border=1><tr>"
        Print #1, "</tr></table>"
        Print #1, "<table Width=100% border=0><tr>"
           Print #1, "<td>" & RS2.Fields("glava1") & "</td>"
        Print #1, "</tr>"
           Print #1, "<td>" & RS2.Fields("glava2") & "</td>"
           Print #1, "</tr>"
           Print #1, "<td>" & RS2.Fields("glava3") & "</td>"
           Print #1, "</tr>"
           Print #1, "<td>" & RS2.Fields("glava4") & "</td>"
          Print #1, "</tr>"
           Print #1, "<td>" & RS2.Fields("glava5") & "</td>"
          Print #1, "</tr>"
         Print #1, "<td>" & titiles & "</td>"
          Print #1, "</tr>"
            Print #1, "<tr>"
           Print #1, "<td bgcolor=#B4C0DC Height=10>===========================================================</td>"
        Print #1, "<tr>"
        'Print #1, titiles)
        Print #1, "<table Width=100% border=1><tr>"
        Print #1, "</tr></table>"
        Print #1, "<table Width=100% border=0><tr>"
        ''Headings
        For Each fld In Rs1.Fields
            Print #1, "<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>"
          '  Print #1, "<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>"
        Next fld
        'First row
        Dim ii As Double
         Dim iix As Double
        Dim too As Integer
        too = 0
        
         Dim iii As Double
             Print #1, "<tr>"
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
           ' If Left(titiles, 1) = "X" Then
             
             'End If
            If i Mod 2 <> 0 Then
                If fld.Name = "vrednost" Or fld.Name = "kol" Then
                   Print #1, "<Align=Right VAlign=Middle>"
                Print #1, "<td>" & Format(fld.Value, "standard") & "</td>"
                
                Else
            
                Print #1, "<td>" & fld.Value & "</td>"
                End If
                If fld.Name = "vrednost" Then
                 
                ii = ii + fld.Value
                End If
                 If fld.Name = "kol" Then
                 If Not IsNull(fld.Value) Then
                  iix = iix + fld.Value
                  End If
                End If
            Else
            If fld.Name = "vrednost" Or fld.Name = "kol" Then
             Print #1, "<Align=Right VAlign=Middle>"
                Print #1, "<td bgcolor=#CCCCC2>" & Format(fld.Value, "standard") & "</td>"
                Else
            
              Print #1, "<td bgcolor=#CCCCC2>" & fld.Value & "</td>"
                End If
                
                If fld.Name = "vrednost" Then
                ii = ii + fld.Value
                End If
                If fld.Name = "kol" Then
                 If Not IsNull(fld.Value) Then
                iix = iix + fld.Value
                End If
                End If
                 
            End If
    Next fld
            Print #1, "</tr>"
            Rs1.MoveNext
        Wend
        Print #1, "</tr>"
         Print #1, "<table Width=100% border=1><tr>"
        Print #1, "</tr></table>"
        Print #1, "<table Width=100% border=0><tr>"
         Print #1, "<tr>"
                Print #1, "<td bgcolor=#B4C0DC Height=10>===========================================================</td>"
                 Print #1, "<tr>"
                   Print #1, "</tr>"
                     Print #1, "<tr>"
                     Print #1, "<table Width=100% border=1><tr>"
        Print #1, "</tr></table>"
        Print #1, "<table Width=100% border=0><tr>"
                     Print #1, "<tr>"
                     If ii <> 0 Then
                     
                    Print #1, "<tr>" & "SKUPAJ VREDNOST      :                 " & Format(ii, "standard") & "</tr>"
              '      Print #1, "<tr>" & "SKUPAJ VREDNOST BREZ DDV:                 " & Format(ii / 1.2, "standard" & "</tr>"
                    End If
                    Print #1, "<tr>"
                     If iix <> 0 Then
                     
                    Print #1, "<tr>" & "SKUPAJ KOLIÈINA :                 " & Format(iix, "standard") & "</tr>"
                   ' Print #1, "<tr>" & "RVC :                 " & Format((ii / 1.2) - iix, "standard" & "</tr>"
                    
                    End If
                     If iii <> 0 Then
                     
                    'Print #1, "<tr>" & "SKUPAJ PO PRODAJNI CENI:                 " & Format(iii, "standard" & "</tr>"
                    End If
                     
                     Print #1, "<tr>"
        
            Print #1, "</td></tr></table></BODY></HTML>"
       'frmPrint.WebBrowser1.Document.Script.Document.clear
       ' frmPrint.WebBrowser1.Document.Script.Document.Close
       


Close #1
Set DCON = Nothing
frmPrint.Show vbModal

End Sub

Public Sub Cre_Page(strSqry As String, titiles As String, sumo As Double, ker As String)
Dim fld, flde As ADODB.Field
Dim i As Integer
Dim vex As String
Dim data2 As Variant
Call GetNewConnection2
    Set Rs1 = New ADODB.Recordset
    Dim RS2 As ADODB.Recordset
Set RS2 = myConection.Execute("select glava1,glava2,glava3,glava4,glava5 from oblikar")
    Set Rs1 = DCON.Execute(strSqry)
    WebSQL = strSqry
    frmView.Wbrow.Navigate2 "about:blank"
        Do While frmView.Wbrow.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        
    With frmView.Wbrow.Document
    
        .Write ("<HTML><head></head><style type='text/css'> body,td{font-family: Courier New;} body,td{font-size:11px;}</style>") 'Style
        .Write ("<BODY Scroll=Yes oncontextmenu='return false';>") '
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
           .Write ("<td>" & RS2.Fields("glava1") & "</td>")
        .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava2") & "</td>")
           .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava3") & "</td>")
           .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava4") & "</td>")
          .Write ("</tr>")
           .Write ("<td>" & RS2.Fields("glava5") & "</td>")
          .Write ("</tr>")
         .Write ("<td>" & titiles & "</td>")
          .Write ("</tr>")
            .Write ("<tr>")
           .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
        .Write ("<tr>")
        '.Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        ''Headings
        'For Each fld In Rs1.Fields
        '    .Write ("<td bgcolor=#B4C0DC Height=10>" & fld.Name & "</td>")
          '  .Write ("<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>")
        'Next fld
        'First row
        Dim ii As Long
        Dim Stan, vred, xvred As Long
        Dim too As String
        too = ""
        Stan = 0
        vred = 0
        xvred = 0
         Dim iii As Long
             .Write ("<tr>")
        While Rs1.EOF <> True
        i = i + 1
            For Each fld In Rs1.Fields
           ' If Left(titiles, 1) = "X" Then
             
             If fld.Name = "sifra" Then
             If too <> fld.Value Then
             .Write ("</td>")
            '  If stan <> 0 Then
            '  .Write ("<tr>")
            '  .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
            '  .Write ("</tr>")
            '  .Write ("<tr>")
            '        .Write ("<tr>" & "STANJE :                           " & Format(stan, "standard") & "</tr>")
            '.Write ("</tr>")
            '        stan = 0
            '  End If
            .Write ("<table Width=100% border=0><tr>")
            If ker = "0" Then
            vred = sumo
            End If
        If vred <> 0 Then
             .Write ("<td>" & "Vrednost zaloge :                  " & FormatNumber(vred, 3) & "</td>")
             Else
             .Write ("<td>" & "Vrednost zaloge :                  " & FormatNumber(0, 3) & "</td>")
             End If
        .Write ("</tr></table>")
             .Write ("<td></td>")
             .Write ("<td></td>")
            .Write ("<table Width=100% border=1><tr>")
        
             .Write ("<td>" & fld.Value & "  " & Getnazi("select madanazi from mada where madasifr='" & fld.Value & "'") & "</td>")
        .Write ("</tr></table>")
             Stan = 0
             vred = 0
             .Write ("</td>")
           .Write ("<tr>")
            .Write ("</tr>")
            .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
        If UCase(fld.Name) = "TD" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "SKKOL" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "DATUM" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "SIFRA" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "VREDNOSTZ" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "STANJE" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "VEZA_TD" Then
        vex = " WIDTH='5%' "
        End If
        If UCase(fld.Name) = "VEZA_ID" Then
        vex = " WIDTH='5%' "
        End If
         If UCase(fld.Name) = "ID_DOKUMENTA" Then
        vex = " WIDTH='5%' "
        End If
         If UCase(fld.Name) = "PROSTA" Then
        vex = " WIDTH='5%' "
        End If
        For Each flde In Rs1.Fields
        
         
            .Write ("<td" & vex & "bgcolor=#B4C0DC Height=10 Align=right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & flde.Name & "<BR></FONT></B></TD>")
          '  .Write ("<TD WIDTH='100%' Align=Center VAlign=Middle><HR Size=0 NoShade></TD>")
        Next flde
        '.Write ("<table Width=100% border=0><tr>")
        '    .Write ("<tr>")
        '   .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
        .Write ("<tr>")
        '.Write (titiles)
        .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
            .Write ("<tr>")
             too = fld.Value
             End If
             End If
            'End If
          '  If i Mod 2 <> 0 Then
                If fld.Name = "skkol" Or fld.Name = "madampcd" Or fld.Name = "vrednostz" Then
                   .Write ("<Align=Right VAlign=Middle>")
                   '<TD WIDTH='10%' Align=Left VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & rs.Fields("madasifr") & "<BR></FONT></B></TD>
                .Write ("<td" & vex & " Align=right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(fld.Value, 3) & "<BR></FONT></B></td>")
                
                Else
           '
             If fld.Name = "stanje" Then
                   .Write ("<Align=Right VAlign=Middle>")
                .Write ("<td" & vex & " Align=right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & FormatNumber(Stan, 3) & "<BR></FONT></B></td>")
               Else
                
                .Write ("<td" & vex & " Align=right VAlign=Bottom><B><FONT SIZE=1 FACE='Helvetica'>" & fld.Value & "<BR></FONT></B></td>")
            End If
                End If
                If fld.Name = "skkol" Then
                 
                Stan = Stan + fld.Value
                'MsgBox strSqry
                vred = vred + Rs1.Fields("prosta") * Rs1.Fields("madampcd")
              '  xvred = xvred + Rs1.Fields("prosta") * Rs1.Fields("madampcd")
                
                 End If
                
           ' Else
           ' If fld.Name = "skkol" Or fld.Name = "madampcd" Then
           '  .Write ("<Align=Right VAlign=Middle>")
           '     .Write ("<td>" & Format(fld.Value, "standard") & "</td>")
           '     Else
           '
           '   .Write ("<td>" & fld.Value & "</td>")
           '     End If
                
           '     If fld.Name = "skkol" Then
           '     stan = stan + fld.Value
           '     End If
                 
           ' End If
            
    Next fld
            .Write ("</tr>")
            ' xvred = xvred + Rs1.Fields("stanjez")
            Rs1.MoveNext
        Wend
        .Write ("</tr>")
         .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
         .Write ("<tr>")
                .Write ("<td bgcolor=#B4C0DC Height=10>===========================================================</td>")
                 .Write ("<tr>")
                   .Write ("</tr>")
                     .Write ("<tr>")
                     .Write ("<table Width=100% border=1><tr>")
        .Write ("</tr></table>")
        .Write ("<table Width=100% border=0><tr>")
                     .Write ("<tr>")
                    If ker = "1" Then
                     'MsgBox (dod)
                     xvred = Getnumb("select Sum(IIf([veza_td]=[tip_dok],kol,prosta)*[cena]) AS aa from zaloga")
                    End If
                    If xvred <> 0 Then
                     
                   '  xvred = Getnazi("select Sum([prosta]*[cena]) AS aa from zaloga")
                    .Write ("<tr>" & "SKUPAJ VREDNOST :                 " & FormatNumber(xvred, 3) & "</tr>")
                    End If
                     If iii <> 0 Then
                     
                    .Write ("<tr>" & "SKUPAJ PO PRODAJNI CENI:                 " & FormatNumber(iii, 3) & "</tr>")
                    End If
                     
                     .Write ("<tr>")
        
            .Write ("</td></tr></table></BODY></HTML>")

        frmView.Wbrow.Document.Script.Document.clear
        frmView.Wbrow.Document.Script.Document.Close
End With


Set DCON = Nothing
frmView.Show vbModal

End Sub


