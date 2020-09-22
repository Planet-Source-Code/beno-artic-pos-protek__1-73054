VERSION 5.00
Begin VB.Form izpra 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "izpra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim tString  As String
  Dim cPrint As clsMultiPgPreview
    
    Set cPrint = New clsMultiPgPreview
    

    
SendToPrinter:
    picPrinting.Visible = True
    
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
    'cPrint.pPrint " Prodajalec: " & Me.Label3.Caption
    If Me.imes.Text <> "" Then
    
    cPrint.pPrint
    cPrint.pPrint "Stranka:"
    cPrint.pPrint Left(Me.imes.Text, 40)
cPrint.pPrint Mid(Me.imes.Text, 40, 40)
cPrint.pPrint Left(Me.nassl.Text, 40)
cPrint.pPrint Mid(Me.nassl.Text, 40, 40)
cPrint.pPrint "ID.ST.: SI" & Me.dav.Text

    
    End If
    
    cPrint.pPrint
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint "Racun St.", 0.1, True
    cPrint.pPrint txtInvoiceNo.Text, 1, True
    cPrint.pPrint " z dne " & Format(Date, "dd/mm/yyyy")
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
    ddv1 = 0
    ddv2 = 0
    popu = 0
    znep = 0
    sku = 0
    For i = 1 To MSHFlexGrid1.Rows - 1
    
   If Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'") = "20" Then
   ddv1 = ddv1 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    If Replace(Getnazi("select madapd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'"), ",", ".") = "8.5" Then
   ddv2 = ddv2 + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
   End If
    stri = Format(MSHFlexGrid1.TextMatrix(i, 2), "standard")
    stri1 = Format(MSHFlexGrid1.TextMatrix(i, 4), "standard")
    If MSHFlexGrid1.TextMatrix(i, 4) <> "" Then
    sku = sku + FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)
    End If
     If stri1 <> "" Then
     If (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2)) > 0 Then
     znep = znep + (Getnumb("select madampcd from mada where madasifr='" & MSHFlexGrid1.TextMatrix(i, 0) & "'") - FormatNumber(MSHFlexGrid1.TextMatrix(i, 4), 2))
     End If
    'MsgBox (Val(Getnazi("select madampcd from mada where madasifr=" & Val(MSHFlexGrid1.TextMatrix(i, 1)))) - (Val(MSHFlexGrid1.TextMatrix(i, 5)) / Val(MSHFlexGrid1.TextMatrix(i, 4))))
    If MSHFlexGrid1.TextMatrix(i, 6) <> 0 Then
    popu = popu + (Getnumb("select madampcd from mada where madasifr='" & (MSHFlexGrid1.TextMatrix(i, 0)) & "'")) - FormatNumber(MSHFlexGrid1.TextMatrix(i, 3), 2)
    End If
    End If
    'popu = 0
    'popu = FormatNumber(popu, 2)
    cPrint.pPrint "", 0.1, False
    cPrint.pPrint MSHFlexGrid1.TextMatrix(i, 1), 0.1, True
    cPrint.pRightJust stri, 2.7 * tiskdol, True
    
    cPrint.pRightJust Format(MSHFlexGrid1.TextMatrix(i, 6), "standard"), 3.3 * tiskdol, True
    cPrint.pRightJust stri1, 4 * tiskdol, True
    Next
   
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
    
    If Me.kart.Value = True Then
    pl = "KARTICA"
    Else
    pl = "GOTOVINA"
    End If
     If Me.inter.Value = True Then
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


If FileExist("c:\be.txt") Then
Call Shell("print /d:" & LTrim(RTrim(Getnazi("select POSPRINT from lokal"))) & " c:\be.txt", 6)
End If
    picPrinting.Visible = False
    cPrint.pEndDoc
      cPrint.SendToPrinter = True
    cPrint.Orientation = Printer.Orientation
    Set cPrint = Nothing
 
End Sub
