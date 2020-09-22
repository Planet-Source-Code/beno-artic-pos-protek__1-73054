VERSION 5.00
Begin VB.Form narocila 
   BackColor       =   &H008080FF&
   Caption         =   "PREGLED NAROÈIL"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5145
   LinkTopic       =   "Form9"
   ScaleHeight     =   420
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox onode 
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "PREGLEDUJEM ÈE OBSTAJAJO NAROÈILA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "narocila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Const NAME_COLUMN = 0
'Const TYPE_COLUMN = 1
'Const SIZE_COLUMN = 2
'Const DATE_COLUMN = 3
Private sFileName As String
Private xSPozz As String
Private idpa As Integer
    Private kjejee As String
Private Const sglSplitLimit = 500
Private Const M_BufferWidth = 60
Private kolar, cenar As Double
Private xxiddo, xsifra, imes, nasl, post, kra, dav, sifrar, opisar As String
Private mbMoving As Boolean
Private m_oDoc As XMLDocument
Private m_oCurrentElement As CXmlElement
Private Sub FillNode(onode As MSComctlLib.node, oElement As CXmlElement)
    Dim oChild As CXmlElement
    Dim oChNode As MSComctlLib.node
    Dim lIndex As Long
    Dim sKey As String
    
    Me.onode.Text = oElement.Name
    
    For Each oChild In oElement
        sKey = lIndex
        'Me.onode.Text = Me.onode.Text & " " & tvwChild
        'Set oChNode = tvTreeView.Nodes.Add(onode, tvwChild, sKey)
       ' MsgBox (oChild.Name)
        If oChild.Name = "Stranke" Then
        kjejee = "STRANKA"
        End If
       ' MsgBox kjejee
        If kjejee = "STRANKA" Then
        If oChild.Name = "Naziv" Then
        imes = ocisti(oChild.Body)
        End If
        
        If oChild.Name = "Naslov" Then
        nasl = ocisti(oChild.Body)
        End If
        If oChild.Name = "PostnaStevilka" Then
        post = oChild.Body
        End If
        If oChild.Name = "NazivPoste" Then
        kra = ocisti(oChild.Body)
        End If
        If oChild.Name = "DavcnaStevilka" Then
        
        dav = oChild.Body
        
        If Getnazi("select naziv from partner where naziv like '%" & imes & "%'") <> "" Then
        idpa = Getnazi("select sifra from partner where naziv like '%" & imes & "%'")
        Else
        If rs.State = 1 Then rs.Close
        idpa = Getnumb("SELECT MAX(SIFRA) FROM PARTNER") + 1
        rs.Open "insert into partner (sifra,naziv,ulica,mesto,posta,davcna) values (" & idpa & ",'" & imes & "','" & nasl & "','" & kra & "','" & post & "','SI" & dav & "')", myConection, adOpenDynamic, adLockOptimistic
        End If
        
        End If
        
        End If
        If oChild.Name = "Artikli" Then
        kjejee = "ARTIKLI"
        End If
        If oChild.Name = "NarociloVrstice" Then
        kjejee = "POZICIJE"
        End If
        
        If kjejee = "POZICIJE" Then
            If oChild.Name = "SifraArtikla" Then
            xsifra = oChild.Body
            sifrar = Getnazi("SELECT MADASIFR FROM MADA WHERE DOBAVIT_ID like '%" & oChild.Body & "%'")
            
            End If
            If oChild.Name = "Opis" Then
            opisar = ocisti(oChild.Body)
            End If
            If oChild.Name = "Kolicina" Then
            kolar = FormatNumber(oChild.Body, 2)
            End If
            If oChild.Name = "Cena" Then
            cenar = FormatNumber(Replace(oChild.Body, ".", ",")) * (1 + (Getnumb("select madapd from mada where madasifr='" & sifrar & "'") / 100))
            If rs.State = 1 Then rs.Close
            If xSPozz = "" Then
            xSPozz = "   1"
            Else
            xSPozz = levi_pres(Trim(str(Val(xSPozz) + 1)), 4)
            End If
            NAZART = Getnazi("SELECT MADANAZI FROM MADA WHERE MADASIFR='" & sifrar & "'")
            Dim SQLLL As String
            
            SQLLL = "insert into nabasif (tip_dok,id_dok,datum,pozicija,sifra,kol,cena,znes,naziv,faktor)values('NK','" & xxiddo & "','" & Date & "','" & xSPozz & "','" & sifrar & "'," & Replace(kolar, ",", ".") & "," & Replace(cenar, ",", ".") & "," & Replace(cenar * kolar, ",", ".") & ",'" & NAZART & "',0)"
            'MsgBox SQLLL
            If sifrar = "" Then
            MsgBox ("Ne najdem šifre izedelka " & xsifra)
            Else
           
            rs.Open SQLLL, myConection, adOpenDynamic, adLockOptimistic
            If rs.State = 1 Then rs.Close
            rs.Open "update glavna set dod0='" & Getnazi("select naziv from partner where sifra=" & idpa) & "' where tip_dok='NK' and id_dok='" & xxiddo & "'", myConection, adOpenDynamic, adLockOptimistic
            If opisar <> "" Then
            If rs.State = 1 Then rs.Close
            rs.Open "INSERT INTO DOKM (tip_dok,id_dok,ATRIBUT,TEKST) VALUES ('NK','" & xxiddo & "','" & xSPozz & "','" & opisar & "')", myConection, adOpenDynamic, adLockOptimistic
            End If
            Timer1.Enabled = True
            End If
            opisar = ""
            End If
            
        End If
        Call FillNode(oChNode, oChild)
        lIndex = lIndex + 1
    Next
  
      'SQL = "insert into glavna (tip_dok,id_dok,opis,dod0,dod1,dod2,dod3,dod4,dod5,dod6,dod7,skl) values ('" & Left(Me.dok.Caption, 2) & "','" & Mid(Me.dok.Caption, 3) & "','" & Me.Text1.Text & "','" & Me.UserControl11(0).BoundDatax & "','" & Me.UserControl11(1).BoundDatax & "','" & Me.UserControl11(2).BoundDatax & "','" & Me.UserControl11(3).BoundDatax & "','" & Me.UserControl11(4).BoundDatax & "','" & Me.UserControl11(5).BoundDatax & "','" & Me.UserControl11(6).BoundDatax & "','" & Me.UserControl11(7).BoundDatax & "','" & Me.sklad.BoundDatax & "')"
End Sub

Private Sub LoadDoc(oDoc As XMLDocument)
    Dim onode As MSComctlLib.node
    
    Set m_oDoc = oDoc
    
   ' tvTreeView.Nodes.clear
   ' Set oNode = tvTreeView.Nodes.Add(, , ":0")
    Set m_oCurrentElement = m_oDoc.Root
    
    Call FillNode(onode, m_oCurrentElement)
   ' Set tvTreeView.SelectedItem = oNode
End Sub

Private Sub Command1_Click()
  MsgBox imes & nasl & post
End Sub

Private Sub Form_Load()
Dim oDoc As XMLDocument
    
    Set oDoc = OpenFile
    
    If Not oDoc Is Nothing Then Call LoadDoc(oDoc)
   ' FileCopy sFileName, App.path & "\arhivnaro\" & JUSTFileName(sFileName)
    
End Sub
Private Function FileExist(FileName As String) As Boolean

  On Error GoTo FileDoesNotExist
  
  Call FileLen(FileName)
  FileExist = True
  Exit Function
  
FileDoesNotExist:
  FileExist = False
  
End Function

Public Function OpenFile() As XMLDocument
    'Dim oDlg As clsFileDlg
    
    Dim sFileData As String
    Dim hFile As Integer
    hFile = FreeFile
    sFileName = App.path & "\naro\" & AllFiles(App.path & "\naro")
    If AllFiles(App.path & "\naro") = "" Then
    Me.Timer1.Enabled = True
    Exit Function
    End If
    If JUSTFileName(sFileName) = "stock.xml" Then
    Kill App.path & "\naro\" & JUSTFileName(sFileName)
    Me.Timer1.Enabled = True
    Exit Function
    End If
    
    xxiddo = novast(Val(Getnazi("select max(id_dok) from nabasif where tip_dok='NK'")) + 1, 6)
    If rs.State = 1 Then rs.Close
    rs.Open "insert into glavna (tip_dok,id_dok,opis)values('NK','" & xxiddo & "','" & sFileName & "')", myConection, adOpenDynamic, adLockOptimistic
   
    Open sFileName For Input As hFile
    sFileData = Input(LOF(hFile), hFile)
    
    'sFileData = App.path & "\ben.xml"
    
    Set OpenFile = New XMLDocument
    Call OpenFile.LoadData(sFileData)
    Close hFile
End Function


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo bbb:
Dim fso As scripting.FileSystemObject
    Set fso = New scripting.FileSystemObject
 If AllFiles(App.path & "\naro") <> "" Then
 If FileExist(App.path & "\arhivnaro\" & JUSTFileName(sFileName)) Then
 Kill App.path & "\arhivnaro\" & JUSTFileName(sFileName)
 End If
    fso.MoveFile sFileName, App.path & "\arhivnaro\" & JUSTFileName(sFileName)
'Kill sFileName
End If
    'Kill sFileName
   ' Unload Me
bbb:
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
