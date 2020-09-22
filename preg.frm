VERSION 5.00
Begin VB.Form preg 
   Caption         =   "Pregled"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13965
   LinkTopic       =   "Form7"
   ScaleHeight     =   9855
   ScaleWidth      =   13965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "TISKAJ"
      Height          =   615
      Left            =   12480
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8700
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "preg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PrintListBox List1
End Sub
Public Sub PrintListBox(TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Printer.FontSize = 10
    Printer.FontName = "Courier New"

    For SaveList& = 0 To TheList.ListCount - 1
        Printer.Print TheList.List(SaveList&)
    Next SaveList&
    Printer.EndDoc
End Sub



Private Sub Form_Load()
If izp_fifx = 1 Then


izp_fifx = 0
Else
If RS.State = 1 Then RS.Close
RS.Open "SELECT normati.sifr,mada.dobavit_id,normati.naz, normati.kol,  mada.MADAENME, mada.MADAGRUP FROM normati LEFT JOIN mada ON normati.sifr = mada.MADASIFR order by mada.madagrup", myConection, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then

RS.MoveFirst

End If
Dim i
With List1
If Not RS.EOF Then
RS.MoveFirst
End If
Dim gr, xgr As Integer
gr = RS.Fields("madagrup")
xgr = 0
Do While Not RS.EOF
gr = RS.Fields("madagrup")
If gr <> xgr Then
If xgr <> 0 Then
.AddItem " _________________________________________________________________________________________________________"
End If
.AddItem ""
.AddItem RS.Fields("madagrup") & "  " & Getnazi("select grupa from grupa where sifra=" & RS.Fields("madagrup"))
.AddItem " _________________________________________________________________________________________________________"
End If
.AddItem presled(Trim(RS.Fields(0)), 6) & "  " & presled(Left(RS.Fields(1), 12), 15) & "  " & presled(Left(RS.Fields(2), 40), 40) & " " & levi_pres(FormatNumber(RS.Fields(3), 3), 10) & "   " & presled(Trim(RS.Fields(4)), 3)
If Not IsNull(RS.Fields("madagrup")) Then
xgr = RS.Fields("madagrup")
Else
xgr = 0
End If
RS.MoveNext
Loop
End With
End If
End Sub
