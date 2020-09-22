Attribute VB_Name = "modNumber"
Option Explicit

Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
     
Private Function WordTens(ByVal SNUM As Long) As String
Select Case SNUM
    Case 1
        WordTens = " One"
    Case 2
        WordTens = " Two"
    Case 3
        WordTens = " Three"
    Case 4
        WordTens = " Four"
    Case 5
        WordTens = " Five"
    Case 6
        WordTens = " Six"
    Case 7
        WordTens = " Seven"
    Case 8
        WordTens = " Eight"
    Case 9
        WordTens = " Nine"
    Case 10
        WordTens = " Ten"
    Case 11
        WordTens = " Eleven"
    Case 12
        WordTens = " Twelve"
    Case 13
        WordTens = " Thirteen"
    Case 14
        WordTens = " Fourteen"
    Case 15
        WordTens = " Fifteen"
    Case 16
        WordTens = " Sixteen"
    Case 17
        WordTens = " Seventeen"
    Case 18
        WordTens = " Eighteen"
    Case 19
        WordTens = " Nineteen"
    Case 20
        WordTens = " Twenty"
    Case 30
        WordTens = " Thirty"
    Case 40
        WordTens = " Fourty"
    Case 50
        WordTens = " Fifty"
    Case 60
        WordTens = " Sixty"
    Case 70
        WordTens = " Seventy"
    Case 80
        WordTens = " Eighty"
    Case 90
        WordTens = " Ninty"
End Select
End Function





Public Function NumToWord(ByVal src_num As String) As String
Dim SNUM  As Double
SNUM = Val(src_num)
If SNUM > 999999999999999# Then
    NumToWord = "Error: To much number."
    Exit Function
End If
Dim WHOLE As String
Dim EXTRA As String
Dim WORD  As String
Dim NWHOLE As Double

If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
   WHOLE = Split(Str$(SNUM), ".")(0)
    EXTRA = Split(src_num, ".")(1)
Else
    WHOLE = SNUM
End If

If SNUM < 1 Then WORD = "Zero"

NWHOLE = Val(WHOLE)
'Check for One and Tens
If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
    WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
ElseIf Val(Right(NWHOLE, 2)) > 20 Then
    WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
    WORD = WORD & WordTens(Right(NWHOLE, 1))
End If
'Check for Hundred
If NWHOLE > 99 Then
   If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Hundred" & WORD
End If
'Check for Thousand
If NWHOLE > 999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Thousand" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Million
If NWHOLE > 999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 6))) & " Million" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 6), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 2, 1)) & " Million" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 99 Then
            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Billion
If NWHOLE > 999999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " Billion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " Billion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 99 Then
            If Left(Right(NWHOLE, 12), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 12), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Trillion
If NWHOLE > 999999999999# Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 12))) & " Trillion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 12), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 2, 1)) & " Trillion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 99 Then
            If Left(Right(NWHOLE, 15), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 15), 1)) & " Hundred" & WORD
        End If
    End If
End If
If EXTRA = "" Then
    WORD = WORD & "   and   00/100"
Else
    If Val(EXTRA) < 10 Then EXTRA = "0" & EXTRA
    WORD = WORD & "   and   " & EXTRA & "/100"
End If
NumToWord = WORD

NWHOLE = 0
WORD = ""
EXTRA = ""
WHOLE = ""
End Function
Public Function AUTONUM(ByVal Active_Connection As ADODB.Connection, ByVal TABLE As String, ByVal PKEY As String, ByVal PREFIX As String, Optional ByVal DISPLAY As Object)
Dim AUTORS As New ADODB.Recordset
Dim ID1 As String
TABLE = UCase(TABLE)
PKEY = UCase(PKEY)


Set AUTORS = New Recordset
Set AUTORS = Active_Connection.Execute("Select * from " & TABLE & " WHERE Val(Right(" & PKEY & ",6))='" _
                                        & Val(Format(Date, "mmddyy")) & "'")


If AUTORS.RecordCount >= 1 Then
AUTORS.MoveLast
ID1 = Val(Mid(AUTORS.Fields(PKEY), 6, 4)) + 1 & "-" & Format(Date, "mmddyy")

Else
ID1 = "0"
ID1 = Val(ID) + 1 & "-" & Format(Date, "mmddyy")

End If

If Val(Right(ID1, 6)) <> Val(Format(Date, "mmddyy")) Then
    ID1 = "0"
    ID1 = Val(ID1) + 1 & "-" & Format(Date, "mmddyy")

End If

  If Val(Mid(ID1, 1, Len(PREFIX))) >= 1000 Then
        
     
            ID1 = PREFIX & "-" + ID1
   
    
    ElseIf Val(Mid(ID1, 1, Len(PREFIX))) >= 100 Then
   
            ID1 = PREFIX & "-0" & ID1
     
    ElseIf Val(Mid(ID1, 1, Len(PREFIX))) >= 10 Then
      
            ID1 = PREFIX & "-00" & ID1
      
   
   ElseIf Val(Mid(ID1, 1, Len(PREFIX))) >= 1 Then
      
            ID1 = PREFIX & "-000" & ID1
       
        
    End If
     
     If Not DISPLAY Is Nothing Then
        DISPLAY = ID1
     End If
     
     
    Set AUTORS = Nothing
    Set Active_Connection = Nothing
    

End Function


Public Function Decimals(Key_ASCII As Integer, ByVal ControlName As Object, ByVal DecimalPlace As Integer)
On Error GoTo DECERR

Static DecPlace As Integer

If InStr(1, ControlName, ".") Then
    If Key_ASCII <> 13 And Key_ASCII <> 8 Then
    If Key_ASCII < 48 Or Key_ASCII > 57 Then Key_ASCII = 0

    
    End If
    If DecPlace = 0 Then
    DecPlace = Val(Len(ControlName) + DecimalPlace)
    ControlName.MaxLength = DecPlace
   
    End If
Else
    DecPlace = 0
    If Key_ASCII <> 13 And Key_ASCII <> 8 And Key_ASCII <> 46 Then
    If Key_ASCII < 48 Or Key_ASCII > 57 Then Key_ASCII = 0
    End If
End If


Exit Function

DECERR:
    MsgBox Err.Description & vbTab, vbInformation
    
End Function
Public Function CMB1(ByVal TABLE As String, ByVal Field As String, CMB As ComboBox)

CMB.Clear
Call DBCONNECT


Set RS1 = New Recordset
RS1.Open "Select * from " + TABLE, DCON, adOpenDynamic, adLockPessimistic

If RS1.RecordCount > 0 Then
    While Not RS1.EOF
       
        CMB.AddItem RS1.Collect(1)
       
        RS1.MoveNext
    Wend

    
    
End If

Set RS1 = Nothing
Set DCON = Nothing


End Function
