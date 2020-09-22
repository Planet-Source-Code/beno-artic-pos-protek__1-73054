Attribute VB_Name = "Mod_Interfacing"
Option Explicit
'.////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Avoid Public Statment As for Decalration As far as Possible
Public FRMVAL As Form
Public Sub KeyEvent(X As Integer, Optional ByRef CNTROL As Object)
If X = 13 Then
   CNTROL.SetFocus
End If
End Sub
Public Sub clear(frm As Form)
Dim ctr As Integer
    For ctr = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(ctr) Is TextBox Or TypeOf frm.Controls(ctr) Is ComboBox Then
                frm.Controls(ctr).text = ""
        End If
    Next
End Sub
Public Sub highLight()
   With Screen.ActiveForm
      If (TypeOf .ActiveControl Is TextBox) Then
         .ActiveControl.SelStart = 0
         .ActiveControl.SelLength = (Len(.ActiveControl))
         End If
   End With
   Exit Sub
End Sub
Public Function moveShape(shape As Object, text As Object)
        shape.Visible = True
        shape.Move text.Left + 10, text.Top + 10, text.Width + 10, text.Height + 10
End Function


Public Sub CONTLOCK(ByVal frm As Form)
Dim CONTROL1 As Object
For Each CONTROL1 In frm
    If TypeOf CONTROL1 Is TextBox Then
        CONTROL1.text = ""
        CONTROL1.Locked = True
    ElseIf TypeOf CONTROL1 Is ComboBox Then
        CONTROL1.ListIndex = -1
        CONTROL1.Locked = True
    ElseIf TypeOf CONTROL1 Is ListBox Then
        CONTROL1.Enabled = False
    End If
Next
End Sub
Public Sub CONTUNLOCK(ByVal frm As Form)
Dim CONTROL1 As Object
For Each CONTROL1 In frm
    If TypeOf CONTROL1 Is TextBox Then
        CONTROL1.text = ""
        CONTROL1.Locked = False
    ElseIf TypeOf CONTROL1 Is ComboBox Then
        CONTROL1.ListIndex = -1
        CONTROL1.Locked = False
    ElseIf TypeOf CONTROL1 Is ListBox Then
        CONTROL1.Enabled = True
    End If
 Next
End Sub
Public Sub CONTUNLOCK2(ByVal frm As Form)
Dim CONTROL1 As Object
For Each CONTROL1 In frm
    If TypeOf CONTROL1 Is TextBox Then
        CONTROL1.Locked = False
    ElseIf TypeOf CONTROL1 Is ComboBox Then
        CONTROL1.Locked = False
     ElseIf TypeOf CONTROL1 Is ListBox Then
        CONTROL1.Enabled = True
    End If
Next
End Sub
'Example - MakeTopMost Me.hwnd' and so on forth for other arrtribs in CONSTANTS
Public Sub MakeTopMost(Handle As Long)
    SetWindowPos Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
' Example - MakeNormal Me.hwnd
Public Sub MakeNormal(Handle As Long)
SetWindowPos Handle, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

