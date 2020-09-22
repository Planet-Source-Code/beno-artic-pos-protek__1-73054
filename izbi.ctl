VERSION 5.00
Begin VB.UserControl UserControl1 
   BackStyle       =   0  'Transparent
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   435
   ScaleWidth      =   3870
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   120
   End
   Begin VB.CommandButton cmdClear 
      DisabledPicture =   "izbi.ctx":0000
      Height          =   345
      Left            =   3090
      Picture         =   "izbi.ctx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmdPicker 
      DisabledPicture =   "izbi.ctx":0B14
      Height          =   345
      Left            =   2760
      Picture         =   "izbi.ctx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   345
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim x_ssql As String
Dim mDisplayData As String
Dim x_Top As Integer
Dim x_left1 As Integer
Dim ki As String
Dim x_polje As String
Dim i As String
Dim widd As Integer
Private mBoundDatax As String
Public Event BeforeDropDown()

Public Event Change()



Private Sub cmdClear_Click()
txtDisplay.Text = ""
'MsgBox ((Screen.Width / 2) - (Parent.Width / 2) + Me.Left1)


End Sub
Private Sub cmdPicker_lostfocus()

Me.BoundDatax = txtDisplay.Text


End Sub
Private Sub cmdPicker_Click()


Dim sDT As String
    Dim sBT As String
    Dim OldBT As String
    Dim sqlukaz As String
    
    Load Izbor
    RaiseEvent BeforeDropDown
   ' MsgBox (Screen.TwipsPerPixelX)
    '((Screen.Width / 2) - (Parent.Width / 2) + Me.Left1)
If txtDisplay.Text = "" Then
sqlukaz = Me.sSQL
Else
sqlukaz = Me.sSQL & " where " & Me.polje & " like '" & txtDisplay.Text & "%'"
End If
txtDisplay.Text = (Izbor.odpri((Screen.Width / 2) - (Parent.Width / 2) + Me.Left1, (Screen.Height / 2) - (Parent.Height / 2) + Me.Top1 + (UserControl.Height * 2), sqlukaz, sBT, sDT, Me, UserControl.Parent, Me.polje))


End Sub
Public Property Get BoundDatax() As String


    BoundDatax = mBoundDatax

End Property
Public Property Let BoundDatax(ByVal NewValue As String)
    mBoundDatax = NewValue
End Property

Public Property Get Left1() As Integer
  Left1 = x_left1
End Property

Public Property Let Left1(ByVal New_Left1 As Integer)
  x_left1 = New_Left1
    PropertyChanged "Left1"
End Property
Public Property Get Top1() As Integer
  Top1 = x_Top
End Property

Public Property Let Top1(ByVal New_Top As Integer)
  x_Top = New_Top
    PropertyChanged "Top1"
End Property
Public Property Get sSQL() As String
  sSQL = x_ssql
End Property

Public Property Let sSQL(ByVal New_ssql As String)
  x_ssql = New_ssql
    PropertyChanged "ssql"
End Property
Public Property Get polje() As String
  polje = x_polje
End Property

Public Property Let polje(ByVal New_polje As String)
  x_polje = New_polje
    PropertyChanged "polje"
End Property
Public Sub opentime()
Timer1.Enabled = True

End Sub
Public Sub Timer1_Timer()


If sSQL = "" Then
cmdPicker.Visible = False
cmdClear.Visible = False
Else
cmdPicker.Visible = True
cmdClear.Visible = True

End If
If mBoundDatax <> "" Then
txtDisplay.Text = mBoundDatax
End If

Timer1.Enabled = False
End Sub

Private Sub txtDisplay_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
If cmdPicker.Visible = True Then
Call cmdPicker_Click
Else
Sendkeys "{TAB}"
End If
End If

End Sub

Private Sub txtDisplay_lostfocus()
Me.BoundDatax = txtDisplay.Text
End Sub
Private Sub txtDisplay_Change()
Me.BoundDatax = txtDisplay.Text
End Sub

Private Sub txtDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 txtDisplay.ToolTipText = txtDisplay.Text
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
      x_ssql = PropBag.ReadProperty("ssql", x_def_ssql)
        x_polje = PropBag.ReadProperty("polje", x_def_polje)
         x_Top = PropBag.ReadProperty("Top1", x_def_Top)
         x_left1 = PropBag.ReadProperty("Left1", x_def_Left1)
       txtDisplay.Locked = PropBag.ReadProperty("TextLocked", True)
       UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
           txtDisplay.Locked = PropBag.ReadProperty("Locked", True)
           txtDisplay.Text = PropBag.ReadProperty("tex", x_def_mBoundDatax)
End Sub
'Initialize Properties for User Control
Public Sub UserControl_InitProperties()

      x_ssql = x_def_ssql
       x_Top = x_def_Top
         x_left1 = x_def_Left1
           x_polje = x_def_polje
           mBoundDatax = x_def_mBoundDatax
       RaiseEvent Change
End Sub

Private Sub UserControl_Resize()
    
    If GetWidth < 58 Then
        UserControl.Width = 58 * 15
    End If
    If GetHeight < 21 Then
        UserControl.Height = 21 * 15
    End If
   ' MsgBox (GetHeight)
    txtDisplay.Move 0, 1, UserControl.Width - cmdClear.Width - cmdPicker.Width - 4, UserControl.Height - 1
    cmdPicker.Move UserControl.Width - cmdPicker.Width - cmdClear.Width - 2, 0, cmdPicker.Width, UserControl.Height - 1
    cmdClear.Move UserControl.Width - cmdClear.Width, 0, cmdClear.Width, UserControl.Height - 1
widd = txtDisplay.Width
End Sub
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
    
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)


    Call PropBag.WriteProperty("Locked", txtDisplay.Locked, True)
    Call PropBag.WriteProperty("polje", x_polje, x_def_polje)
     Call PropBag.WriteProperty("ssql", x_ssql, x_def_ssql)
     Call PropBag.WriteProperty("tex", mBoundDatax, x_def_mBoundDatax)
       Call PropBag.WriteProperty("Top1", x_Top, x_def_Top)
        Call PropBag.WriteProperty("Left1", x_left1, x_def_Left1)
     Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
      Call PropBag.WriteProperty("TextLocked", txtDisplay.Locked, True)

End Sub
Public Property Get TextLocked() As Boolean
    TextLocked = txtDisplay.Locked
End Property

Public Property Let TextLocked(ByVal New_TextLocked As Boolean)
    txtDisplay.Locked() = New_TextLocked
    PropertyChanged "TextLocked"
End Property
Public Property Get Locked() As Boolean
    Locked = txtDisplay.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtDisplay.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtDisplay.Enabled = New_Enabled
    cmdPicker.Enabled = New_Enabled
    cmdClear.Enabled = New_Enabled
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get DisplayData() As String
    DisplayData = txtDisplay.Text
End Property
Public Property Let DisplayData(ByVal NewValue As String)
    txtDisplay.Text = NewValue
End Property
