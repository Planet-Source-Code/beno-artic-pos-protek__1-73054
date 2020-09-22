VERSION 5.00
Begin VB.Form FKeyboard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipkovnica"
   ClientHeight    =   4425
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTyping 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10960
      Begin ProsVent.xcKeyboard xcKeyboard1 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5953
      End
   End
End
Attribute VB_Name = "FKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------
' copyright(c) 2000, 2006 Original Software Designs
'-------------------------------------------------------------------------

Option Explicit

Private Sub mnuDemo_Keyboard_Click()

   
    
End Sub

Private Sub mnuDemo_Keypad_Click()

    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub xcKeyboard1_Keyboard(KeyPressed As String)
    
    Select Case KeyPressed
        Case "BS"
            If Len(txtTyping.Text) > 0 Then
                txtTyping.Text = Left$(txtTyping.Text, Len(txtTyping.Text) - 1)
            End If
        Case "CR"
            'Debug.Print txtTyping.Text
            If idtipk < 5 Then
            pregledr.Text2(idtipk).Text = txtTyping.Text
            Else
            dobavnice.Text2(idtipk).Text = txtTyping.Text
            End If
            txtTyping.Text = ""
            Unload Me
        Case Else
            txtTyping.Text = txtTyping.Text & KeyPressed
    End Select
    
End Sub

