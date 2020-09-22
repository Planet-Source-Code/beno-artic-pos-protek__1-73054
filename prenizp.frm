VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSInet.ocx"
Begin VB.Form prenizp 
   Caption         =   "prenizp"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "prenizp"
   ScaleHeight     =   660
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet inetDownload 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "POÈAKAJ TRENUTEK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "prenizp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CoInitialize Lib "ole32" (ByVal pvReserved As Long) As Long


Private Sub Form_Activate()
Dim bytes() As Byte
Dim fnum As Integer

   
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Get the file.
    bytes() = inetDownload.OpenURL( _
        "http://90.157.186.14/izpisi.mdb", icByteArray)

    ' Save the file.
    fnum = FreeFile
    Open App.path & "/database/izpisi.mdb" For Binary Access Write As #fnum
    Put #fnum, , bytes()
    Close #fnum
 Screen.MousePointer = vbDefault
Unload Me


End Sub

Private Sub Form_Unload(Cancel As Integer)
    CoInitialize 0
End Sub


