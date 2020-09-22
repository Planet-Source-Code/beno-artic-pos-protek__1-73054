VERSION 5.00
Begin VB.MDIForm pla 
   BackColor       =   &H8000000C&
   Caption         =   "Poslovanje"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   -60
   ClientWidth     =   14700
   LinkTopic       =   "placa"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FF8080&
      FillColor       =   &H00E0E0E0&
      Height          =   8625
      Left            =   0
      ScaleHeight     =   8565
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox i32x32 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   6240
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox SmallImages 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   10560
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   4
         Top             =   150
         Width           =   1200
      End
   End
   Begin VB.PictureBox iml16 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14670
      TabIndex        =   1
      Top             =   0
      Width           =   14700
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14640
      TabIndex        =   2
      Top             =   8625
      Width           =   14700
   End
   Begin VB.Menu mnuTop 
      Caption         =   "Datoteka"
      Begin VB.Menu mnuFileNew 
         Caption         =   "Nova      "
         Begin VB.Menu mnuNew 
            Caption         =   "Stranke...."
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Artikli"
            Index           =   3
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuNew 
            Caption         =   "Kategorije"
            Index           =   4
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Printaj"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Nastavi printanje"
      End
      Begin VB.Menu mnuPrintPrv 
         Caption         =   "Predogled printanja"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Shrani"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Izhod"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Urejanje"
      Visible         =   0   'False
      Begin VB.Menu mnuModify 
         Caption         =   "Uredi "
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Briši       "
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Podrobnosti      "
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Osveži"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Najdi"
      End
   End
   Begin VB.Menu mnUsers 
      Caption         =   "Uporabniki"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Dodaj uporabnika"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Briši uporabnika"
      End
      Begin VB.Menu mnuChangeUsername 
         Caption         =   "Spremeni username"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Spremeni  Password"
      End
      Begin VB.Menu mnuViewall 
         Caption         =   "Pregled vseh uporabnikov"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Orodja"
      Begin VB.Menu mnuvoz 
         Caption         =   "UVOZ"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
   End
End
Attribute VB_Name = "pla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
