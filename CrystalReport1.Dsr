VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrystalReport1 
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   OleObjectBlob   =   "CrystalReport1.dsx":0000
End
Attribute VB_Name = "CrystalReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Report_Initialize()
If Rs1.State = 1 Then Rs1.Close
Rs1.Open SQLREP, myConection, adOpenDynamic, adLockOptimistic
Me.Text5.SetText (Getnazi("select glava1 from oblikar"))
Me.Database.SetDataSource Rs1
'Getnazi("select glava1 from oblikar")
End Sub

