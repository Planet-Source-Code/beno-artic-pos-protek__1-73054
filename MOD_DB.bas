Attribute VB_Name = "MOD_DB"
Option Explicit
Public DCON As New ADODB.Connection
Public RS1 As New ADODB.Recordset
Public RS2 As New ADODB.Recordset
Public RS3 As New ADODB.Recordset
Public SQL As String
Public ID As String
Public EDT As Boolean
Public CRED As Boolean


Public Sub MAIN()

Load frmPOS
frmPOS.Show




End Sub
Public Sub DBCONNECT()
Dim CONSTR As String
Set DCON = New Connection

DCON.CursorLocation = adUseClient

CONSTR = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" + App.Path + "\Database\INVENT2000V.MDB"

DCON.Open CONSTR


End Sub
