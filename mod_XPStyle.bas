Attribute VB_Name = "mod_XPStyle"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Public Sub ManifestWrite()

  Dim lngFN        As Long
  Dim mstrEXEName  As String
  
   '/* Standard manifest file
   '/* http://support.microsoft.com/default.aspx?scid=kb;en-us;309366
   Const C_XPLookXML As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
     "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
     "<assemblyIdentity" & vbCrLf & _
     "   version=""1.0.0.0""" & vbCrLf & _
     "   processorArchitecture=""X86""" & vbCrLf & _
     "   name=""$""" & vbCrLf & _
     "   type=""win32""" & vbCrLf & _
     "/>" & vbCrLf & _
     "<description>XP-Look</description>" & vbCrLf & _
     "<dependency>" & vbCrLf & _
     "   <dependentAssembly>" & vbCrLf & _
     "      <assemblyIdentity" & vbCrLf & _
     "         type=""win32""" & vbCrLf & _
     "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
     "         version=""6.0.0.0""" & vbCrLf & _
     "         processorArchitecture=""X86""" & vbCrLf & _
     "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
     "         language=""*""" & vbCrLf & _
     "      />" & vbCrLf & _
     "   </dependentAssembly>" & vbCrLf & _
     "</dependency>" & vbCrLf & _
     "</assembly>"
  
   On Error Resume Next
   
   '/* Create manifest file if it is missing
   mstrEXEName = App.Path & "\" & App.EXEName & ".exe"
   '/* valid directory; create manifest file
   If LenB(Dir$(mstrEXEName & ".Manifest")) = 0 Then
        lngFN = FreeFile
        Open mstrEXEName & ".Manifest" For Output As lngFN
        Print #lngFN, Replace$(C_XPLookXML, "$", Replace$(Mid$(mstrEXEName, InStrRev(mstrEXEName, "\") + 1), ".exe", vbNullString, , , vbTextCompare))
        Close lngFN
   End If
   DoEvents
   
   '/* Link XP themes to application
   Call InitCommonControls

   On Error GoTo 0

End Sub

Public Sub EndApp()
  
  '/* Call from the closing Form's Form_Unload event
  '/* Example:
  '/*   Call EndApp
  '/*   Set FormName = Nothing
  '/*   '/* End Program

  Dim Frm As Form
  Const SEM_NOGPFAULTERRORBOX As Long = &H2
  
   On Error Resume Next
    
   '/* Close all open Forms
   For Each Frm In Forms
       Unload Frm
       Set Frm = Nothing
   Next Frm

   '/* Some versions of ComCtl32.DLL version 6.0 cause a crash at shutdown
   '/* when you enable XP Visual Styles in an application that has VB User Controls.
   '/* This instructs Windows to not display the UAE message box that invites you to send
   '/* Microsoft information about the problem.
   If CBool(VB.App.LogMode()) Then '/* Not running in IDE
      Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
   End If
         
End Sub

