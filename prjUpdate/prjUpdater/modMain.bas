Attribute VB_Name = "modMain"
Rem // ----------------------------------------------------------------
Rem // |                      ...Project Info...                      |
Rem // |                                                              |
Rem // |               Programmed by David Nedved (Neddy)             |
Rem // |                                                              |
Rem // |                  em. dnedved@datosoftware.com                |
Rem // |                    ws. www.datosoftware.com                  |
Rem // |                                                              |
Rem // |            This is a great little update utility             |
Rem // |           Just Use the Update Editor to create an            |
Rem // |        update information file, and you'r set to go!         |
Rem // ----------------------------------------------------------------
Rem // |                       ...Shout Outs...                       |
Rem // |                                                              |
Rem // | Many thanks to Mario Flores for his1 wonderfull progressbar! |
Rem // |                                                              |
Rem // |            Also Thanks to MaRiØ G for his                    |
Rem // |     help with the Downloader Control.. Works Great!          |
Rem // |                                                              |
Rem // |    Thanks to CubeSolver for his great Firefox Control..      |
Rem // |         and thanks to Firefox for inventing it!              |
Rem // |                                                              |
Rem // |          Thankyou SKoW for your NT Process List.             |
Rem // ----------------------------------------------------------------

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long

Public Sub KillProcessById(p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
    End If
End Sub

Private Sub RetrieveError()
  Dim strBuffer As String
    
    strBuffer = Space(200)
    
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
End Sub



Private Sub Pause(iSecs As Integer)
    Dim i As Integer
    
    For i = 1 To iSecs * 10
        Sleep 100
        DoEvents
    Next
End Sub


Public Sub MAIN()
Rem // Start The Application
frmMain.Show
End Sub
