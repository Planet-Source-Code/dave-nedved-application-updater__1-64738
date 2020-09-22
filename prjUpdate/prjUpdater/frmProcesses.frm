VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcesses 
   Caption         =   "Processes"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSComctlLib.ListView lvProcesses 
      Height          =   2940
      Left            =   480
      TabIndex        =   0
      Top             =   2.45745e5
      Visible         =   0   'False
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   5186
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmProcesses.frx":000C
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
    End Type
Dim killPROCESSNAME As Long


Public Sub RefreshProcesses()
    Rem // Just a sub to call the function to refresh the process list
    GetWTSProcesses
End Sub

Private Function GetStringFromLP(ByVal StrPtr As Long) As String
    Dim b As Byte
    Dim tempStr As String
    Dim bufferStr As String
    Dim Done As Boolean

    Done = False
    Do
        Rem // Get the byte/character that StrPtr is pointing to.
        CopyMemory b, ByVal StrPtr, 1
        
        Rem // If you've found a null character, then you're done.
        If b = 0 Then
            Done = True
        Else
            Rem // Get the character for the byte's value
            tempStr = Chr$(b)
            
            Rem // Add it to the string
            bufferStr = bufferStr & tempStr
                
            Rem // Increment the pointer to next byte/char
            StrPtr = StrPtr + 1
        End If
    Loop Until Done
    GetStringFromLP = bufferStr
    End Function

Private Sub Form_Load()
Dim i As Integer
    
    Rem // Display all the Processes running on the system in the list box
    
    lvProcesses.View = lvwReport

    lvProcesses.ColumnHeaders.Add 1, "SessionID", "Session ID"
    lvProcesses.ColumnHeaders.Add 2, "ProcessID", "Process ID"
    lvProcesses.ColumnHeaders.Add 3, "ProcessName", "Process Name"
    lvProcesses.ColumnHeaders.Add 4, "UserID", "User ID"
    lvProcesses.ColumnHeaders(4).Width = lvProcesses.Width - (lvProcesses.ColumnHeaders(1).Width * 3) - 300

    GetWTSProcesses
    
    Rem // Find the process running and kill it.
    Rem // The process we are looking for is the process we are updating.
    For i = 1 To Me.lvProcesses.ListItems.Count
     If Me.lvProcesses.ListItems(i).SubItems(2) = frmMain.Tag Then
      Me.lvProcesses.ListItems(i).Selected = True
      KillProcessById Me.lvProcesses.SelectedItem.SubItems(1)
      frmMain.lvData.Tag = "killprocess=true"
     End If
    Next i
    
    Unload Me
    End Sub

Private Sub lvProcesses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Rem // When a ColumnHeader object is clicked, the ListView control is
    Rem // sorted by the subitems of that column.
    Rem // Set the SortKey to the Index of the ColumnHeader - 1
    
    lvProcesses.SortKey = ColumnHeader.Index - 1
    
    Rem // Set Sorted to True to sort the list.
    lvProcesses.Sorted = True
    End Sub

Private Sub GetWTSProcesses()
   Dim RetVal As Long
   Dim Count As Long
   Dim i As Integer
   Dim lpBuffer As Long
   Dim p As Long
   Dim udtProcessInfo As WTS_PROCESS_INFO
   Dim itmAdd As ListItem

   lvProcesses.ListItems.Clear
   RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
   Rem // WTSEnumerateProcesses was successful
   If RetVal Then
      p = lpBuffer
        For i = 1 To Count
            Rem // Count is the number of Structures in the buffer
            Rem // WTSEnumerateProcesses returns a pointer, so copy it to a
            Rem // WTS_PROCESS_INO UDT so you can access its members
            CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
            Rem // Add items to the ListView control
            Set itmAdd = lvProcesses.ListItems.Add(i, , CStr(udtProcessInfo.SessionID))
                itmAdd.SubItems(1) = CStr(udtProcessInfo.ProcessID)
                Rem // Since pProcessName contains a pointer, call GetStringFromLP to get the
                Rem // variable length string it points to
                If udtProcessInfo.ProcessID = 0 Then
                    itmAdd.SubItems(2) = "System Idle Process"
                Else
                    itmAdd.SubItems(2) = GetStringFromLP(udtProcessInfo.pProcessName)
                End If
                
                Rem // itmAdd.SubItems(3) = CStr(udtProcessInfo.pUserSid)
                itmAdd.SubItems(3) = GetUserName(udtProcessInfo.pUserSid)

                Rem // Increment to next WTS_PROCESS_INO structure in the buffer
                p = p + LenB(udtProcessInfo)
        Next i

        Set itmAdd = Nothing
        
        Rem // Free your memory buffer
        WTSFreeMemory lpBuffer
    Else
        Rem // Error occurred calling WTSEnumerateProcesses
        Rem // Check Err.LastDllError for error code
        MsgBox "Error occurred calling WTSEnumerateProcesses.  " & "Check the Platform SDK error codes in the MSDN Documentation" & " for more information.", vbCritical, "Error " & Err.LastDllError
    End If
    End Sub

Function GetUserName(sID As Long) As String
Rem // Get the username of the person running the process
    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    GetUserName = Left$(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left$(retname, InStr(retname, vbNullChar) - 1)
End Function

