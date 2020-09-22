VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Safe Guard Update"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin prjUpdate.DL DL2 
      Left            =   2280
      Top             =   4080
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin prjUpdate.ucHorizontal3DLine ucHorizontal3DLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   4250
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   53
   End
   Begin VB.PictureBox picUpdate 
      BackColor       =   &H00FFFFFF&
      Height          =   3210
      Left            =   120
      ScaleHeight     =   3150
      ScaleWidth      =   5835
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer tmrCopyFileandResume 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   2160
      End
      Begin VB.Timer tmrNextUpdate 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   1680
      End
      Begin prjUpdate.ucFirefoxWait ucFfUpdate 
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         BackColor       =   16777215
      End
      Begin prjUpdate.XP_ProgressBar pbUpdate 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   6956042
         Scrolling       =   4
         Value           =   -1
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Checking For Updates..."
         Height          =   735
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   1680
         Width           =   4935
      End
   End
   Begin VB.PictureBox picFinished 
      BackColor       =   &H80000005&
      Height          =   3210
      Left            =   120
      ScaleHeight     =   3150
      ScaleWidth      =   5835
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label lblUpdateMessage 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":6852
         Height          =   2535
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Updates Complete..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   5655
      End
   End
   Begin prjUpdate.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   53
   End
   Begin VB.PictureBox picLoading 
      BackColor       =   &H00FFFFFF&
      Height          =   3210
      Left            =   120
      ScaleHeight     =   3150
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   960
      Width           =   5895
      Begin VB.Timer tmrDownloadUpdates 
         Interval        =   1000
         Left            =   1800
         Top             =   2280
      End
      Begin VB.Timer tmrSearchPBAR 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1320
         Top             =   2280
      End
      Begin prjUpdate.XP_ProgressBar pbLoad 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   6956042
         Scrolling       =   2
      End
      Begin prjUpdate.ucFirefoxWait ucFfWait 
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         BackColor       =   16777215
      End
      Begin prjUpdate.DL DL1 
         Left            =   720
         Top             =   2280
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Checking For Updates..."
         Height          =   615
         Index           =   1
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   4935
      End
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Update"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8643
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Path"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "File Name"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblUpdateStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   400
      Width           =   5175
   End
   Begin VB.Image imgUpdate 
      Height          =   720
      Left            =   5450
      Picture         =   "frmMain.frx":6938
      Top             =   60
      Width           =   600
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Safe Guard Updates..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3255
   End
   Begin VB.Shape shpBACK 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
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
 Dim DownloadVal As Integer
 Dim DownloadCount As Integer
 Dim DownloadCurrent As Integer
 Dim DownloadFileName As String
 Dim DownloadingFile As Boolean
 Dim DownloadSelectedCount As Integer
 Dim TempStatus As Integer
 Dim ProcessName As String
 Dim UpdateLocation As String
 Dim MTime As Long
 Const XPBlue_ProgressBar = &H2BD228
 Dim killPROCESSNAME As String

Private Sub cmdCancel_Click()
Dim MsgRes

Rem // If the user is half way through updating a file then it may corrupt the update, warn the user
If Me.picUpdate.Visible = True Then
 MsgRes = MsgBox("Are you sure you want to cancel the update process?" & vbNewLine & "If you exit, this may cause data corruption with the files you are updating!", vbExclamation + vbYesNo + vbSystemModal + vbDefaultButton2, "Abort Update?")
 If MsgRes = vbYes Then End
Else
 Rem // If the user isnt updating, then just ask if the user want's to exit
 MsgRes = MsgBox("Are you sure you want to cancel the update process?" & vbNewLine & "You can download the updates at a later time.", vbQuestion + vbYesNo + vbSystemModal + vbDefaultButton2, "Abort Update?")
 If MsgRes = vbYes Then End
End If
End Sub

Private Sub cmdNext_Click()
Dim i
Rem // If the Download value = 1 then download the checked items
If DownloadVal = "1" Then
 Rem // Setup the download...
 DownloadVal = 2
 Me.ucFfUpdate.PlayWait
 Me.lvData.Visible = False
 Me.picLoading.Visible = False
 Me.picUpdate.Visible = True
 Me.cmdNext.Enabled = False
 DownloadSelectedCount = 0
 TempStatus = 0
 Rem // For each item that is checked, count + 1. this will give us our download total.
 For i = 1 To Me.lvData.ListItems.Count
  If Me.lvData.ListItems(i).Checked = True Then
   DownloadCount = DownloadCount + 1
  End If
 Next i
 Rem // More setting up of download.
 Me.lblUpdateStatus.Caption = "Downloading Selected Updates..."
 
 Rem // Call the DownloadUpdates sub
 DownloadUpdates
 Exit Sub
End If

Rem // If the user has updated, then when they click the button, exit the update.
If DownloadVal = "3" Then
 End
End If
End Sub


Sub DownloadUpdates()
Dim i
 For i = 1 To Me.lvData.ListItems.Count
  Rem // If the user isn't updating a file then go for it, otherwise opt out, so we dont start downloading another file
  If DownloadingFile = False Then
   If Me.lvData.ListItems(i).Checked = True Then
    Rem // Setup The Download.
    DownloadCurrent = i
    DoEvents
    Rem // Set the Download File Name.
    DownloadFileName = App.Path & "\" & Me.lvData.ListItems(i).SubItems(2)
    Rem // Download The file, (File address (http://www... etc), Download To Location (in this case a temp .data file)
    Me.DL1.Download Me.lvData.ListItems(i).Tag, App.Path & "\" & Me.lvData.ListItems(i).SubItems(2) & ".1231287318637146.data"
    killPROCESSNAME = Me.lvData.ListItems(i).SubItems(3)
    Rem // Update the UpdateLocation Sting, That is the location of the update.
    UpdateLocation = Me.lvData.ListItems(i).SubItems(2)
    DownloadingFile = True
    Me.lvData.ListItems(i).Checked = False
    DoEvents
    'MsgBox Me.lvData.ListItems(i).Tag
   End If
  End If
 Next i
 
 Rem // This will set +1 to the downloadselectedcount
 Rem // This is for the.. "Downloading x of x files"
 For i = 1 To Me.lvData.ListItems.Count
  If Me.lvData.ListItems(i).Checked = True Then
   DownloadSelectedCount = DownloadSelectedCount + 1: Exit Sub
  End If
 Next i

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Rem // Unsubclass the form, and remove the link in the menu.
  Call modExtendMenu.ReleaseHook(Me.hwnd)
  End
End Sub

Private Sub tmrCopyFileandResume_Timer()
On Error GoTo ErrorCode

    Rem // When the download is complete, it will call this timer
    Rem // The reason the code is in the timer is so we have a few ms wait
    Rem // This is so it dosn't start Downloading the next file right away.
    Rem // Why is that? that is so the computer has a chance to copy & update the file.
    
    Rem // Why did i use filecopy? face it, the user is downloading from the internet.
    Rem // The most we are looking @ to update is about 100meg, so i just used the filecopy feature.
    Rem // Updates are normally under 1mb anyway.
    FileCopy App.Path & UpdateLocation & ".1231287318637146.data", App.Path & UpdateLocation
    
    Rem // Kill the temp data file.
    Kill App.Path & UpdateLocation & ".1231287318637146.data"
    
    Rem // If the Process is running, the kill it.
    Rem // e.g. i am updating the antivirus main mod exe. so it is always running
    Rem // Kill the exe update the file, and load it again
    Rem // the most the user will be unprotected for is about half a seccond.
    If Me.lvData.Tag = "killprocess=true" Then Shell App.Path & UpdateLocation, vbMinimizedFocus
    Me.SetFocus
    
    Rem // Resume the Download Process
    DownloadingFile = False
    TempStatus = 0
    Me.tmrNextUpdate.Enabled = True
    Me.tmrCopyFileandResume.Enabled = False
    Exit Sub


ErrorCode:
Rem // If there is an error, 99% of the time the process is running
Rem // Just store the Temp file, and tell the user it could not be updated
Rem // Then continue the update.
MsgBox "The Dato Software Update Utility Could not Update the selected file because it is in use." & vbNewLine & "Please Close the following File." & UpdateLocation & vbNewLine & vbNewLine & "The updated version has been saved to: " & vbNewLine & App.Path & UpdateLocation & ".1231287318637146.data" & vbNewLine & vbNewLine & "Please manually update this file.", vbExclamation + vbSystemModal, "Dato Software"
DownloadingFile = False
TempStatus = 0
Me.tmrNextUpdate.Enabled = True
Me.tmrCopyFileandResume.Enabled = False
End Sub

Private Sub tmrDownloadUpdates_Timer()
 Rem // This is what starts the application off.
 Rem // I stuck this here, to insert loading code..
 Rem // There isn't much to load so it dosn't matter anyway..
 DownloadVal = "1"
 Me.DL1.Download "http://updates.datosoftware.com/updatecreator/updates.txt", App.Path & "\~1231287318637146.tmp"
 Me.tmrDownloadUpdates.Enabled = False
End Sub

Private Sub tmrNextUpdate_Timer()
Dim i

Rem // This is in a timer to give the last download a half a sec break befour it starts again.
Rem // That way it has enough time to clear its data and start again
For i = 1 To Me.lvData.ListItems.Count
  If Me.lvData.ListItems(i).Checked = True Then
   TempStatus = TempStatus + 1
  End If
 Next i
  If TempStatus = 0 Then
   DownloadVal = "3"
   DL1_Completed 0, 0
  End If

 Rem // Continue Downloadin The Updates... (That is if there is more than 1 update selected)
 DoEvents
 DownloadUpdates
 Me.tmrNextUpdate.Enabled = False
End Sub

Private Sub DL1_Completed(Bytes As Long, sID As String)
Dim TempNAME, TempDESCRIPTION, TempLOCATION, TempADDRESS, LC, TempLCcount, TempFILENAME, X
Dim TempVal As String, TempUpdateVal


Rem // This is where a lot of the project is.
Rem // This will download the update list from the online server.
Rem // Then update the project, or whatever the user selects.
On Error Resume Next

If DownloadVal = "1" Then

   Rem // If the download val is "1" then read the update file, then kill it
   Rem // The update file will contain all the data like links to the updates
   
   Rem // Read the Updates from the temp file, and store the data on temp strings
   TempLCcount = ReadINI("List", "Count", App.Path & "\~1231287318637146.tmp")
   TempUpdateVal = ReadINI("UpdateDate", "date", App.Path & "\settings.cfg")
   TempVal = ReadINI("UpdateDate", "date", App.Path & "\~1231287318637146.tmp")
   
   If TempUpdateVal = TempVal Then TempLCcount = 0
   
   Rem // If the list count = "0" then Send the user to the finished screen & tell them that there are no new updates
   If TempLCcount = "0" Then
    Me.lblUpdateMessage.Caption = "Dato Software Could not Find Any Updates." & vbNewLine & "Safe Guard Appears to be fully updated." & vbNewLine & vbNewLine & "Please check back for new updates later."
    Me.cmdNext.Enabled = True
    Me.cmdCancel.Enabled = False
    Me.cmdNext.Caption = "Finish!"
    DownloadVal = "3"
    Me.picUpdate.Visible = False
    Me.picLoading.Visible = False
    Me.picFinished.Visible = True
    Me.lblUpdateStatus.Caption = "No new updates found."
    Me.lblInfo(3).Caption = "No new updates."
    Kill App.Path & "\~1231287318637146.tmp"
    Exit Sub
   End If
   
   Rem // If there is no listcount then it means that the file on the webserver has been deleted, or modified, or the user has internet problems e.g. not connected to the internet.
   If TempLCcount = "" Then
    Me.lblUpdateMessage.Caption = "Dato Software Could not Contact the Update Server." & vbNewLine & "The Update Server may be down for maintenance, or the update address specified may be invalid." & vbNewLine & vbNewLine & "It may also be possible that you are not connected to the internet, or your internet connection is refusing Update Requests. Please contact your system Administrator." & vbNewLine & vbNewLine & vbNewLine & "Please check back for new updates later."
    Me.cmdNext.Enabled = True
    Me.cmdCancel.Enabled = False
    Me.cmdNext.Caption = "Finish!"
    DownloadVal = "3"
    Me.picUpdate.Visible = False
    Me.picLoading.Visible = False
    Me.picFinished.Visible = True
    Me.lblUpdateStatus.Caption = "Update server not found."
    Me.lblInfo(3).Caption = "No updates."
    Kill App.Path & "\~1231287318637146.tmp"
    Exit Sub
   End If
    
   Rem // Read the update file into the list view
   Rem // Add the checkboxes etc..
   For LC = 1 To TempLCcount
     TempNAME = ReadINI("UpdateNAME", "C:" & LC, App.Path & "\~1231287318637146.tmp")
     TempDESCRIPTION = ReadINI("UpdateDESCRIPTION", "C:" & LC, App.Path & "\~1231287318637146.tmp")
     TempLOCATION = ReadINI("UpdateLOCATION", "C:" & LC, App.Path & "\~1231287318637146.tmp")
     TempADDRESS = ReadINI("UpdateADDRESS", "C:" & LC, App.Path & "\~1231287318637146.tmp")
     TempFILENAME = ReadINI("ProcessName", "C:" & LC, App.Path & "\~1231287318637146.tmp")
   
     Rem // Set the sub items.. e.g. info about the update, update path, update location etc etc..
     Set X = Me.lvData.ListItems.Add(LC, , TempNAME)
         X.SubItems(1) = TempDESCRIPTION
         X.SubItems(2) = TempLOCATION
         X.SubItems(3) = TempFILENAME
     Me.lvData.ListItems(LC).Tag = TempADDRESS
        
   Next LC
   
   Rem // Set the update message, the message that will be displayed when the user finishes updating.
   Me.lblUpdateMessage.Caption = ReadINI("UpdateMessage", "Message", App.Path & "\~1231287318637146.tmp")
   
   Rem // Set the update date.
   TempVal = ReadINI("UpdateDate", "date", App.Path & "\~1231287318637146.tmp")
   WriteINI "UpdateDate", "date", TempVal, App.Path & "\settings.cfg"
   
   Rem // Set the Update Expire date, that is the date that the user should of checked for updates again by.
   TempVal = ReadINI("UpdateExpireDate", "date", App.Path & "\~1231287318637146.tmp")
   WriteINI "UpdateExpireDate", "date", TempVal, App.Path & "\settings.cfg"
      
   Rem // Kill the update information file
   Kill App.Path & "\~1231287318637146.tmp"
   
   Rem // Stop the update visual  stuff and show the selection box.
   Me.ucFfWait.StopWait
   Me.picLoading.Visible = False
   Me.lvData.Visible = True
   Me.cmdNext.Enabled = True
   
   Me.lblUpdateStatus.Caption = "Select Your Updates."
   
End If


If DownloadVal = "2" Then
        
    Rem // if the user is running the file we want to update then terminate the file.
    Me.Tag = killPROCESSNAME
    Load frmProcesses
    
    Rem // Then continue updating
    Me.tmrCopyFileandResume.Enabled = True
End If

If DownloadVal = "3" Then
 
 Rem // If the user has finished the updates then show the finished screen
 Rem // The below is just trying to make it look neater...
 Rem // like if there is no update complete message then show the default one.
 If Me.lblUpdateMessage.Caption = "" Then Me.lblUpdateMessage.Caption = "The updates you selected were successfully Downloaded and installed." & vbNewLine & "For a complete Log on the Update Process that just completed please check the 'updateslog.rtf' file located in the " & App.Path & " folder."
 If Me.lblUpdateMessage.Caption = " " Then Me.lblUpdateMessage.Caption = "The updates you selected were successfully Downloaded and installed." & vbNewLine & "For a complete Log on the Update Process that just completed please check the 'updateslog.rtf' file located in the " & App.Path & " folder."
 If Me.lblUpdateMessage.Caption = "." Then Me.lblUpdateMessage.Caption = "The updates you selected were successfully Downloaded and installed." & vbNewLine & "For a complete Log on the Update Process that just completed please check the 'updateslog.rtf' file located in the " & App.Path & " folder."
 If Me.lblUpdateMessage.Caption = "  " Then Me.lblUpdateMessage.Caption = "The updates you selected were successfully Downloaded and installed." & vbNewLine & "For a complete Log on the Update Process that just completed please check the 'updateslog.rtf' file located in the " & App.Path & " folder."
 If Me.lblUpdateMessage.Caption = "-" Then Me.lblUpdateMessage.Caption = "The updates you selected were successfully Downloaded and installed." & vbNewLine & "For a complete Log on the Update Process that just completed please check the 'updateslog.rtf' file located in the " & App.Path & " folder."
 
 Me.picUpdate.Visible = False
 Me.picFinished.Visible = True
 Me.cmdCancel.Enabled = False
 Me.cmdNext.Caption = "Finished!"
 Me.lblUpdateStatus.Caption = "Finished Updating Files."
 Me.cmdNext.Enabled = True
End If
End Sub

Private Sub DL1_Progress(DownLoadedBytes As Long, TotalBytes As Long, sID As String)
Rem // When the user is downloading the updates, show the progress.
If DownloadVal = "1" Then
 If DownLoadedBytes > 0 Then
    Rem // if it is downloading the first file... the info data file then just show the bytes completed
    Me.lblInfo(1).Caption = "Downloading Updates..." & vbNewLine & "Downloaded: " & DownLoadedBytes & " of " & TotalBytes & "."
 End If
Else
 If DownLoadedBytes > 0 Then
    Rem // If it is the update screen then show the update progressbare
    Rem // and show they bytes completed.
    Me.pbUpdate.Value = (DownLoadedBytes * 100) / TotalBytes
    Me.lblInfo(2).Caption = "Downloading Updates... " & DownloadSelectedCount & " of " & DownloadCount & " files." & vbNewLine & "Downloaded: " & DownLoadedBytes & " of " & TotalBytes & "."
 End If
End If
End Sub


Private Sub Form_Load()
On Error Resume Next
 
Rem // On the form load process check if the user is running in the vb ide, or in the compiled state
Rem // if the user is in the vb ide then do nothing
Rem // If the user is running the compiled application the subclass the form.
 If App.LogMode = 0 Then
 'Do Nothing
 Else
  If modExtendMenu.CreateMenuEntry("Dato Software Online", Me.hwnd, 5, MenuByPosition) = True Then
   Call modExtendMenu.CreateHook(Me.hwnd)
  End If
 End If
 
 Rem // Set the Progressbar Colors, setup the form make it look nice :D
 Me.pbLoad.Color = XPBlue_ProgressBar
 Me.pbUpdate.Color = XPBlue_ProgressBar
 MTime = 1
 Me.tmrSearchPBAR.Enabled = True
 Me.lblUpdateStatus.Caption = "Downloading Update List..."
 ucFfWait.PlayWait
 Kill App.Path & "\~1231287318637146.tmp"
End Sub

Private Sub tmrSearchPBAR_Timer()
Rem // This is just a gimic thing ... rolf
Rem // This will make the search bar thing go @ the start of the project, when the
Rem // user is checking for updates... etc etc.

MTime = MTime + 1
If MTime > pbLoad.Max Then
    MTime = pbLoad.Min
End If
Me.pbLoad.Value = MTime
End Sub
