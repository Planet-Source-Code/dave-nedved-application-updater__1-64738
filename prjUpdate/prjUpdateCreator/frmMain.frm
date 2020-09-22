VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dato Software Update Creator"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLoadUpdate 
      Caption         =   "Load Update"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   5640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdWriteToFile 
      Caption         =   "Write Update To File"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelUpdate 
      Caption         =   "Delete Update"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddUpdate 
      Caption         =   "Add Update"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvUpdates 
      Height          =   2055
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Update"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Update Address"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtUpdateFinished 
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox txtUpdateExpireDate 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtUpdateDate 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin prjUpdateCreator.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   53
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Creator.. Create && Edit Updates..."
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
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image imgUpdate 
      Height          =   480
      Left            =   5640
      Picture         =   "frmMain.frx":6852
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Updates"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Message"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Expire Date"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Date"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape shpBACK 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
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
Rem // |      use this program to help you creaste update files.      |
Rem // ----------------------------------------------------------------

Option Explicit

Private Sub cmdAddUpdate_Click()
Rem // show the add new itm screen
frmAddUpdate.Show vbModal
End Sub

Sub SaveUpdateToFile(sUpdateFile As String)
    Dim lC
    
    Rem // Save the update to file...
    Rem // Write the product name @ the top.
    WriteINI "Dato Software, Safe Guard", "", "", sUpdateFile
    
    Rem // Then write the header info... like how many itms to update, update dates, titles etc...
    WriteINI "List", "Count", Me.lvUpdates.ListItems.Count, sUpdateFile
    WriteINI "UpdateDate", "date", Me.txtUpdateDate, sUpdateFile
    WriteINI "UpdateExpireDate", "date", Me.txtUpdateExpireDate, sUpdateFile
    WriteINI "UpdateMessage", "message", Me.txtUpdateFinished, sUpdateFile

    Rem // For each item in the list, write it & its sub info to file.
    For lC = 1 To Me.lvUpdates.ListItems.Count
     WriteINI "UpdateName", "C:" & lC, Me.lvUpdates.ListItems.Item(lC).Text, sUpdateFile
     WriteINI "UpdateDescription", "C:" & lC, Me.lvUpdates.ListItems.Item(lC).SubItems(1), sUpdateFile
     WriteINI "UpdateLocation", "C:" & lC, Me.lvUpdates.ListItems.Item(lC).SubItems(2), sUpdateFile
     WriteINI "UpdateAddress", "C:" & lC, Me.lvUpdates.ListItems.Item(lC).SubItems(4), sUpdateFile
     WriteINI "ProcessName", "C:" & lC, Me.lvUpdates.ListItems.Item(lC).SubItems(3), sUpdateFile
     WriteINI "RegisterFile", "C:" & lC, "0", sUpdateFile
    Next lC
    
    Rem // Let the user know that the file is written
    MsgBox "Update Written to File." & vbNewLine & sUpdateFile, vbInformation, "Dato Software Update Creator"
End Sub

Sub LoadUpdate(sUpdateFile As String)
    Dim lC, tmpLC, X
    
    Rem // Load the update file...
    Rem // Load the update count.
    tmpLC = ReadINI("List", "Count", sUpdateFile)
    
    Rem // Load all the info into the text boxes.
    Me.txtUpdateDate.Text = ReadINI("UpdateDate", "date", sUpdateFile)
    Me.txtUpdateExpireDate = ReadINI("UpdateExpireDate", "date", sUpdateFile)
    Me.txtUpdateFinished = ReadINI("UpdateMessage", "message", sUpdateFile)

    Rem // For each item in the file, add it to the list.
    For lC = 1 To tmpLC
     Set X = frmMain.lvUpdates.ListItems.Add(, , ReadINI("UpdateName", "C:" & lC, sUpdateFile))
         X.SubItems(1) = ReadINI("UpdateDescription", "C:" & lC, sUpdateFile)
         X.SubItems(2) = ReadINI("UpdateLocation", "C:" & lC, sUpdateFile)
         X.SubItems(4) = ReadINI("UpdateAddress", "C:" & lC, sUpdateFile)
         X.SubItems(3) = ReadINI("ProcessName", "C:" & lC, sUpdateFile)
    Next lC
    
End Sub


Private Sub cmdDelUpdate_Click()
Rem // Delete the selected itm
Me.lvUpdates.ListItems.Remove (Me.lvUpdates.SelectedItem.Index)
End Sub

Private Sub cmdLoadUpdate_Click()
Dim sFile As String

Rem // Load an Update File to the form..
Rem // There is errorcode incase the user clicks the cancel button
Rem // If the user does, this will cause a cancel error, and will exit the sub
On Error GoTo ErrorCode

Rem // With the common dialoug.
With cdFile
 Rem // Set the options needed
 .CancelError = True
 .DialogTitle = "Load Update File"
 .Filter = "Text File (*.txt)|*.txt|Update Database (*.db)|*.db|Information File (*.inf)|*.inf"
 .ShowOpen
 
 Rem //Set the file to load as a temp string.
 sFile = .FileName
End With
 Rem // Load the update file.
 LoadUpdate sFile
 Exit Sub
 
ErrorCode:
End Sub

Private Sub cmdOK_Click()
Rem // Exit the application
End
End Sub

Private Sub cmdWriteToFile_Click()
Dim sFile As String

Rem // Write the update & data to a update file.
Rem // There is errorcode incase the user clicks the cancel button
Rem // If the user does, this will cause a cancel error, and will exit the sub.
On Error GoTo ErrorCode

Rem // With the common dialoug
With cdFile
 Rem // Set the options needed
 .CancelError = True
 .DialogTitle = "Save Update File"
 .Filter = "Text File (*.txt)|*.txt|Update Database (*.db)|*.db|Information File (*.inf)|*.inf"
 .ShowSave
 
 Rem // Set the file thet the user choose to write to as a temp tring.
 sFile = .FileName
End With

 Rem // Save / Write the update file.
 SaveUpdateToFile sFile
 Exit Sub
 
ErrorCode:
End Sub
