VERSION 5.00
Begin VB.Form frmAddUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Update"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   Icon            =   "frmAddUpdate.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjUpdateCreator.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtUpdateProcessName 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "sample update.exe"
      Top             =   2280
      Width           =   4575
   End
   Begin VB.TextBox txtUpdateAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "http://updates.datosoftware.com/updatecreator/sampleupdate.exe.txt"
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtUpdateLocation 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "\sample update.exe"
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox txtUpdateDescription 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "This is a Sample Update, Used for Testing Purposes Only"
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox txtUpdateName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Sample Update"
      Top             =   840
      Width           =   4575
   End
   Begin prjUpdateCreator.ucHorizontal3DLine hozLine 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   53
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Process Name:"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Address:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Location:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Description:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Update Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Update"
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
      TabIndex        =   7
      Top             =   240
      Width           =   4815
   End
   Begin VB.Image imgUpdate 
      Height          =   480
      Left            =   5640
      Picture         =   "frmAddUpdate.frx":6852
      Top             =   120
      Width           =   480
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
Attribute VB_Name = "frmAddUpdate"
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

Private Sub cmdCancel_Click()
Rem // Unload this form (ME)
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim X
Rem // When the usere clicks the ok button, to add the file... write the data to the frmMain

     Set X = frmMain.lvUpdates.ListItems.Add(, , Me.txtUpdateName)
         X.SubItems(1) = Me.txtUpdateDescription
         X.SubItems(2) = Me.txtUpdateLocation
         X.SubItems(3) = Me.txtUpdateProcessName
         X.SubItems(4) = Me.txtUpdateAddress
         Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Rem // When the form loads, select all the text in the textboxes... this will make
Rem // The text selected in the text box, so the user can type over the
Rem // Example text, without backspacing it.
Me.txtUpdateAddress.SelStart = 0: Me.txtUpdateAddress.SelLength = 9999
Me.txtUpdateName.SelStart = 0: Me.txtUpdateName.SelLength = 9999
Me.txtUpdateDescription.SelStart = 0: Me.txtUpdateDescription.SelLength = 9999
Me.txtUpdateLocation.SelStart = 0: Me.txtUpdateLocation.SelLength = 9999
Me.txtUpdateProcessName.SelStart = 0: Me.txtUpdateProcessName.SelLength = 9999

DoEvents
Me.txtUpdateName.SetFocus
End Sub

Private Sub txtUpdateAddress_GotFocus()
Rem // Select the text
If Me.txtUpdateAddress.Text = "http://updates.datosoftware.com/updatecreator/sampleupdate.exe.txt" Then Me.txtUpdateAddress.SelStart = 0: Me.txtUpdateAddress.SelLength = 9999
End Sub

Private Sub txtUpdateAddress_KeyPress(KeyAscii As Integer)
Rem // Jump to the next text box
If KeyAscii = 13 Then Me.txtUpdateProcessName.SetFocus
End Sub

Private Sub txtUpdateDescription_GotFocus()
Rem // Select the text
If Me.txtUpdateDescription.Text = "This is a Sample Update, Used for Testing Purposes Only" Then Me.txtUpdateDescription.SelStart = 0: Me.txtUpdateDescription.SelLength = 9999
End Sub

Private Sub txtUpdateDescription_KeyPress(KeyAscii As Integer)
Rem // Jump to the next text box
If KeyAscii = 13 Then Me.txtUpdateLocation.SetFocus
End Sub

Private Sub txtUpdateLocation_GotFocus()
Rem // Select the text
If Me.txtUpdateLocation.Text = "\sample update.exe" Then Me.txtUpdateLocation.SelStart = 0: Me.txtUpdateLocation.SelLength = 9999
End Sub

Private Sub txtUpdateLocation_KeyPress(KeyAscii As Integer)
Rem // Jump to the next text box
If KeyAscii = 13 Then Me.txtUpdateAddress.SetFocus
End Sub

Private Sub txtUpdateName_GotFocus()
Rem // Select the text
If Me.txtUpdateName.Text = "Sample Update" Then Me.txtUpdateName.SelStart = 0: Me.txtUpdateName.SelLength = 9999
End Sub

Private Sub txtUpdateName_KeyPress(KeyAscii As Integer)
Rem // Jump to the next text box
If KeyAscii = 13 Then Me.txtUpdateDescription.SetFocus
End Sub

Private Sub txtUpdateProcessName_GotFocus()
Rem // Jump to the next text box
If Me.txtUpdateProcessName.Text = "sample update.exe" Then Me.txtUpdateProcessName.SelStart = 0: Me.txtUpdateProcessName.SelLength = 9999
End Sub

Private Sub txtUpdateProcessName_KeyPress(KeyAscii As Integer)
Rem // Jump to the ok command button.
If KeyAscii = 13 Then Me.cmdOK.SetFocus
End Sub
