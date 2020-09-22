VERSION 5.00
Begin VB.UserControl DL 
   BackColor       =   &H00FF00FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   Enabled         =   0   'False
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   495
   ScaleWidth      =   495
   ToolboxBitmap   =   "DL.ctx":0000
   Begin VB.CheckBox butBDR 
      Height          =   495
      Left            =   0
      MaskColor       =   &H00FF00FF&
      Picture         =   "DL.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "DL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
 Public Event Progress(DownLoadedBytes As Long, TotalBytes As Long, sID As String)
 Public Event Completed(Bytes As Long, sID As String)
 Private colDest As New Collection

Public Sub Download(sWWWFile As String, sDestination As String, Optional sID As String = "Id")
On Error Resume Next
    colDest.Add sDestination, sID
    UserControl.AsyncRead sWWWFile, vbAsyncTypeFile, sID, vbAsyncReadForceUpdate
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error Resume Next
   Name AsyncProp.Value As colDest.Item(AsyncProp.PropertyName)
   colDest.Remove AsyncProp.PropertyName
   RaiseEvent Completed(AsyncProp.BytesRead, AsyncProp.PropertyName)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent Progress(AsyncProp.BytesRead, AsyncProp.BytesMax, AsyncProp.PropertyName)
End Sub

Public Sub CancelDownload(Optional sID As String = "Id")
    UserControl.CancelAsyncRead sID
End Sub

Private Sub UserControl_Resize()
 UserControl.Height = "495"
 UserControl.Width = "495"
End Sub

Rem // ----------------------------------------------------------------------------------------
Rem // Note to rest of team...
Rem // The below code is commented out beacuse it will catche the file.
Rem // It all works fine, just the file will be catched after downloading it.
Rem // I dont want that to happen cause i want the update to be realtime, not catched
Rem // Cheers, dave
Rem // ----------------------------------------------------------------------------------------

'Event Declarations:
'Event Progress(Max As Long, Min As Long)
'Event Complete(Data As Variant)
'
'
'Public Sub Download(URL As String)
'    On Error Resume Next
'    AsyncRead URL, 1
'End Sub
'
'
'Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
'Dim inData As String
'On Error Resume Next
'Open AsyncProp.Value For Binary As #1
'inData = Space(FileLen(AsyncProp.Value))
'Get #1, , inData
'Close
'RaiseEvent Complete(inData)
'End Sub
'
'Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
'RaiseEvent Progress(AsyncProp.BytesMax, AsyncProp.BytesRead)
'End Sub
'
'
'Private Sub UserControl_Resize()
'UserControl.Height = "495"
'UserControl.Width = "495"
'End Sub

Rem // ----------------------------------------------------------------------------------------
