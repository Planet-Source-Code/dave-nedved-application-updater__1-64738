Attribute VB_Name = "modExtendMenu"
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

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function InsertMenu Lib "user32" Alias _
        "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition _
        As Long, ByVal wFlags As Long, ByVal wIDNewItem As _
        Long, ByVal lpNewItem As Any) As Long
        
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const MF_SEPARATOR = &H800&
Private Const MF_BYPOSITION = &H400&
Private Const GWL_WNDPROC = (-4)
Private Const WM_SYSCOMMAND = &H112

Private lngPrevProc As Long
Private intItemID As Integer

Public Enum MenuInsertType
  Separator = 0
  MenuByPosition = 1
End Enum

Public Function CreateMenuEntry(ByVal strName As String, ByVal lngFormHwnd As Long, ByVal intMenuPosition As Integer, _
                                ByVal mnuType As MenuInsertType) As Boolean

Dim lngMnuHandle As Long
Dim lngRetValue As Long
Dim intFlag As Integer

  On Error GoTo errHandler
  lngMnuHandle = GetSystemMenu(lngFormHwnd, False)
  Select Case mnuType
    Case MenuInsertType.MenuByPosition: intFlag = MF_BYPOSITION
    Case MenuInsertType.Separator: intFlag = MF_SEPARATOR
  End Select
  
  If lngMnuHandle Then
    lngRetValue = InsertMenu(lngMnuHandle, intMenuPosition, intFlag, intItemID, strName)
  End If
  
  CreateMenuEntry = True
  Exit Function
  
errHandler:
  CreateMenuEntry = False
  
End Function

Public Sub CreateHook(ByVal lngHwnd As Long)
  lngPrevProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub ReleaseHook(ByVal lngHwnd As Long)
  
  SetWindowLong lngHwnd, GWL_WNDPROC, lngPrevProc

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next
    WindowProc = CallWindowProc(lngPrevProc, hWnd, uMsg, wParam, lParam)
    If uMsg = WM_SYSCOMMAND Then
      If wParam = intItemID Then
        ShellExecute frmMain.hWnd, "open", "http://www.datosoftware.com", 0&, "", vbNormalFocus
      End If
    End If
    
End Function


