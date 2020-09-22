Attribute VB_Name = "modProgressInStatusBar"
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

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Type RECT
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
End Type

Const WM_USER    As Long = &H400
Const SB_GETRECT As Long = (WM_USER + 10)


Public Sub ShowProgressInStatusBar(ByRef Progress As Control, ByRef StatusBar As StatusBar, ByVal PanelNumber As Long)

    Dim TRC As RECT
    
        StatusBar.Panels(PanelNumber).Width = Progress.Width + 15
        SendMessageAny StatusBar.hWnd, SB_GETRECT, PanelNumber - 1, TRC
               
        With Progress
            SetParent .hWnd, StatusBar.hWnd
            .Move TRC.Left + ((TRC.Right - TRC.Left) / 2) - (.Width / 2), TRC.Top + ((TRC.Bottom - TRC.Top) / 2) - (.Height / 2), .Width, .Height
            .Visible = True
            .Value = 0
        End With
        
End Sub
