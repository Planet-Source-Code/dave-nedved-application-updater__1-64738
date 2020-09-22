VERSION 5.00
Begin VB.UserControl ucHorizontal3DLine 
   ClientHeight    =   180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   ScaleHeight     =   180
   ScaleWidth      =   1845
   ToolboxBitmap   =   "ucHorizontal3DLine.ctx":0000
   Begin VB.Line LineA 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   1800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LineB 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   1800
      X2              =   0
      Y1              =   75
      Y2              =   75
   End
End
Attribute VB_Name = "ucHorizontal3DLine"
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
Rem // |      use this program to help you creaste update files.      |
Rem // ----------------------------------------------------------------

Option Explicit

Private Sub UserControl_Resize()
    UserControl.Height = 30
    LineB.X1 = 0
    LineA.X1 = 0
    LineA.X2 = UserControl.Width
    LineB.X2 = UserControl.Width
    LineB.Y1 = 15
    LineB.Y2 = 15
End Sub
