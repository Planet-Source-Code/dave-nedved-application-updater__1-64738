Attribute VB_Name = "modMain"
Option Explicit

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

   Rem // we need to call InitCommonControls before we
   Rem // can use XP visual styles.  Here I'm using
   Rem // InitCommonControlsEx, which is the extended
   Rem // version provided in v4.72 upwards (you need
   Rem // v6.00 or higher to get XP styles)
   On Error Resume Next
   Rem // this will fail if Comctl not available
   Rem //  - unlikely now though!
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
   Rem // now start the application
   On Error GoTo 0
   Rem // Start the application
   Rem // The reason i coded it to start in a module is so i can skin the application.
   frmMain.Show
   
End Sub
