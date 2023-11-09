Attribute VB_Name = "lbLabelHand"
Option Explicit

'namespace=vba-files\Libraries

Public Const IDC_HAND = 32649&
#If VBA7 Then
Public Declare PtrSafe Function SetCursor Lib "user32" (Byval hCursor As Long) As Long
Public Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (Byval hInstance As Long, Byval lpCursorName As Long) As Long
    #Else
Public Declare Function SetCursor Lib "user32" (Byval hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (Byval hInstance As Long, Byval lpCursorName As Long) As Long
    #End If
