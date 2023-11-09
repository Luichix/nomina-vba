Attribute VB_Name = "lbRemoveHeader"
Option Explicit

'namespace=vba-files\Libraries

'----------------------------------------- APIS ELIMINAR BARRA TITULO FORMULARIO
#If VBA7 Then
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (Byval lpClassName As String, Byval lpWindowName As String) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (Byval hwnd As Long, Byval nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (Byval hwnd As Long, Byval nIndex As Long, Byval dwNewLong As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (Byval hwnd As Long) As Long
    #Else
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (Byval lpClassName As String, Byval lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (Byval hwnd As Long, Byval nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (Byval hwnd As Long, Byval nIndex As Long, Byval dwNewLong As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (Byval hwnd As Long) As Long
    #End If

Sub RemoveHeader(MeCaption)
    Dim lStyle As Long
    Dim hMenu As Long
    Dim mhWndForm As Long
    mhWndForm = FindWindow("ThunderDFrame", MeCaption)
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub
