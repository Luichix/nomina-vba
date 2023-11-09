Attribute VB_Name = "lbFileOpen"
Option Explicit

'namespace=vba-files\libraries

Function IsFileOpen(FileName As String)
    Dim iFilenum As Long
    Dim iErr As Long

    On Error Resume Next
    iFilenum = FreeFile()
        Open FileName For Input Lock Read As #iFilenum
        Close iFilenum
        iErr = Err
        On Error Goto 0

            Select Case iErr
             Case 0:    IsFileOpen = False
             Case 70:   IsFileOpen = True
             Case Else: Error iErr
            End Select

End Function

