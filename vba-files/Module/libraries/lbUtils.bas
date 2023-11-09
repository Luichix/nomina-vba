Attribute VB_Name = "lbUtils"
Option Explicit

'namespace=vba-files\libraries

Public Function ErasedText(texbox As MSForms.textbox, KeyCode As MSForms.ReturnInteger) As Integer
    If KeyCode = vbKeyBack Then
        texbox.Text = Empty
    End If
    If KeyCode = vbKeyDelete Then
        texbox.Text = Empty
    End If
End Function


Public Function DoublePoint(texbox As MSForms.textbox, KeyAscii As MSForms.ReturnInteger) As Integer

    If KeyAscii > 47 And KeyAscii < 59 Then

        KeyAscii = KeyAscii

    Else
        KeyAscii = 0

    End If

    DoublePoint = KeyAscii

End Function


Public Function GetLastRecord(Sheet As Worksheet) As Integer
    GetLastRecord = GetNewRecord(Sheet) - 1
End Function


Public Function GetNewRecord(Sheet As Worksheet) As Integer

    Dim Row As Long
    Row = 2

    Do While Sheet.Cells(Row, 1) <> ""
        Row = Row + 1
    Loop

    GetNewRecord = Row

End Function