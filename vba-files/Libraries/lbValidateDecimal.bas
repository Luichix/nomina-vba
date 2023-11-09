Attribute VB_Name = "lbValidateDecimal"
Option Explicit

'namespace=vba-files\Libraries


Public Function ValidateDecimal(texbox As MSForms.textbox, KeyAscii As MSForms.ReturnInteger) As Integer

    Dim size As Integer
    Dim text As String

    text = texbox.Text
    size = Len(text)
    ' Verify If the first character is a number 
    ' After that any character can be a number Or dot
    If size > 0 Then

        If CountDecimalPoints(text) >= 1 Then

            If KeyAscii > 47 And KeyAscii < 58 Then
                KeyAscii = KeyAscii
            Else
                KeyAscii = 0
            End If

        Else
            ' Entry only numbers Or point
            If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Then

                KeyAscii = KeyAscii

            Else
                KeyAscii = 0

            End If

        End If

    Else

        ' Entry only number
        If KeyAscii > 47 And KeyAscii < 58 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    End If

    ValidateDecimal = KeyAscii
End Function


' Function To count the number of point in the string
Private Function CountDecimalPoints(text As String) As Integer

    Dim size As Integer
    Dim i As Integer
    Dim character As String
    Dim counter As Integer
    counter = 0

    size = Len(text)
    For i = 1 To size
        character = Mid(text, i, 1)
        If character = "." Then
            counter = counter + 1

        End If
        Next
        CountDecimalPoints = counter
End Function

