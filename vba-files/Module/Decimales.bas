Attribute VB_Name = "Decimales"
Public Function ValidarDecimales(texbox As MSForms.textbox, KeyAscii As MSForms.ReturnInteger) As Integer
 
Dim tama�o As Integer
Dim texto As String
 
texto = texbox.Text
tama�o = Len(texto)
Rem verifica si el 1 caracter sea solo numero y
Rem a partir del segun caracter puede ser numero o punto
If tama�o > 0 Then
 
If ContarPuntosDecimales(texto) >= 1 Then
 
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
 
Else
Rem ingreso solo numeros y punto
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 46 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If
 
End If
 
Else
 
Rem Ingreso solo numeros
If KeyAscii > 47 And KeyAscii < 58 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If
End If
 
ValidarDecimales = KeyAscii
End Function


Rem Funcion que Cuenta el numero de puntos de la cadena
Function ContarPuntosDecimales(texto As String) As Integer
 
Dim largo As Integer
Dim i As Integer
Dim caracter As String
Dim contador As Integer
contador = 0
 
largo = Len(texto)
For i = 1 To largo
caracter = Mid(texto, i, 1)
If caracter = "." Then
contador = contador + 1
 
End If
Next
ContarPuntosDecimales = contador
End Function
Public Function TextoBorrado(texbox As MSForms.textbox, KeyCode As MSForms.ReturnInteger) As Integer
If KeyCode = vbKeyBack Then
texbox.Text = Empty
End If
If KeyCode = vbKeyDelete Then
texbox.Text = Empty
End If
End Function
Public Function DoblePunto(texbox As MSForms.textbox, KeyAscii As MSForms.ReturnInteger) As Integer
 
If KeyAscii > 47 And KeyAscii < 59 Then
 
KeyAscii = KeyAscii
 
Else
KeyAscii = 0
 
End If

DoblePunto = KeyAscii

End Function
