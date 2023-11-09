Attribute VB_Name = "lbSecurity"
Option Explicit

'namespace=vba-files\Libraries

Public banderaUnprotect As Long
Public banderaProtect As Long

Sub Unprotec()
  Dim Seguridad As String

  Seguridad = Hoja83.Range("L1").Text

  Select Case banderaUnprotect

   Case 1 
    Hoja1.Unprotect (Seguridad)
    Hoja10.Unprotect (Seguridad)
    Hoja3.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)

   Case 2 
    Hoja1.Unprotect (Seguridad)
    Hoja10.Unprotect (Seguridad)
    Hoja3.Unprotect (Seguridad)
    Hoja4.Unprotect (Seguridad)

   Case Else
    MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
  End Select

End Sub

Sub Protect()
  Seguridad As String

  Seguridad = Hoja83.Range("L1")

  Select Case banderaProtect

   Case 1 'Contrataci�n
    Hoja1.Protect (Seguridad)
    Hoja10.Protect (Seguridad)
    Hoja3.Protect (Seguridad)
    Hoja4.Protect (Seguridad)

   Case 2 'Contrataci�n
    Hoja1.Protect (Seguridad)
    Hoja10.Protect (Seguridad)
    Hoja3.Protect (Seguridad)
    Hoja4.Protect (Seguridad)

   Case Else
    MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
  End Select

End Sub

