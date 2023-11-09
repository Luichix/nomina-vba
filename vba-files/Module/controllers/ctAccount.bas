Attribute VB_Name = "ctAccount"
Option Explicit

'namespace=vba-files\controllers

Public banderaCuenta As Long

Public Function LanzarCuenta(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_Cuentapersonal

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_Cuentapersonal.Show

End Function
Sub Insertarcuenta()

    If frm_Cuentapersonal.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una cuenta", vbInformation
        frm_Cuentapersonal.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaCuenta
     Case 1
        frm_Cuenta.txt_ccuenta = frm_Cuentapersonal.lbx_cuenta.Column(0)
        frm_Cuenta.txt_cuenta = frm_Cuentapersonal.lbx_cuenta.Column(1)

        Unload frm_Cuentapersonal

     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub
