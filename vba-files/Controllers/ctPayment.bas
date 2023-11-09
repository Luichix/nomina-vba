Attribute VB_Name = "ctPayment"
Option Explicit

'namespace=vba-files\Controllers

Public banderaPago As Long

Public Function LanzarPago(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_Pago

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Pago.StartUpPosition = 0
            frm_Pago.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Pago.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_Pago.Show

End Function

Sub InsertarPago()

    If frm_Pago.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una categoria", vbInformation
        frm_Pago.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaPago
     Case 1
        frm_Personal.txt_Pago = frm_Pago.lbx_cuenta.Column(0)

        Unload frm_Pago
     Case 2
        frm_Personal.txt_APago = frm_Pago.lbx_cuenta.Column(0)

        Unload frm_Pago

     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub

