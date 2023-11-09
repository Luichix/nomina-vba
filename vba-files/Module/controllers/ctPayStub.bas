Attribute VB_Name = "ctPayStub"
Option Explicit

'namespace=vba-files\controllers

Public banderaColillaPago As Long

Public Function LanzarColillaPago(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_ColillaPago

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_ColillaPago.StartUpPosition = 0
            frm_ColillaPago.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_ColillaPago.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_ColillaPago.Show

End Function

Sub InsertarColillaPago()

    If frm_ColillaPago.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una categoria..!", vbInformation
        frm_ColillaPago.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaColillaPago
     Case 1
        frm_General.txt_ColillaPago = frm_ColillaPago.lbx_cuenta.Column(0)

        Unload frm_ColillaPago


     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub
