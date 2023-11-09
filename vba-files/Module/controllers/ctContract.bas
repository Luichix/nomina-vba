Attribute VB_Name = "ctContract"
Option Explicit

'namespace=vba-files\controllers

Public banderaContrato As Long

Public Function LanzarContrato(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_Contrato

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Contrato.StartUpPosition = 0
            frm_Contrato.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Contrato.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_Contrato.Show

End Function
Sub Insertarcontrato()

    If frm_Contrato.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una categoria", vbInformation
        frm_Contrato.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaContrato
     Case 1
        frm_Personal.txt_Contrato = frm_Contrato.lbx_cuenta.Column(0)

        Unload frm_Contrato
     Case 2
        frm_Personal.txt_Acontrato = frm_Contrato.lbx_cuenta.Column(0)

        Unload frm_Contrato


     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub

