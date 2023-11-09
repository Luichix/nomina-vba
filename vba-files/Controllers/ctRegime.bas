Attribute VB_Name = "ctRegime"
Option Explicit

'namespace=vba-files\Controllers

Public banderaRegimen As Long

Public Function LanzarRegimen(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_Regimen

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Regimen.StartUpPosition = 0
            frm_Regimen.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Regimen.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_Regimen.Show

End Function
Sub InsertarRegimen()

    If frm_Regimen.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una categoria", vbInformation
        frm_Regimen.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaRegimen
     Case 1
        frm_Personal.txt_Regimen = frm_Regimen.lbx_cuenta.Column(0)

        Unload frm_Regimen
     Case 2
        frm_Personal.txt_Aregimen = frm_Regimen.lbx_cuenta.Column(0)

        Unload frm_Regimen

     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub
