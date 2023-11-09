Attribute VB_Name = "ctSchedule"
Option Explicit

'namespace=vba-files\controllers

Public banderaJornada As Long

Public Function LanzarJornada(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_Jornada

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_Jornada.StartUpPosition = 0
            frm_Jornada.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_Jornada.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_Jornada.Show

End Function

Sub InsertarJornada()

    If frm_Jornada.lbx_cuenta.ListIndex = -1 Then
        MsgBox "Debe seleccionar una categoria", vbInformation
        frm_Jornada.lbx_cuenta.SetFocus
     Exit Sub
    End If

    Select Case banderaJornada
     Case 1
        frm_Personal.txt_Jornada = frm_Jornada.lbx_cuenta.Column(0)

        Unload frm_Jornada
     Case 2
        frm_Personal.txt_Ajornada = frm_Jornada.lbx_cuenta.Column(0)

        Unload frm_Jornada

     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub
