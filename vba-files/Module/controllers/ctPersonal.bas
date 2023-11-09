Attribute VB_Name = "ctPersonal"
Option Explicit

'namespace=vba-files\controllers

Public banderaPersonal As Long

Public Function LanzarListadoPersonal(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frm_ListadoPersonal

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frm_ListadoPersonal.StartUpPosition = 0
            frm_ListadoPersonal.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frm_ListadoPersonal.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frm_ListadoPersonal.Show

End Function

Sub InsertarPersonal()

    If frm_ListadoPersonal.lbx_Personal.ListIndex = -1 Then
        MsgBox "Debe seleccionar un Colaborador", vbInformation
        frm_ListadoPersonal.lbx_Personal.SetFocus
     Exit Sub
    End If

    Select Case banderaPersonal


     Case 4
        frm_Personal.txt_Aid = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Personal.txt_Anombre = frm_ListadoPersonal.lbx_Personal.Column(1)
        Unload frm_ListadoPersonal

     Case 5
        Hoja58.Range("K6") = frm_ListadoPersonal.lbx_Personal.Column(0)
        Unload frm_ListadoPersonal

     Case 6
        frm_Comisiones.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Comisiones.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal
     Case 9
        frm_Exonera.cbx_personal = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Exonera.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal

     Case 10
        frm_Anular.cbx_personal = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Anular.cbx_nombre = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal

     Case 11
        frm_Ajuste.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Ajuste.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal

     Case 12
        frm_Viatico.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Viatico.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal

     Case 13
        frm_Reporte_Jornada.txt_Id = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_Reporte_Jornada.txt_Nombre = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal

     Case 14

     Case 15
        frm_ISR.ComboBox1 = frm_ListadoPersonal.lbx_Personal.Column(0)
        frm_ISR.ComboBox2 = frm_ListadoPersonal.lbx_Personal.Column(1)

        Unload frm_ListadoPersonal



     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select

End Sub

