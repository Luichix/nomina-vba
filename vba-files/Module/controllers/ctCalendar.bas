Attribute VB_Name = "ctCalendar"
Option Explicit

'namespace=vba-files\controllers

Public banderaCalendario As Long

Public Function LanzarCalendario(CualquierFormulario As Object, xTextBox As String)
    Dim xCtrl As Control

    Load frmCalendario

    For Each xCtrl In CualquierFormulario.Controls
        If xCtrl.Name = xTextBox Then
            frmCalendario.StartUpPosition = 0
            frmCalendario.Left = CualquierFormulario.Left + xCtrl.Left + 5
            frmCalendario.Top = CualquierFormulario.Top + xCtrl.Top + xCtrl.Height + 25
        End If
        Next

        frmCalendario.Show

End Function

Public Function InsertarFecha(Fecha As Date)

    Dim i As Byte
    Dim txt_Fecha As textbox

    Select Case banderaCalendario
     Case 1
        frm_Personal.txt_Inicio.Text = Fecha

     Case 2
        frm_Personal.txt_Fin.Text = Fecha

     Case 5
        frm_Incapacidad.txt_Inicio.Text = Fecha

     Case 6
        frm_Incapacidad.txt_Fin.Text = Fecha

     Case 7
        frm_Comisiones.txt_Fecha.Text = Fecha

     Case 8
        frm_Hora_Marca.txt_fecha1.Text = Fecha
     Case 9
        frm_Hora_Marca.txt_fecha2.Text = Fecha
     Case 10
        frm_Hora_Marca.txt_fecha3.Text = Fecha
     Case 11
        frm_Hora_Marca.txt_fecha4.Text = Fecha
     Case 12
        frm_Hora_Marca.txt_fecha5.Text = Fecha
     Case 13
        frm_Hora_Marca.txt_fecha6.Text = Fecha
     Case 14
        frm_Hora_Marca.txt_fecha7.Text = Fecha
     Case 15
        frm_Hora_Marca.txt_fecha8.Text = Fecha
     Case 16
        frm_Hora_Marca.txt_fecha9.Text = Fecha
     Case 17
        frm_Hora_Marca.txt_fecha10.Text = Fecha
     Case 18
        frm_Hora_Marca.txt_fecha11.Text = Fecha
     Case 19
        frm_Hora_Marca.txt_fecha12.Text = Fecha
     Case 20
        frm_Hora_Marca.txt_fecha13.Text = Fecha
     Case 21
        frm_Hora_Marca.txt_fecha14.Text = Fecha
     Case 22
        frm_Hora_Marca.txt_fecha15.Text = Fecha
     Case 23
        frm_Hora_Marca.txt_fecha16.Text = Fecha

     Case 25


     Case 27
        frm_Viatico.txt_Fecha.Text = Fecha

     Case 28
        frm_Personal.txt_Ainicio.Text = Fecha

     Case 29
        frm_Personal.txt_Afin.Text = Fecha
     Case Else
        MsgBox "La petici�n solicitada, a�n no se ha establecido dentro de la declaraci�n Select Case", vbCritical
    End Select
End Function
