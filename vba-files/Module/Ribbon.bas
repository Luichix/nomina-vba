Attribute VB_Name = "Ribbon"
Option Explicit
Option Base 1
Public CintaDeRibbon As IRibbonUI
Public RetVal(54) As Boolean

#If VBA7 And Win64 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (Byval hwnd As LongPtr, Byval lpOperation As String, Byval lpFile As String, Byval lpParameters As String, Byval lpDirectory As String, Byval nShowCmd As Long) As LongPtr
    #Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (Byval hwnd As Long, Byval lpOperation As String, Byval lpFile As String, Byval lpParameters As String, Byval lpDirectory As String, Byval nShowCmd As Long) As Long
    #End If


Sub CargarCinta(CintaDeExcel As IRibbonUI)
    Set CintaDeRibbon = CintaDeExcel
    form_iniciosesion.Show
End Sub

'////////////////////// Llamadas desde la Cinta para ejectuar cada formulario ///////////////////////////////////

Sub Boton1(Control As IRibbonControl)
    Hoja0.Select
    Hoja0.Cells(1, 1).Select

End Sub

Sub Boton2(Control As IRibbonControl)

    Application.ScreenUpdating = False

    If Hoja1.Visible = xlSheetVisible Then
        Hoja3.Visible = xlSheetVisible
        Hoja4.Visible = xlSheetVisible

        If Hoja10.Visible = xlSheetVisible Then

            frm_Personal.Show
        Elseif Hoja10.Visible = xlSheetVeryHidden Then
            Hoja10.Visible = xlSheetVisible
            frm_Personal.Show
            Hoja10.Visible = xlSheetVeryHidden
        End If

    Elseif Hoja1.Visible = xlSheetVeryHidden Then
        Hoja1.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible
        Hoja4.Visible = xlSheetVisible
        Hoja10.Visible = xlSheetVisible

        frm_Personal.Show

        Hoja1.Visible = xlSheetVeryHidden
        Hoja3.Visible = xlSheetVeryHidden
        Hoja4.Visible = xlSheetVeryHidden
        Hoja10.Visible = xlSheetVeryHidden

    End If

    Application.ScreenUpdating = True

End Sub

Sub Boton25(Control As IRibbonControl)

    Application.ScreenUpdating = False


    Dim Seguridad As String
    Seguridad = Hoja83.Range("L1").Text


    If Hoja2.Visible = xlSheetVisible Then
        If Hoja83.Visible = xlSheetVisible Then
            Hoja58.Unprotect (Seguridad)
            frm_Hora_Marca.Show
            Hoja58.Protect (Seguridad)
        Elseif Hoja83.Visible = xlSheetVeryHidden Then
            Hoja83.Visible = xlSheetVisible
            Hoja58.Unprotect (Seguridad)
            frm_Hora_Marca.Show
            Hoja58.Protect (Seguridad)
            Hoja83.Visible = xlSheetVeryHidden
        End If

    Elseif Hoja2.Visible = xlSheetVeryHidden Then
        Hoja2.Visible = xlSheetVisible
        Hoja58.Unprotect (Seguridad)
        frm_Hora_Marca.Show
        Hoja58.Protect (Seguridad)
        Hoja2.Visible = xlSheetVeryHidden
    End If

    Application.ScreenUpdating = True
End Sub

Sub Boton26(Control As IRibbonControl)

    Application.ScreenUpdating = False
    Importar_Data
    Application.ScreenUpdating = True
End Sub

Sub Boton27(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja13.Visible = xlSheetVisible Then
        frm_Comisiones.Show
    Elseif Hoja13.Visible = xlSheetVeryHidden Then
        Hoja13.Visible = xlSheetVisible
        frm_Comisiones.Show
        Hoja13.Visible = xlSheetVeryHidden
    End If

    Application.ScreenUpdating = True
End Sub
Sub Boton31(Control As IRibbonControl)

End Sub
Sub Boton32(Control As IRibbonControl)

End Sub
Sub Boton4(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja18.Visible = xlSheetVisible Then
        frm_Viatico.Show
    Elseif Hoja18.Visible = xlSheetVeryHidden Then
        Hoja18.Visible = xlSheetVisible
        frm_Viatico.Show
        Hoja18.Visible = xlSheetVeryHidden
    End If

    Application.ScreenUpdating = True

End Sub

Sub Boton22(Control As IRibbonControl)


End Sub

Sub Boton23(Control As IRibbonControl)

End Sub
Sub Boton24(Control As IRibbonControl)

End Sub
Sub Boton7(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja17.Visible = xlSheetVisible Then
        Hoja18.Visible = xlSheetVisible
        frm_Ajuste.Show
    Elseif Hoja17.Visible = xlSheetVeryHidden Then
        Hoja17.Visible = xlSheetVisible
        Hoja18.Visible = xlSheetVisible
        frm_Ajuste.Show
        Hoja17.Visible = xlSheetVeryHidden
        Hoja18.Visible = xlSheetVeryHidden
    End If

    Application.ScreenUpdating = True
End Sub
Sub Boton8(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja5.Visible = xlSheetVisible Then
        Hoja4.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible
        Hoja21.Visible = xlSheetVisible

        frm_Reporte_Jornada.Show

    Elseif Hoja5.Visible = xlSheetVeryHidden Then
        MsgBox ("Acceso no Autorizado: Debe de ingresar desde un usuario Administrativo..!"), vbCritical, "Gestor de Recursos Humanos"
    End If

    Application.ScreenUpdating = True


End Sub

Sub Boton9(Control As IRibbonControl)
    Application.ScreenUpdating = False

    Dim Seguridad As String
    Seguridad = Hoja83.Range("L1").Text

    Hoja58.Unprotect (Seguridad)
    frm_Calendario_Asistencia.Show
    Hoja58.Protect (Seguridad)

    Application.ScreenUpdating = True

End Sub

Sub Boton28(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja5.Visible = xlSheetVisible Then
        Hoja4.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible

        frm_Colilla.Show

    Elseif Hoja5.Visible = xlSheetVeryHidden Then
        MsgBox ("Acceso no Autorizado: Debe de ingresar desde un usuario Administrativo..!"), vbCritical, "Gestor de Recursos Humanos"
    End If

    Application.ScreenUpdating = True

End Sub

Sub Boton29(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja5.Visible = xlSheetVisible Then
        Hoja4.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible

        frm_General.Show

    Elseif Hoja5.Visible = xlSheetVeryHidden Then
        MsgBox ("Acceso no Autorizado: Debe de ingresar desde un usuario Administrativo..!"), vbCritical, "Gestor de Recursos Humanos"
    End If

    Application.ScreenUpdating = True
End Sub
Sub Boton30(Control As IRibbonControl)
    Dim Mensaje As String
    Dim xNombre As String
    Dim Contrase�a As String
    Dim Seguridad As String

    Application.ScreenUpdating = False

    Seguridad = Hoja83.Range("L1").Text

    If Hoja20.Visible = xlSheetVisible Then
        Hoja4.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible
        Hoja5.Visible = xlSheetVisible

        xNombre = Hoja11.Range("J2").Text

        Mensaje = MsgBox("Esta seguro que desea almacenar los datos de la quincena?" + Chr(13) + "�Desea Continuar?", _
        vbYesNo + vbQuestion, "Grabar Quincena")

        On Error Resume Next

        If Mensaje = vbYes Then

            Contrase�a = InputBox("Digite la clave de permiso", "Reporte de Base de Datos")

            If Contrase�a = Seguridad Then
                Reporte_Historico
                MsgBox "Datos grabados con �xito...!", vbInformation, "Gestor de Recursos Humanos"

            Else
                MsgBox "No se han grabado los datos..!", vbInformation, "Gestor de Recursos Humanos"
            End If
        End If


    Elseif Hoja20.Visible = xlSheetVeryHidden Then
        MsgBox ("Acceso no Autorizado: Debe de ingresar desde un usuario Administrativo..!"), vbCritical, "Gestor de Recursos Humanos"
    End If

    Application.ScreenUpdating = True


End Sub
Sub Boton38(Control As IRibbonControl)
    Dim Mensaje As String
    Dim xNombre As String
    Dim Contrase�a As String
    Dim Seguridad As String

    Application.ScreenUpdating = False

    Seguridad = Hoja83.Range("L1").Text

    If Hoja20.Visible = xlSheetVisible Then
        Hoja4.Visible = xlSheetVisible
        Hoja3.Visible = xlSheetVisible
        Hoja5.Visible = xlSheetVisible

        xNombre = Hoja11.Range("J2").Text

        Mensaje = MsgBox("Esta seguro que desea exportar el reporte quincenal " & xNombre & "?" + Chr(13) + "�Desea Continuar?", _
        vbYesNo + vbQuestion, "Exportar Excel")

        On Error Resume Next

        If Mensaje = vbYes Then

            Contrase�a = InputBox("Digite la clave de permiso", "Reporte de Base de Datos")

            If Contrase�a = Seguridad Then
                Exportar_Excel
                MsgBox "Reporte generado con �xito...!", vbInformation, "Gestor de Recursos Humanos"

            Else
                MsgBox "No se ha generado el reporte..!", vbInformation, "Gestor de Recursos Humanos"
            End If
        End If


    Elseif Hoja20.Visible = xlSheetVeryHidden Then
        MsgBox ("Acceso no Autorizado: Debe de ingresar desde un usuario Administrativo..!"), vbCritical, "Gestor de Recursos Humanos"
    End If

    Application.ScreenUpdating = True


End Sub
Sub Boton39(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja27.Visible = xlSheetVisible Then
        frm_Incapacidad.Show
    Elseif Hoja27.Visible = xlSheetVeryHidden Then
        Hoja27.Visible = xlSheetVisible
        frm_Incapacidad.Show
        Hoja27.Visible = xlSheetVeryHidden
    End If

    Application.ScreenUpdating = True
End Sub
Sub Boton13(Control As IRibbonControl)
End Sub

Sub Boton14(Control As IRibbonControl)

    Application.ScreenUpdating = False
    If Hoja14.Visible = xlSheetVisible Then
        If Hoja11.Visible = xlSheetVisible Then

            frm_Exonera.Show

        Elseif Hoja11.Visible = xlSheetVeryHidden Then
            Hoja11.Visible = xlSheetVisible

            frm_Exonera.Show

            Hoja11.Visible = xlSheetVeryHidden
        End If
    Elseif Hoja14.Visible = xlSheetVeryHidden Then
        Hoja14.Visible = xlSheetVisible
        Hoja11.Visible = xlSheetVisible

        frm_Exonera.Show

        Hoja14.Visible = xlSheetVeryHidden
        Hoja11.Visible = xlSheetVeryHidden

    End If

    Application.ScreenUpdating = True

End Sub

Sub Boton15(Control As IRibbonControl)
    Application.ScreenUpdating = False

    If Hoja15.Visible = xlSheetVisible Then
        If Hoja11.Visible = xlSheetVisible Then
            frm_Anular.Show
        Elseif Hoja11.Visible = xlSheetVeryHidden Then
            Hoja11.Visible = xlSheetVisible

            frm_Anular.Show

            Hoja11.Visible = xlSheetVeryHidden
        End If
    Elseif Hoja15.Visible = xlSheetVeryHidden Then
        Hoja15.Visible = xlSheetVisible
        Hoja11.Visible = xlSheetVisible
        frm_Anular.Show
        Hoja15.Visible = xlSheetVeryHidden
        Hoja11.Visible = xlSheetVeryHidden
    End If
    Application.ScreenUpdating = True

End Sub
Sub Boton16(Control As IRibbonControl)

    Application.ScreenUpdating = False

    If Hoja82.Visible = xlSheetVisible Then

        frm_NuevoUsuario.Show

    Elseif Hoja82.Visible = xlSheetVeryHidden Then
        If Hoja83.Range("H1").Text = "ADMINISTRADOR" Then
            Hoja82.Visible = xlSheetVisible

            frm_NuevoUsuario.Show

            Hoja82.Visible = xlSheetVeryHidden
        Else
            MsgBox ("Acceso no Autorizado: El usuario actual no posee los privilegios para realizar esta acci�n..!"), vbCritical, "Gestor de Recursos Humanos"
        End If
    End If

    Hoja0.Select
    Application.ScreenUpdating = True

End Sub
Sub Boton17(Control As IRibbonControl)

    Application.ScreenUpdating = False

    If Hoja82.Visible = xlSheetVisible Then

        frm_EliminarUsuario.Show

    Elseif Hoja82.Visible = xlSheetVeryHidden Then
        If Hoja83.Range("H1").Text = "ADMINISTRADOR" Then
            Hoja82.Visible = xlSheetVisible

            frm_EliminarUsuario.Show

            Hoja82.Visible = xlSheetVeryHidden
        Else
            MsgBox ("Acceso no Autorizado: El usuario actual no posee los privilegios para realizar esta acci�n..!"), vbCritical, "Gestor de Recursos Humanos"
        End If
    End If

    Hoja0.Select
    Application.ScreenUpdating = True

End Sub
Sub Boton18(Control As IRibbonControl)


    Application.ScreenUpdating = False

    If Hoja82.Visible = xlSheetVisible Then

        frm_Modificar_Permisos.Show

    Elseif Hoja82.Visible = xlSheetVeryHidden Then
        If Hoja83.Range("H1").Text = "ADMINISTRADOR" Then
            Hoja82.Visible = xlSheetVisible

            frm_Modificar_Permisos.Show

            Hoja82.Visible = xlSheetVeryHidden
        Else
            MsgBox ("Acceso no Autorizado: El usuario actual no posee los privilegios para realizar esta acci�n..!"), vbCritical, "Gestor de Recursos Humanos"
        End If
    End If

    Hoja0.Select
    Application.ScreenUpdating = True


End Sub
Sub Boton37(Control As IRibbonControl)
End Sub

Sub Boton33(Control As IRibbonControl)


End Sub

Sub Boton34(Control As IRibbonControl)
End Sub
Sub Boton35(Control As IRibbonControl)


End Sub
Sub Boton36(Control As IRibbonControl)

End Sub

Sub Boton19(Control As IRibbonControl)
    form_iniciosesion.Show
End Sub
Sub Boton20(Control As IRibbonControl)
    ThisWorkbook.Save
End Sub



'//////////////////// Retornos del estado de cada bot�n ////////////////////////


Public Sub DesactivarBoton1(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(1)

End Sub


Public Sub DesactivarBoton2(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(2)

End Sub

Public Sub DesactivarBoton3(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(3)

End Sub

Public Sub DesactivarBoton4(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(4)

End Sub


Public Sub DesactivarBoton5(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(5)

End Sub


Public Sub DesactivarBoton6(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(6)

End Sub


Public Sub DesactivarBoton7(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(7)

End Sub

Public Sub DesactivarBoton8(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(8)

End Sub

Public Sub DesactivarBoton9(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(9)

End Sub
Public Sub DesactivarBoton10(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(10)

End Sub
Public Sub DesactivarBoton11(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(11)

End Sub
Public Sub DesactivarBoton12(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(12)

End Sub
Public Sub DesactivarBoton13(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(13)

End Sub
Public Sub DesactivarBoton14(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(14)

End Sub
Public Sub DesactivarBoton15(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(15)

End Sub
Public Sub DesactivarBoton16(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(16)

End Sub
Public Sub DesactivarBoton17(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(17)

End Sub
Public Sub DesactivarBoton18(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(18)

End Sub
Public Sub DesactivarBoton19(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(19)

End Sub
Public Sub DesactivarBoton20(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(20)

End Sub

Public Sub DesactivarBoton28(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(28)

End Sub

Public Sub DesactivarBoton29(Control As IRibbonControl, Byref ValorBloqueo)
    ValorBloqueo = RetVal(29)

End Sub
