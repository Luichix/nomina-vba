VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
    Caption         =   "Seleccione una fecha"
    ClientHeight    =   3150
    ClientLeft      =   48
    ClientTop       =   396
    ClientWidth     =   2664
    OleObjectBlob   =   "frmCalendario.frx":0000
    StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub cboMes_Click()
    lbCalendar.CambioDeMes
End Sub



Private Sub lbl1_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl1.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl1.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl1_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl1 As Control
    Set Control_lbl1 = frmCalendario.lbl1

    Call lbCalendar.MarcarDia(Control_lbl1)
End Sub

Private Sub lbl10_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl10.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl10.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl10_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl10 As Control
    Set Control_lbl10 = frmCalendario.lbl10

    Call lbCalendar.MarcarDia(Control_lbl10)
End Sub

Private Sub lbl11_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl11.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl11.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl11_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl11 As Control
    Set Control_lbl11 = frmCalendario.lbl11

    Call lbCalendar.MarcarDia(Control_lbl11)
End Sub

Private Sub lbl12_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl12.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl12.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl12_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl12 As Control
    Set Control_lbl12 = frmCalendario.lbl12

    Call lbCalendar.MarcarDia(Control_lbl12)
End Sub

Private Sub lbl13_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl13.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl13.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl13_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl13 As Control
    Set Control_lbl13 = frmCalendario.lbl13

    Call lbCalendar.MarcarDia(Control_lbl13)
End Sub

Private Sub lbl14_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl14.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl14.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl14_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl14 As Control
    Set Control_lbl14 = frmCalendario.lbl14

    Call lbCalendar.MarcarDia(Control_lbl14)
End Sub

Private Sub lbl15_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl15.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl15.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl15_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl15 As Control
    Set Control_lbl15 = frmCalendario.lbl15

    Call lbCalendar.MarcarDia(Control_lbl15)
End Sub

Private Sub lbl16_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl16.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl16.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl16_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl16 As Control
    Set Control_lbl16 = frmCalendario.lbl16

    Call lbCalendar.MarcarDia(Control_lbl16)
End Sub

Private Sub lbl17_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl17.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl17.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl17_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl17 As Control
    Set Control_lbl17 = frmCalendario.lbl17

    Call lbCalendar.MarcarDia(Control_lbl17)
End Sub

Private Sub lbl18_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl18.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl18.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl18_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl18 As Control
    Set Control_lbl18 = frmCalendario.lbl18

    Call lbCalendar.MarcarDia(Control_lbl18)
End Sub

Private Sub lbl19_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl19.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl19.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl19_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl19 As Control
    Set Control_lbl19 = frmCalendario.lbl19

    Call lbCalendar.MarcarDia(Control_lbl19)
End Sub

Private Sub lbl2_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl2.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl2.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl2_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl2 As Control
    Set Control_lbl2 = frmCalendario.lbl2

    Call lbCalendar.MarcarDia(Control_lbl2)
End Sub

Private Sub lbl20_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl20.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl20.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl20_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl20 As Control
    Set Control_lbl20 = frmCalendario.lbl20

    Call lbCalendar.MarcarDia(Control_lbl20)
End Sub

Private Sub lbl21_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl21.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl21.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl21_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl21 As Control
    Set Control_lbl21 = frmCalendario.lbl21

    Call lbCalendar.MarcarDia(Control_lbl21)
End Sub

Private Sub lbl22_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl22.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl22.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl22_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl22 As Control
    Set Control_lbl22 = frmCalendario.lbl22

    Call lbCalendar.MarcarDia(Control_lbl22)
End Sub

Private Sub lbl23_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl23.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl23.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl23_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl23 As Control
    Set Control_lbl23 = frmCalendario.lbl23

    Call lbCalendar.MarcarDia(Control_lbl23)
End Sub

Private Sub lbl24_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl24.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl24.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl24_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl24 As Control
    Set Control_lbl24 = frmCalendario.lbl24

    Call lbCalendar.MarcarDia(Control_lbl24)
End Sub

Private Sub lbl25_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl25.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl25.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl25_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl25 As Control
    Set Control_lbl25 = frmCalendario.lbl25

    Call lbCalendar.MarcarDia(Control_lbl25)
End Sub

Private Sub lbl26_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl26.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl26.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl26_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl26 As Control
    Set Control_lbl26 = frmCalendario.lbl26

    Call lbCalendar.MarcarDia(Control_lbl26)
End Sub

Private Sub lbl27_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl27.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl27.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl27_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl27 As Control
    Set Control_lbl27 = frmCalendario.lbl27

    Call lbCalendar.MarcarDia(Control_lbl27)
End Sub

Private Sub lbl28_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl28.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl28.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl28_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl28 As Control
    Set Control_lbl28 = frmCalendario.lbl28

    Call lbCalendar.MarcarDia(Control_lbl28)
End Sub

Private Sub lbl29_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl29.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl29.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl29_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl29 As Control
    Set Control_lbl29 = frmCalendario.lbl29

    Call lbCalendar.MarcarDia(Control_lbl29)
End Sub

Private Sub lbl3_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl3.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl3.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl3_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl3 As Control
    Set Control_lbl3 = frmCalendario.lbl3

    Call lbCalendar.MarcarDia(Control_lbl3)
End Sub

Private Sub lbl30_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl30.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl30.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl30_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl30 As Control
    Set Control_lbl30 = frmCalendario.lbl30

    Call lbCalendar.MarcarDia(Control_lbl30)
End Sub

Private Sub lbl31_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl31.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl31.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl31_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl31 As Control
    Set Control_lbl31 = frmCalendario.lbl31

    Call lbCalendar.MarcarDia(Control_lbl31)
End Sub

Private Sub lbl32_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl32.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl32.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl32_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl32 As Control
    Set Control_lbl32 = frmCalendario.lbl32

    Call lbCalendar.MarcarDia(Control_lbl32)
End Sub

Private Sub lbl33_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl33.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl33.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl33_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl33 As Control
    Set Control_lbl33 = frmCalendario.lbl33

    Call lbCalendar.MarcarDia(Control_lbl33)
End Sub

Private Sub lbl34_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl34.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl34.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl34_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl34 As Control
    Set Control_lbl34 = frmCalendario.lbl34

    Call lbCalendar.MarcarDia(Control_lbl34)
End Sub

Private Sub lbl35_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl35.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl35.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl35_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl35 As Control
    Set Control_lbl35 = frmCalendario.lbl35

    Call lbCalendar.MarcarDia(Control_lbl35)
End Sub

Private Sub lbl36_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl36.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl36.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl36_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl36 As Control
    Set Control_lbl36 = frmCalendario.lbl36

    Call lbCalendar.MarcarDia(Control_lbl36)
End Sub

Private Sub lbl37_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl37.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl37.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl37_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl37 As Control
    Set Control_lbl37 = frmCalendario.lbl37

    Call lbCalendar.MarcarDia(Control_lbl37)
End Sub

Private Sub lbl38_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl38.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl38.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl38_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl38 As Control
    Set Control_lbl38 = frmCalendario.lbl38

    Call lbCalendar.MarcarDia(Control_lbl38)
End Sub

Private Sub lbl39_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl39.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl39.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl39_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl39 As Control
    Set Control_lbl39 = frmCalendario.lbl39

    Call lbCalendar.MarcarDia(Control_lbl39)
End Sub

Private Sub lbl4_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl4.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl4.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl4_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl4 As Control
    Set Control_lbl4 = frmCalendario.lbl4

    Call lbCalendar.MarcarDia(Control_lbl4)
End Sub

Private Sub lbl40_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl40.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl40.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl40_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl40 As Control
    Set Control_lbl40 = frmCalendario.lbl40

    Call lbCalendar.MarcarDia(Control_lbl40)
End Sub

Private Sub lbl41_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl41.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl41.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl41_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl41 As Control
    Set Control_lbl41 = frmCalendario.lbl41

    Call lbCalendar.MarcarDia(Control_lbl41)
End Sub

Private Sub lbl42_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl42.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl42.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl42_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl42 As Control
    Set Control_lbl42 = frmCalendario.lbl42

    Call lbCalendar.MarcarDia(Control_lbl42)
End Sub

Private Sub lbl5_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl5.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl5.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl5_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl5 As Control
    Set Control_lbl5 = frmCalendario.lbl5

    Call lbCalendar.MarcarDia(Control_lbl5)
End Sub

Private Sub lbl6_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl6.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl6.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl6_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl6 As Control
    Set Control_lbl6 = frmCalendario.lbl6

    Call lbCalendar.MarcarDia(Control_lbl6)
End Sub

Private Sub lbl7_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl7.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl7.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl7_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl7 As Control
    Set Control_lbl7 = frmCalendario.lbl7

    Call lbCalendar.MarcarDia(Control_lbl7)
End Sub

Private Sub lbl8_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl8.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl8.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl8_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl8 As Control
    Set Control_lbl8 = frmCalendario.lbl8

    Call lbCalendar.MarcarDia(Control_lbl8)
End Sub

Private Sub lbl9_DblClick(Byval Cancel As MSForms.ReturnBoolean)
    If frmCalendario.lbl9.Caption <> "-" Then
        Dim Dia As Long, Mes As Long, Ano As Long

        Dia = VBA.CLng(frmCalendario.lbl9.Caption)
        Mes = VBA.CLng(frmCalendario.cboMes.Value)
        Ano = VBA.CLng(frmCalendario.lblAno.Caption)

        Unload frmCalendario
        Call lbCalendar.RecibeLaFecha(Dia, Mes, Ano)
    End If
End Sub

Private Sub lbl9_MouseMove(Byval Button As Integer, Byval Shift As Integer, Byval X As Single, Byval Y As Single)
    Dim Control_lbl9 As Control
    Set Control_lbl9 = frmCalendario.lbl9

    Call lbCalendar.MarcarDia(Control_lbl9)
End Sub

Private Sub cmdSalirConEscape_Click()
    Call lbCalendar.SalirConEscape
End Sub

Private Sub lblHoy_Click()
    lbCalendar.UnClickEnHoyEs
End Sub

Private Sub mrcDias_Click()

End Sub

Private Sub spbA�o_Change()
    lbCalendar.CambioDeAno
End Sub

Private Sub UserForm_Initialize()
    Call lbCalendar.InicializaFormularioCalendario
End Sub
