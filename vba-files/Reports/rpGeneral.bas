Attribute VB_Name = "rpGeneral"
Option Explicit

'namespace=vba-files\Reports

Public Sub Reporte_General()
    Attribute Reporte_General.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Id As String
    Dim Colaborador As String
    Dim Fila As Long
    Dim Final As Long
    Dim TOTAL As String
    Dim separador As String
    Dim CK As String
    Dim ACH As String
    Dim xCol As Long


    CK = Hoja81.Range("G2").Text
    ACH = Hoja81.Range("G3").Text

    separador = Application.International(xlListSeparator)

    Id = "ID"
    Colaborador = "COLABORADOR"
    TOTAL = "TOTAL"

    Hoja16.Activate
    Hoja16.Cells.Select
    Selection.Clear

    Hoja4.Activate
    Hoja4.Cells.Select
    Application.CutCopyMode = False
    Hoja4.Cells.Copy

    Hoja16.Activate
    Hoja16.Cells(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Hoja16.Columns("C:BF").Select
    Selection.Delete

    Hoja16.Columns("S:AC").Select
    Selection.Delete

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("1:2").Select
    Selection.RowHeight = 25

    Rows("3:500").Select
    Selection.RowHeight = 20

    Fila = Hoja16.Range("A" & Rows.Count).End(xlUp).Row
    Final = Fila + 1


    Hoja16.Activate
    Hoja16.Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .VerticalAlignment = xlCenter
    End With


    Hoja16.Activate
    Hoja16.Cells(2, 1).Select
    Hoja16.Cells(2, 1) = Id

    Hoja16.Cells(2, 2).Select
    Hoja16.Cells(2, 2) = Colaborador

    Hoja16.Columns("A:A").Select
    Selection.ColumnWidth = 7
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Hoja16.Columns("C:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Hoja16.Range("A1:R1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With


    Hoja16.Range("A2:R2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With


    Hoja16.Select


    Hoja16.Cells(Fila + 1, 1).Select
    Hoja16.Cells(Fila + 1, 1) = TOTAL



    Hoja16.Range(Cells(2, 1), Cells(Final, 18)).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone


    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With



    Range("A2:R2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range(Cells(2, 1), Cells(Final, 2)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With

    Range("A2:R2").Select
    Selection.Font.Bold = True
    Range("A1:R1").Select
    Selection.Font.Size = 10
    Selection.Font.Bold = True
    Range(Cells(Final, 1), Cells(Final, 18)).Select
    Selection.Font.Bold = True
    Range(Cells(2, 1), Cells(Final, 18)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A2:R2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(Cells(Final, 1), Cells(Final, 18)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Hoja16.Cells(Fila + 1, 4) = WorksheetFunction.CountA(Range(Cells(3, 3), Cells(Fila, 3)))

    Application.DisplayAlerts = False

    Hoja16.Cells(Fila + 1, 1) = "TOTAL PLANILLA " & CK & ":"
    Hoja16.Range(Cells(Fila + 1, 1), Cells(Fila + 1, 3)).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .MergeCells = True
        .InsertIndent 2
    End With
    Application.DisplayAlerts = True

End Sub


