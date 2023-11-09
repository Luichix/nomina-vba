Attribute VB_Name = "seCalculateHours"
Option Explicit

'namespace=vba-files\services

Sub ProcessHourRecords()
    ' Assuming your data is in a sheet named "HourRecords"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HourRecords")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Assuming your data starts from row 2 With headers in row 1
    Dim currentRow As Long
    For currentRow = 2 To lastRow
        Dim penalized As Boolean
        Dim startHour As Integer
        Dim endHour As Integer
        Dim entryHour As Integer
        Dim maxDelayTime As Double
        Dim result As Double

        ' Assuming your data is organized in columns A To E (adjust If needed)
        penalized = ws.Cells(currentRow, 1).Value
        startHour = ws.Cells(currentRow, 2).Value
        endHour = ws.Cells(currentRow, 3).Value
        entryHour = ws.Cells(currentRow, 4).Value
        maxDelayTime = ws.Cells(currentRow, 5).Value

        result = CalculateHoursWorked(penalized, startHour, endHour, entryHour, maxDelayTime)

        ' Assuming you want To store the result in column F (adjust If needed)
        ws.Cells(currentRow, 6).Value = result
    Next currentRow
End Sub

Function CalculateHoursWorked(penalized As Boolean, startHour As Integer, endHour As Integer, entryHour As Integer, maxDelayTime As Double) As Double
    If penalized Then
        MsgBox penalized
    End If

    If Not penalized Then
        Dim maxEntryHour As Double
        Dim analyzedEntryHour As Double
        Dim hours As Double

        maxEntryHour = entryHour + maxDelayTime

        If startHour < maxEntryHour Then
            analyzedEntryHour = entryHour
        Else
            analyzedEntryHour = startHour
        End If

        If endHour > analyzedEntryHour Then
            hours = endHour - analyzedEntryHour
        Else
            hours = 0
        End If

        CalculateHoursWorked = hours
    End If
End Function

