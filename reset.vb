Sub reset()

' Declare Current as a worksheet object variable.
Dim wSheet As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each wSheet In ThisWorkbook.Worksheets

        wSheet.Activate

Range("H1:S50000").Value = ""
Range("H1:S50000").Interior.ColorIndex = 0


Next wSheet

End Sub
