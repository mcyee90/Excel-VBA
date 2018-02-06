Sub stockdata()

' Declare Current as a worksheet object variable.
Dim wSheet As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each wSheet In ThisWorkbook.Worksheets

        wSheet.Activate

'===================================================================================================================================================
        'Sub stockvolume()

        Range("J1").Value = "Ticker"
        Range("I1").Value = "Total Volume"
            'set stock name variable
            Dim stock As String

            'set intial volume, always = 0 to start
            Dim volume As Double
            volume = 0
         
            'track summary table data
            Dim SummaryTableRow As Integer
                SummaryTableRow = 2
                   
            Dim LastRow As Long
                LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
            For j = 2 To LastRow
  
                If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
    
                    stock = Cells(j, 1).Value
                    volume = volume + Cells(j, 7).Value
                    Range("J" & SummaryTableRow).Value = stock
                    Range("I" & SummaryTableRow).Value = volume
                    SummaryTableRow = SummaryTableRow + 1
                    volume = 0

                Else
                    volume = volume + Cells(j, 7).Value

                End If
        
            Next j
        
        'End Sub
        
'=====================================================================================================================================================

        'Sub openprice()

        'Dim LastRow As Double
            LastRow = Cells(Rows.Count, 2).End(xlUp).Row

        Range("K1").Value = "Open"

        Dim opendate As Double

        'Dim SummaryTableRow As Integer
            SummaryTableRow = 2
       
        opendate = WorksheetFunction.Min(Range("B2:B" & LastRow))

        For i = 2 To LastRow
  
            If Cells(i, 2).Value = opendate Then
        
                Range("K" & SummaryTableRow).Value = Cells(i, 3).Value
                SummaryTableRow = SummaryTableRow + 1

            End If
        
        Next i
            

        'End Sub
'======================================================================================================================================================

        'Sub closeprice()

        'Dim LastRow As Long
            LastRow = Cells(Rows.Count, 2).End(xlUp).Row

        Range("L1").Value = "Close"

        Dim closedate As Double

        'Dim SummaryTableRow As Integer
            SummaryTableRow = 2
       
        closedate = WorksheetFunction.Max(Range("B2:B" & LastRow))

        For i = 2 To LastRow
  
            If Cells(i, 2).Value = closedate Then
        
                Range("L" & SummaryTableRow).Value = Cells(i, 6).Value
                SummaryTableRow = SummaryTableRow + 1

            End If
        
        Next i


        'End Sub


'=====================================================================================================================================================

        'Sub percentchange()

        Dim SummaryLastRow As Long
            SummaryLastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
        Range("M1").Value = "Difference"
        Range("H1").Value = "% Change"
    
        For i = 2 To SummaryLastRow
            Cells(i, 13).Value = Cells(i, 12).Value - Cells(i, 11).Value
    
            If Cells(i, 12).Value = 0 Then
            Cells(i, 8).Value = Cells(i, 13).Value
    
            Else
            Cells(i, 8).Value = Cells(i, 13).Value / Cells(i, 12).Value
    
            End If
    
            If Cells(i, 13) > 0 Then
                Cells(i, 13).Interior.ColorIndex = 4
            ElseIf Cells(i, 13) = 0 Then
                Cells(i, 13).Interior.ColorIndex = 6
            Else
                Cells(i, 13).Interior.ColorIndex = 3
            End If
    
    
        Next i

        'End Sub

'===================================================================================================================================================

'Sub maxmin()

Range("Q1").Value = "Value"
Range("p1").Value = "Ticker"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

    SummaryLastRow = Cells(Rows.Count, 8).End(xlUp).Row
    'MsgBox (SummaryLastRow)

Dim Maxincrease As Double
    Maxincrease = WorksheetFunction.Max(Range("H2:H" & SummaryLastRow))
    
Dim Maxdecrease As Double
    Maxdecrease = WorksheetFunction.Min(Range("H2:H" & SummaryLastRow))
    
Dim Maxvolume As Double
    Maxvolume = WorksheetFunction.Max(Range("I2:I" & SummaryLastRow))
    
    Range("Q2").Value = Maxincrease
    Range("Q2").NumberFormat = "0.0%"
    
    Range("Q3").Value = Maxdecrease
    Range("Q3").NumberFormat = "0.0%"
    
    Range("Q4").Value = Maxvolume
    
    
Dim stockname As String

'For m = 2 To SummaryLastRow
    
    'If Range("Q2").Value = Cells(m, 14).Value Then
    '    Range("P2").Value = Cells(m, 9).Value
    '    MsgBox ("True")
    'Else
    '    MsgBox ("False")
    'End If
    
    'Exit For
    
'Next m
    
'For n = 2 To SummaryLastRow
    
    'If Range("Q3") = Cells(n, 14).Value Then
    '    Cells(n, 9).Value = Range("P3").Value
    'End If
    
'Exit For
    
'Next n
Range("P2").Value = WorksheetFunction.VLookup(Maxincrease, Range("H2:M" & SummaryLastRow), 3, False)
Range("P3").Value = WorksheetFunction.VLookup(Maxdecrease, Range("H2:M" & SummaryLastRow), 3, False)
Range("P4").Value = WorksheetFunction.VLookup(Maxvolume, Range("I2:M" & SummaryLastRow), 2, False)

Range("H1").EntireColumn.Insert
    

'End Sub

Range("I2:I" & SummaryLastRow).NumberFormat = "0.0%"
Range("J2:J" & SummaryLastRow).NumberFormat = "0.00"
Range("L2:L" & SummaryLastRow).NumberFormat = "0.00"
Range("M2:M" & SummaryLastRow).NumberFormat = "0.00"
Range("N2:N" & SummaryLastRow).NumberFormat = "0.00"


'================================================================================================================================

    Next wSheet


End Sub
