

Sub sheetloop()
Dim ws As Worksheet
For Each ws In Worksheets
ws.Select
Call stockanalyzer
Next

End Sub


Sub stockanalyzer()



' Set headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' oval - opening stock value, cval - closing value
Dim oval As Double
Dim cval As Double
' totvol - total stock volume,
' ticksumrow - count the row where ticker summary is stored
Dim totvol As Double
Dim ticksumrow As Integer

totvol = 0
ticksumrow = 1

Dim lastRow As Long

' Determine the length of sheet by counting number of rows
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


oval = Cells(2, 3).Value 'assign opening stock value
totvol = Cells(2, 7).Value 'assign initial volume

' go through every row
    For i = 2 To lastRow
    
    'start adding up volume first time ticker starts
    'move down a row for storing next ticker's summary info
  
            
        ' if the ticker is the same
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        
            totvol = Cells(i + 1, 7).Value + totvol
        
        
        ' if the ticker changes
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' move to the next row for summary info
             ticksumrow = ticksumrow + 1
            ' record closing value
            cval = Cells(i, 6).Value
        
            ' record name of ticker in summary colum
            Cells(ticksumrow, 9).Value = Cells(i, 1).Value
        
            ' determine yearly change by subtracting starting stock value from the year-end closing value
            Cells(ticksumrow, 10).Value = cval - oval
            
            
            ' if the yearly change is greater than zero, color cell green
            If Cells(ticksumrow, 10).Value > 0 Then
            
                Cells(ticksumrow, 10).Interior.ColorIndex = 4
                
            ' if the yearly change is zero or less, then color the cell red
            ElseIf Cells(ticksumrow, 10).Value <= 0 Then
            
                Cells(ticksumrow, 10).Interior.ColorIndex = 3
                
            End If
                    
            
            ' determine percentage change by dividing the yearly change by the starting value
            Cells(ticksumrow, 11).Value = Cells(ticksumrow, 10).Value / oval
            Cells(ticksumrow, 11).Value = Format(Cells(ticksumrow, 11).Value, "#.##%")
        
            'assign total stock volume to the last column of summary
            Cells(ticksumrow, 12).Value = totvol

            'assign opening value to the next ticker's
            oval = Cells(i + 1, 3).Value
            'assign volume value to the next ticker's
            totvol = Cells(i + 1, 7).Value
                        
        End If
        
    Next i
    
End Sub

