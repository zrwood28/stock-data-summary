Attribute VB_Name = "Module1"
Sub VBACHALLENGE()

    ' https://www.youtube.com/watch?v=AlC8a7KyJq0&t=189s
    ' The following code was inspired by the above video, retrieved on 07/01/22
    Dim ws As Integer
        ws = Application.Worksheets.Count

    For j = 1 To ws
        Worksheets(j).Activate
        
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim current_row As Double
        current_row = 2
        
    Range("J2").Value = Range("C2").Value
        
            For i = 2 To lastrow
            
                Dim current_cell As String
                    current_cell = Cells(i, 1).Value
                    
                Dim next_cell As String
                    next_cell = Cells(i + 1, 1).Value
                    
                Dim open_row As Double
                    open_row = Cells(i + 1, 3).Value
                    
                Dim close_row As Double
                    close_row = Cells(i, 6).Value
                    
                Dim vol_row As Double
                    vol_row = Cells(i, 7).Value
                    
                If next_cell = current_cell Then
                
                    Range("L" & current_row).Value = (vol_row) + Range("L" & current_row).Value
            
                ElseIf next_cell <> current_cell Then
            
                    ticker_id = current_cell
                    
                    Range("I" & current_row).Value = ticker_id
                    
                    Range("K" & current_row).Value = ((close_row) - (Range("J" & current_row).Value)) / (Range("J" & current_row).Value)
                    
                    Range("K" & current_row).NumberFormat = "0.00%"
                    
                    Range("J" & current_row).Value = (close_row) - (Range("J" & current_row).Value)
                    
                        If Range("J" & current_row).Value = 0 Then
                        
                            Range("J" & current_row).Interior.ColorIndex = 36
                            
                        ElseIf Range("J" & current_row).Value > 0 Then
                        
                            Range("J" & current_row).Interior.ColorIndex = 35
                            
                        Else
                        
                            Range("J" & current_row).Interior.ColorIndex = 22
                            
                        End If
                    
                    Range("L" & current_row).Value = (vol_row) + Range("L" & current_row).Value
                    
                    current_row = current_row + 1
                    
                    Range("J" & current_row).Value = (open_row)
                        
                        If open_row = 0 Then
                            
                            Range("J" & current_row).Value = " "
                        
                        End If
                
                End If
            
            Next i
            
        'https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofit
        ' The code below was found on this Microsoft forum on 07/04/22
        ' Added to automatically change column widths to fit headers and stock volume numbers
        Worksheets(j).Columns("J:L").AutoFit
        
    Next j
    
    Worksheets(1).Activate
    
End Sub
