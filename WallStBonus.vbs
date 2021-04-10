Sub Wallst()

    For Each ws In Worksheets
    
        Dim Ticker As String
        Dim ticker_amount As Double
        Dim counter As Long
        Dim year_change As Double
        Dim percent_change As Double
        Dim summary As Long
         
        ticker_amount = 0
        counter = 0
        year_change = 0
        percent_change = 0
        summary = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To last_row

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
            
            If ws.Cells(i - counter, 3) <> 0 Then
                ticker_amount = ticker_amount + ws.Cells(i, 7).Value
                year_change = ws.Cells(i, 6).Value - ws.Cells(i - counter, 3).Value
                percent_change = year_change / ws.Cells(i - counter, 3).Value
                
            Else
                ticker_amount = 0
                year_change = 0
                percent_change = 0

            End If
                
                ws.Range("I" & summary).Value = Ticker
                ws.Range("L" & summary).Value = ticker_amount
                ws.Range("J" & summary).Value = year_change
                ws.Range("K" & summary).Value = percent_change
                    
                summary = summary + 1
                ticker_amount = 0
                counter = 0

            Else
                ticker_amount = ticker_amount + ws.Cells(i, 7).Value
                counter = counter + 1

            End If
        
        Next i
        
        summary_last = ws.Cells(Rows.Count, "I").End(xlUp).Row

        For i = 2 To summary_last
           
            If ws.Cells(i, "J") < 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 3
            Else
                ws.Cells(i, "J").Interior.ColorIndex = 4
            End If

            ws.Cells(i, "K").NumberFormat = "0.00%"

        Next i
   
        Dim volume_high As Double
        Dim volume_high_ticker As String
        Dim percent_high As Double
        Dim percent_high_ticker As String
        Dim percent_low As Double
        Dim percent_low_ticker As String
        
        volume_high = 0
        percent_high = 0
        percent_low = 0
        
        For i = 2 To summary_last
        
            If ws.Cells(i, "L").Value > volume_high Then
            volume_high = ws.Cells(i, "L").Value
            volume_high_ticker = ws.Cells(i, "I").Value
                
            End If
            
            If ws.Cells(i, "K").Value > percent_high Then
            percent_high = ws.Cells(i, "K").Value
            percent_high_ticker = ws.Cells(i, "I").Value
                
            End If
            
            If ws.Cells(i, "K").Value < percent_low Then
                
            percent_low = ws.Cells(i, "K").Value
            percent_low_ticker = ws.Cells(i, "I").Value
            
            End If
                    
        Next i
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("P2").Value = percent_high_Ticker
        ws.Range("P3").Value = percent_low_ticker
        ws.Range("P4").Value = volume_high_ticker       
        ws.Range("Q2").Value = percent_high
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = percent_low
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = volume_high
        ws.Columns("I:Q").AutoFit

        Next ws
 
 End Sub