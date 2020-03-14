Sub symbolnew():

    'loop through each worksheet
    For Each ws In Worksheets
    
    'Declare all variables
        Dim Worksheetname As String
        Dim tickername As String
        Dim summary_row As Integer
        Dim yearopen As Double
        Dim yearclose As Double
        Dim LastRow As Double
        Dim LastRowNew As Double
        Dim max As Double
        Dim min As Double
        Dim mintick As String
        Dim tickmax As String
        Dim tickvol As String
        Dim maxvol As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        
        'set cell values for column names

        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = " Total Volume"

        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest Percent Increase"
        ws.Cells(3, 16).Value = "Greatest Percent Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"

        Worksheetname = ws.Name
        'MsgBox (Worksheetname)

        'Get number of rows in the sheet
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)

        summary_row = 2
        totalvolume = 0
        max = 0
        maxvol = 0
        yearopen = Cells(2, 3).Value

            'loop through the sheet until the last row is reached to find ticker name,
             'yearly change,percent change,total volume and populate records
            
            For i = 2 To LastRow

    
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                    tickername = ws.Cells(i, 1).Value
            
                    totalvolume = totalvolume + ws.Cells(i, 7).Value
        
                    ws.Cells(summary_row, 10).Value = tickername
             
                    ws.Cells(summary_row, 13).Value = totalvolume
             
                    yearclose = ws.Cells(i, 6).Value
             
                    yearlychange = yearclose - yearopen
              
                    ws.Cells(summary_row, 11).Value = yearlychange
              
                        If (yearopen = 0) Then
                        
                            percentchange = 0
                        
                        Else
                        
                            percentchange = (yearlychange / yearopen)
                            
                            ws.Cells(summary_row, 12).Value = percentchange
                            ws.Cells(summary_row, 12).Style = "Percent"
                           
                            summary_row = summary_row + 1
                            
                        
                        End If
                                
             
                        yearopen = ws.Cells(i + 1, 3).Value
            
                        totalvolume = 0
             
             
             
                  Else
    
                        totalvolume = totalvolume + ws.Cells(i, 7).Value
         
               
               
                  End If
     
        
            Next i

            'Get the number of rows of the summary table 1
            
            LastRowNew = ws.Cells(Rows.Count, 12).End(xlUp).Row
            'MsgBox (LastRowNew)
               
               'Loop through the summary table 1 till the last row to set the colour formatting
                For i = 2 To LastRowNew
                
                    If ws.Cells(i, 12).Value > 0 Then
                    
                        ws.Cells(i, 12).Interior.Color = vbGreen
                
                    Else
                    
                        ws.Cells(i, 12).Interior.Color = vbRed
                        
                    End If
                    
                Next i
              
              'Find the greatest percent increase and corresponding ticker value
              
                For i = 2 To LastRowNew

                    If ws.Cells(i, 12).Value > max Then
                        max = ws.Cells(i, 12).Value
                        
                        tickmax = ws.Cells(i, 10).Value
        
                    End If
    
                Next i

                ws.Cells(2, 18).Value = max
                ws.Cells(2, 18).Style = "Percent"
                ws.Cells(2, 17).Value = tickmax
    
    'Find the greatest volume
    
                For i = 2 To LastRowNew

                    If ws.Cells(i, 13).Value > maxvol Then
                        maxvol = ws.Cells(i, 13).Value
                        tickvol = ws.Cells(i, 10).Value
        
                    End If
    
                Next i

                ws.Cells(4, 18).Value = maxvol
                
                ws.Cells(4, 17).Value = tickvol



'Find greatest percent decrease and the corresponding ticker value

        min = Application.WorksheetFunction.min(ws.Columns("L"))

                For i = 2 To LastRowNew

                    If (ws.Cells(i, 12).Value = min) Then
                        mintick = ws.Cells(i, 10).Value
        
                    End If
    
                Next i

            ws.Cells(3, 18).Value = min
            ws.Cells(3, 18).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = mintick

        Next ws

End Sub

