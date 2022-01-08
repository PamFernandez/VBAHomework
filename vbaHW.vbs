Sub VBAofWallStreet()

    ' variables for code
    Dim result As Long
    Dim subtotal As Double
    Dim i As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim ws As Worksheet
       
    ' to iterate through the sheets
    For Each ws In Worksheets
    ws.Activate
             
        ' create the result table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
             
        ' put results in a table
        result = 2
        
        ' set initial openprice
         openprice = Cells(2, 3).Value
           
        ' iterate through to the last row
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        ' calculate the subtotal for Total Stock Volume
        subtotal = subtotal + Cells(i, 7).Value
            
            ' If we are about to encounter a new ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ' write the name of the current Ticker in the result table
                Cells(result, 9).Value = Cells(i, 1)
                
                ' calculate the Yearly Change
                ' collect information about closeprice
                closeprice = Cells(i, 6).Value
                Cells(result, 10).Value = (closeprice - openprice)
                
                    ' highlight Yearly Change numbers: pos (green), neg (red)
                    If Cells(result, 10).Value < 0 Then
                       Cells(result, 10).Interior.ColorIndex = 3
                    Else
                       Cells(result, 10).Interior.ColorIndex = 4
                    End If
                                        
                    ' calculate the Percent Change - don't divide by zero and write the Percent Change in the result table
                    If (openprice = 0) Then
                        Cells(result, 11).Value = 0
                    Else
                        Cells(result, 11).Value = Cells(result, 10).Value / openprice
                    End If
                    
                Cells(result, 11).NumberFormat = "0.00%"
                    
                'write the Total Stock Volume in the result table
                Cells(result, 12).Value = subtotal
                Cells(result, 12).NumberFormat = "0,000"
                
                ' move to the next ticker symbol
                 result = result + 1
                
                ' reset the total
                subtotal = 0
             
                 ' reset the open price
                 openprice = Cells(i + 1, 3)
            
            End If
  
        Next i
        
        'auto fit the columns in the results table
        Columns("I:L").AutoFit
                   
    Next ws

End Sub