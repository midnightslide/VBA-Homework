Sub stockSolver()

    'turn screenupdating off to speed up macro
    Application.ScreenUpdating = False  

    For Each ws In Worksheets

        Dim stockName As String
        Dim stockTotal As Double
        Dim column As Integer
        Dim tableRow As Long
        Dim yearlyRow As Long
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim rowCounter As Long
        stockTotal = 0
        column = 1
        tableRow = 2
        yearlyRow = 2
        openingPrice = 0
        closingPrice = 0
        rowCounter = 0

        ' find the last row of info in each worksheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' assign and print column titles for all additional info
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' ------------------------------------------------
        '
        '         TICKER NAMES AND VALUE TOTALS
        '
        ' ------------------------------------------------
        
        For i = 2 To lastrow
    
            ' check if still within the same stock ticker
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                
                ' record stock ticker name
                stockName = ws.Cells(i, 1).Value
        
                ' add to stock volume total
                stockTotal = stockTotal + ws.Cells(i, 7).Value
        
                ' print stock ticker name in ticker column
                ws.Range("I" & tableRow).Value = stockName
        
                ' print stock volume amount in stock volume column
                ws.Range("L" & tableRow).Value = stockTotal
        
                ' go to the next row
                tableRow = tableRow + 1
                       
                ' reset stock volume total
                stockTotal = 0
                        
                Else
                ' add to the stock volume total
                stockTotal = stockTotal + ws.Cells(i, 7).Value
               
            End If
         
        Next i
        
        ' ------------------------------------------------
        '
        '   OPENING PRICE VS CLOSING PRICE AND % CHANGE
        '
        ' ------------------------------------------------
        
        For i = CLng(2) To lastrow
    
            ' check if still within the same stock ticker
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                
                ' assign closing price value
                closingPrice = ws.Cells(i, 6).Value
                ' MsgBox ("The Closing Price is " & closingPrice)
                
                ' print the difference of the stock from yearly open to close
                ws.Range("J" & yearlyRow).Value = closingPrice - openingPrice
                
                    If openingPrice > 0 Then
                    ' avoid 0/0 errors
                    ws.Range("K" & yearlyRow).Value = ((closingPrice - openingPrice) / openingPrice) ' * 100
                    
                    Else
                    
                    ' occupy the cell with a zero value to indicate zero change
                    ws.Range("K" & yearlyRow).Value = 0
                    
                    End If
                
                ' go to the next row of side chart
                yearlyRow = yearlyRow + 1
                
                'reset the row counter to zero to grab only the first opening price for each ticker
                rowCounter = 0
        
                ElseIf ws.Cells(i + 1, column).Value = ws.Cells(i, column).Value And rowCounter = 0 Then
                
                ' assign opening price value
                openingPrice = ws.Cells(i, 3).Value
                ' MsgBox ("The Opening Price is " & openingPrice)
                                                            
                ' add 1 to row counter to avoid recording other values
                rowCounter = rowCounter + 1
                              
            End If
                      
        Next i
        
        ' ----------------------------------------------
        '
        '          COLOR-CHANGE-O-MATIC v1.0.1
        '
        ' ----------------------------------------------
        
        For i = 2 To lastrow
        
            ' Color cells with positive change green and ignore zero
            If ws.Range("J" & i).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            ' Color cells with negative change red and ignore zero
            ElseIf ws.Range("J" & i).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i
        
        ' ------------------------------------------------
        '
        '              FORMATTING WIZARDRY
        '
        ' ------------------------------------------------
        
        ' format column width and content type
        ws.Columns("J:J").AutoFit
        ws.Columns("J:J").NumberFormat = "0.000000000"
        ws.Columns("K:K").AutoFit
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Columns("L:L").AutoFit
                    
    Next ws

    ' let me know when the macro is successfully done doing its thing
    MsgBox ("Success!! All stock information has been analyzed and calculated!")

    ' turn screenupdating back on
    Application.ScreenUpdating = True
End Sub