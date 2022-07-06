Attribute VB_Name = "Module1"
Sub alph()
For Each ws In Worksheets
    
    'creating the variables that will be used'
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim stockVal As Double
    Dim placeCounter As Integer
    
    'setting the initial values outside of the for loop
    openPrice = ws.Range("C2").Value
    placeCounter = 2 'starting at 2 bc that is the row that the data starts at'
    stockVal = 0
    
    'the ws. function allows you to peruse all the data sheets'
    'also setting the titles for the columns'
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    
    
    For i = 2 To 22771
        'this if statement activates when the forloop is on the last ticker of a set'
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'getting the value for the last close price'
            closePrice = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            
            'printing out the ticker and calculations'
            ws.Cells(placeCounter, 10).Value = ticker
            ws.Cells(placeCounter, 11).Value = (closePrice - openPrice)
            ws.Cells(placeCounter, 12).Value = (closePrice - openPrice) / openPrice
            
            'getting the last stock value and printing it out'
            stockVal = stockVal + ws.Cells(i, 7).Value
            ws.Cells(placeCounter, 13).Value = stockVal
            
            'placeCounter just moved down one row each time a ticker set is finished'
            placeCounter = placeCounter + 1
            stockVal = 0
            
            'getting the open price for the next ticker set'
            openPrice = ws.Cells(i + 1, 3).Value
        
        Else
            'just collecting the volume data'
            stockVal = stockVal + ws.Cells(i, 7).Value
            
        End If
        
    Next i
Next ws
End Sub


