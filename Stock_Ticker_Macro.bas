Attribute VB_Name = "Module1"
Sub StockTicker()

'Create a script that loops through all the stocks for one year and outputs the following information:

'BONUS 2: Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.
For Each ws In Worksheets

Dim TotalRecords As Long
    TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Step 1: The ticker symbol.
    'set initial variable for holding ticker name
    Dim Ticker_Name As String
    Dim Total_Volume As LongLong
    Total_Volume = 0
    
    'keep track of location for each ticker name in summary column
    Dim Summary_Ticker As Long
    Summary_Ticker = 2
       
    'add header
    ws.Range("J1, Q1").Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 18).Value = "Value"
    ws.Range("P1").EntireColumn.AutoFit
    
    'loop through all the ticker names
    For i = 2 To TotalRecords
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("J" & Summary_Ticker).Value = Ticker_Name
            ws.Range("M" & Summary_Ticker).Value = Total_Volume
            
'Step 2: Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        'find open price at year begin, close price at year end; then year end - year begin
        ws.Cells(Summary_Ticker, 11).Value = ws.Cells(i, 6).Value - ws.Cells(i - 250, 3).Value
'Step 5: Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
        'conditional formatting
            If ws.Cells(Summary_Ticker, 11).Value > 0 Then
                ws.Cells(Summary_Ticker, 11).Interior.Color = vbGreen
            Else: ws.Cells(Summary_Ticker, 11).Interior.Color = vbRed
            End If
            
'Step 3: The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        'find percent change
        ws.Cells(Summary_Ticker, 12).Value = ws.Cells(i, 6).Value / ws.Cells(i - 250, 3).Value - 1
        'format to percent
        ws.Range("L2:L" & TotalRecords).NumberFormat = "0.00%"
        'add one
        Summary_Ticker = Summary_Ticker + 1
            Total_Volume = 0
'Step 4: The total stock volume of the stock.
          Else: Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
            End If

    Next i
    
'BONUS 1: Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    'Calculate the maximum and minimum percentage change
Dim MaxVal As Double
Dim MinVal As Double
Dim MaxVol As LongLong
Dim Ticker_Max As String
Dim Ticker_Min As String
Dim Ticker_Vol As String
MaxVal = 0
MinVal = 0
MaxVol = 0
    
For i = 2 To Summary_Ticker
    If ws.Cells(i, 12).Value > MaxVal Then
        MaxVal = ws.Cells(i, 12).Value
        Ticker_Max = ws.Cells(i, 10).Value
    ElseIf ws.Cells(i, 12).Value < MinVal Or MinVal = 0 Then
        MinVal = ws.Cells(i, 12).Value
        Ticker_Min = ws.Cells(i, 10).Value
    End If
    If ws.Cells(i, 13).Value > MaxVol Then
        MaxVol = ws.Cells(i, 13).Value
        Ticker_Vol = ws.Cells(i, 10).Value
    End If
Next i

' Write the results to the worksheet
ws.Cells(2, 18).Value = MaxVal
ws.Cells(3, 18).Value = MinVal
ws.Cells(4, 18).Value = MaxVol
ws.Cells(2, 17).Value = Ticker_Max
ws.Cells(3, 17).Value = Ticker_Min
ws.Cells(4, 17).Value = Ticker_Vol
ws.Range("R2:R3").NumberFormat = "0.00%"
ws.Range("M1, R1").EntireColumn.AutoFit
    
Next

End Sub
