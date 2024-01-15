VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'SCRIPT HAS BEEN DERIVED FROM MANY RESOURCES BUT NOT LIMITED TO YOUTUBE, OVERSTACK, STACKFLOW, GOOGLE, UDEMEY, ZOOM CLASS NOTES


Sub StockAnalysis()

    ' VARIABLES DECLARED FOR THE WORK SHEET
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    'FOR LOOP CREATED TO LOOP THROUH EACH SHEET
    For Each ws In Worksheets
    
' SETTING THE START OF THE ROW OR INITILIZING
summaryRow = 2
        
  ' SETIIGNG VALUES TO BEGIN
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        maxIncreaseTicker = ""
        maxDecreaseTicker = ""
        maxVolumeTicker = ""
   
   
   ' HEADERS TO OUT PUT VALUES IN THE TABLE
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
 ' LAST ROW DATA FROM SHEETS
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
       
       
' LOOPING THROUGH ROWS TO GET DATA AND LATSROW DECLARED HERE
       
        For i = 2 To lastRow
        
   ' IF FUNCTION SEE IF THE TICKER SYMBOL IS THE SAME TO GET THE VALUE
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
         
                
   ' TICKER SYMBOL COLLECTION
                
                ticker = ws.Cells(i, 1).Value
                
   ' COLLECTS OPEN PRICE
                
               openPrice = ws.Cells(i, 3).Value
                
   ' RESETING
                totalVolume = 0
                
            End If
            
   ' ADDING VOLUMES TO CURRENT VOLUME
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
  ' IF FUNCTION CHECKS TO SEE IF THE TICKER IS THE SAME
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
   ' GET CLOSIN PRICE
                closePrice = ws.Cells(i, 6).Value
                
   ' CALCULATE YEARLY CHANGE IN THE PRICE
                yearlyChange = closePrice - openPrice
                
   ' PERCENTAGE CALCULATION  DONE
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
  ' DISPLAY THE TICKER AND THE CHANGES
                
                ws.Range("I" & summaryRow).Value = ticker
                ws.Range("J" & summaryRow).Value = yearlyChange
                ws.Range("K" & summaryRow).Value = percentChange
                ws.Range("L" & summaryRow).Value = totalVolume
                
   ' GIVING FORMAT TO THE CHANGE
                ws.Range("J" & summaryRow).NumberFormat = "0.00"
                ws.Range("K" & summaryRow).NumberFormat = "0.00%"
                
   ' FORMAT COLOR TO THE VALUE BASED ON THE POSITIVE OR NEGATIVE
                If yearlyChange > 0 Then
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 4 ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 0 ' No fill
                End If
                
   ' UPDATE MAX VALUES
                
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If
                
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
    ' Increment the summary row
                summaryRow = summaryRow + 1
                
            End If
            
        Next i
        
   ' Output the max values and tickers to the summary table
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q2").Value = maxIncrease
        ws.Range("Q3").Value = maxDecrease
        ws.Range("Q4").Value = maxVolume
        
  ' Format the max value cells
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0"
        
 
        
    Next ws
    
    
End Sub

