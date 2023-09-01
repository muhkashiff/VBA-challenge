# VBA-challenge
___________________________
Description of the project
___________________________
The project assignment VBA-Challenge for stock analysis has been completed to provide user information about historical trend of stocks.
Key Parameters determined are:
* Yearly change
* Percentage Change
* Total Stock volume
  and top performer and low performers values were determined too.
_________________________________________
Action File:Script file
_________________________________________
Script file has been attached in the repository as "Stock analysis.bas".
____________________________
Code
____________________________
For the description purpose the code has been provided here too.
Sub stock_calculations()


' Setting Variable for worksheet
    Dim Ws As Worksheet
    
' for looping in worksheets
For Each Ws In Worksheets
    
'Ticker Name holder variable
    Dim Ticker As String
    Ticker = " "
        
'Setting Variables to be used
    
    Dim highestTickerName As String
    Dim lowestTickerName As String
    Dim highestPercent As Double
    Dim lowestPercent As Double
    Dim highestVolumeTicker As String
    Dim highestVolume As Double
    Dim totalTickerVolume As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Change_Price As Double
    Dim Change_Percent As Double

'Assiging initial values to variables
    
    highestTickerName = " "
    lowestTickerName = " "
    highestPercent = 0
    lowestPercent = 0
    highestVolumeTicker = " "
    highestVolume = 0
    totalTickerVolume = 0
    openingPrice = 0
    closingPrice = 0
    changePrice = 0
    changePercent = 0
'Rows and last Row assignment & i

Dim summaryTableRow As Long
    summaryTableRow = 2
        
Dim lastRow As Long
Dim i As Long
        
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Table title assignments

Ws.Cells(1, 9).Value = "Ticker"
Ws.Cells(1, 10).Value = "Yearly Change"
Ws.Cells(1, 11).Value = "Percent Change"
Ws.Cells(1, 12).Value = "Total Stock Volume"
Ws.Cells(2, 15).Value = "Greatest % increase"
Ws.Cells(3, 15).Value = "Greatest % decrease"
Ws.Cells(4, 15).Value = "Greatest total Volume"
Ws.Cells(1, 16).Value = "Ticker"
Ws.Cells(1, 17).Value = "Value"

'Setting opening price value
openingPrice = Ws.Cells(2, 3).Value
        
        
        For i = 2 To lastRow
        
    'compare the values in same and next cell in the column
            
    If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
                ' Set the ticker name, we are ready to insert this ticker name data
        Ticker = Ws.Cells(i, 1).Value
        'price and percent calculations
                
        closingPrice = Ws.Cells(i, 6).Value
        changePrice = closingPrice - openingPrice
   
   'Reference:https://github.com/ibaloyan/Stock_Analysis_with_VBA
        
        If openingPrice <> 0 Then
        changePercent = (changePrice / openingPrice) * 100
        End If
                
                totalTickerVolume = totalTickerVolume + Ws.Cells(i, 7).Value
              
                
                
                Ws.Range("I" & summaryTableRow).Value = Ticker
                
               Ws.Range("J" & summaryTableRow).Value = changePrice
                
        'Fill color in the cells
                
                If (changePrice > 0) Then
                    
                   Ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                ElseIf (changePrice <= 0) Then
                    
                   Ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                End If
                
                 
               Ws.Range("K" & summaryTableRow).Value = (CStr(changePercent) & "%")
                
                Ws.Range("L" & summaryTableRow).Value = totalTickerVolume
                
               
                summaryTableRow = summaryTableRow + 1
                
           'set values to zero before going to other loop
                changePrice = 0
                
                closingPrice = 0
                
                openingPrice = Ws.Cells(i + 1, 3).Value
              
                
    'Reference: chatgpt search
                  'Max ,min and highest volume calculation
                     highestPercent = Application.WorksheetFunction.Max(Range("K2:K" & lastRow))
                     Ws.Cells(2, 17).Value = highestPercent
     
                    lowestPercent = Application.WorksheetFunction.Min(Range("K2:K" & lastRow))
                    Ws.Cells(3, 17).Value = lowestPercent
                    
                    highestVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastRow))
                     Ws.Cells(4, 17).Value = highestVolume
                
       
                
                
                changePercent = 0
                totalTickerVolume = 0
                
            Else
                
                totalTickerVolume = totalTickerVolume + Ws.Cells(i, 7).Value
            End If
            
      
        Next i

           
        
     Next Ws
     
     End Sub


___________________________
References:
____________________________
Stock Analysis with VBA:Ibaloyan(2018)https://github.com/ibaloyan/Stock_Analysis_with_VBA
Chat GTP to search Greatest increase,Decrease and volume values. in the code.
Mostly class room assignments and analogies have been used to come to conclusion of this assignment.
____________________________

