Sub Yearly_Stockanalysis()

Dim Rowcounter As Long
Dim lastrow As Long
Dim ws As Worksheet
Dim Total_StockVolume As Double
Dim Openprice As Double
Dim Percentagechange As Double
Dim i As Double
Dim Ticker As String


For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Price Change"
ws.Cells(1, 11).Value = "Yearly percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"




    Rowcounter = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Openprice = Cells(2, 3).Value
    
    For r = 2 To lastrow
    
        If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
        	Pricechange = ws.Cells(r, 6).Value - Openprice
        	
	If Openprice = 0 Then
          	     Percentagechange = 0
          	     Else
          	     Percentagechange = Round(((100 * Pricechange) / Openprice), 2)
          	End If
        
        	Totalstockvolume = Totalstockvolume + ws.Cells(r, 7).Value
        	ws.Cells(Rowcounter, 9).Value = ws.Cells(r, 1).Value
        	ws.Cells(Rowcounter, 10).Value = Pricechange
        	ws.Cells(Rowcounter, 11).Value = Percentagechange
        	ws.Cells(Rowcounter, 12).Value = Totalstockvolume
        
        	If Pricechange < 0 Then
                     ws.Cells(Rowcounter, 10).Interior.ColorIndex = 3
                     Else
                     ws.Cells(Rowcounter, 10).Interior.ColorIndex = 4
        	End If
        
        	Openprice = ws.Cells(r + 1, 3).Value
        	Totalstockvolume = 0
        	Rowcounter = Rowcounter + 1
                
        	Else
        
        	Totalstockvolume = Totalstockvolume + Cells(r, 7).Value
        
        End If
        
    Next r
    
    'Calculating GREATEST PERCENTAGE INCREASE and it's ticker
    i = ws.Cells(2, 11).Value
    
    For r = 2 To lastrow
 
        If ws.Cells(r, 11).Value > i Then
        	i = ws.Cells(r, 11).Value
        	Ticker = ws.Cells(r, 9).Value
        End If
    
    Next r
    
    ws.Cells(2, 15).Value = Ticker
    ws.Cells(2, 16).Value = i
    
    'Calculating GREATEST PERCENTAGE DECREASE and it's ticker
    i = ws.Cells(2, 11).Value
    
    For r = 2 To lastrow
 
        If ws.Cells(r, 11).Value < i Then
        	i = ws.Cells(r, 11).Value
        	Ticker = ws.Cells(r, 9).Value
        End If
    
    Next r
    
    ws.Cells(3, 15).Value = Ticker
    ws.Cells(3, 16).Value = i
    
    'Calculating GREATEST TOTAL VOLUME and it's ticker
    i = ws.Cells(2, 12).Value
    
    For r = 2 To lastrow
 
        If ws.Cells(r, 12).Value > i Then
        	i = ws.Cells(r, 12).Value
        	Ticker = ws.Cells(r, 9).Value
        End If
    
    Next r
    
    ws.Cells(4, 15).Value = Ticker
    ws.Cells(4, 16).Value = i
    
    
Next ws
End Sub
