Attribute VB_Name = "Module1"
Sub VBA_Assignment()
'Sort the data
Range("A:G").Sort Key1:=Range("A1"), Header:=xlYes

'copy the data
Range("A:A").Copy Range("I:I")
    
'Remove Duplictaes
Range("I1").Value = "Ticker"
    
Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes

'Total Stock Value

Range("L1").Value = "Total stock Value"
    
Dim LastRow_A As Long
    
LastRow_A = Cells(Rows.Count, "A").End(xlUp).Row
    
j = 2
    
    For i = 2 To LastRow_A
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            Cells(j, 12).Value = TotalStockValue
            
                j = j + 1
                
                    TotalStockValue = 0
        Else
    
            TotalStockValue = TotalStockValue + Cells(i, 7).Value
        
    End If
    
Next i

'Annual Open Price of Ticker

Range("M1").Value = "Annual Open Price"
 
Range("M2").Value = Range("C2").Value
 
 j = 3
   
           For i = 2 To LastRow_A
           
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                    Cells(j, 13).Value = Cells(i + 1, 3).Value
                         
                         j = j + 1
           End If
           
           Next i
           
'Annual close Price of Ticker


Range("N1").Value = "Annual Closing Price"

   j = 2
           For i = 2 To LastRow_A
           
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                    Cells(j, 14).Value = Cells(i, 6).Value
                         
                         j = j + 1
                         
             End If
             
    Next i
    
'calculation of Yearly Change and Percentage change

             
 Range("J1").Value = "Yearly Change"
 
 Range("K1").Value = "Percent Change"
 
 Dim LastRow_ticker As Long
 
 Dim YearlyChange As Double
 
 Dim PercentChange As Double
 
 LastRow_ticker = Cells(Rows.Count, "I").End(xlUp).Row
 
        For i = 2 To LastRow_ticker
        
           YearlyChange = Cells(i, 14).Value - Cells(i, 13).Value
           
                 Cells(i, 10).Value = YearlyChange
                 
                    If YearlyChange = 0 Then
                    
                        Cells(i, 11).Value = 0
                        
                            ElseIf Cells(i, 13).Value = 0 Then
           
                                Cells(i, 11).Value = 0
                 
                        Else
                        
                            PercentChange = YearlyChange / Cells(i, 13).Value
                      
                            Cells(i, 11).Value = PercentChange
           End If
           
          Next i
          
'finding the greatest percent increase and decrease

Dim GreatestPercentIncrease As Double

Dim GreatestPercentDecrease As Double

Dim GreatestTotalVolume As Single

Range("P2").Value = "Greatest Percent Increase"

Range("P3").Value = "Greatest Percent Decrease"

Range("P4").Value = "Greatest Total Volume "

Range("Q1").Value = "Ticker"

Range("R1").Value = "Value"

GreatestPercentIncrease = Application.WorksheetFunction.Max(Range("K:K"))

Range("R2").Value = GreatestPercentIncrease

GreatestPercentDecrease = Application.WorksheetFunction.Min(Range("K:K"))

Range("R3").Value = GreatestPercentDecrease
 
'greatest total stock value

GreatestTotalVolume = Application.WorksheetFunction.Max(Range("L:L"))

Range("R4").Value = GreatestTotalVolume


'format of numbers

Range("K:K").NumberFormat = "0.0%"

Range("R2:R3").NumberFormat = "0.0%"

'format clours

           For i = 2 To LastRow_ticker
           
                 If Cells(i, 11).Value < 0 Then Cells(i, 11).Interior.ColorIndex = 3
                 
                    If Cells(i, 11).Value > 0 Then Cells(i, 11).Interior.ColorIndex = 10
           Next i
           
        Range("K:K").Font.Color = 1
        
Dim a As Double

Dim b As Double

a = Range("R2").Value

b = Range("R3").Value

LastRow_ticker_K = Cells(Rows.Count, "K").End(xlUp).Row

            For i = 2 To LastRow_ticker_K
            
                If Cells(i, 11).Value = a Then
                
                    Cells(2, 17).Value = Cells(i, 9).Value
            End If
            
                If Cells(i, 11).Value = b Then
                
                    Cells(3, 17).Value = Cells(i, 9).Value
            End If
            
Next i

Dim c As Single

Dim LastRow_stckVol As Double

c = Range("R4").Value

LastRow_stckVol = Cells(Rows.Count, "L").End(xlUp).Row

For j = 2 To LastRow_stckVol

    If Cells(j, 12).Value = c Then
    
        Cells(4, 17).Value = Cells(j, 9).Value
End If

Next j

Range("M:N").Delete

End Sub


