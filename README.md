# VBA_Chanllenge
Sub Multiple_year_stock_data()

    Cells(1, 9).Value = "Ticker"
    
    Cells(1, 10).Value = "Yearly Change"
    
    Cells(1, 11).Value = "Percent Change"
    
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 16).Value = "Ticker"
    
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    
    Cells(3, 15).Value = "Greatest % Decrease"
    
    Cells(4, 15).Value = "Greatest Total Volume"



  Dim Ticker_Name As String
  
  Dim Ticker_Volume As Double
  
  Ticker_Volume = 0
  
  Dim Summary_Table_Row As Integer
  
  Summary_Table_Row = 2
  
   Dim EndRow As Long
   
    EndRow = Cells(Rows.Count, 1).End(xlUp).Row
    
  Dim last_time As Long
  
  last_time = Cells(2, 2).Value
  
  Dim first_time As Long
  
  first_time = Cells(2, 2).Value
  
  Dim val_last_time As Double
  
  val_last_time = Cells(2, 6).Value
  
  Dim val_first_time As Double
  
  val_first_time = Cells(2, 3).Value
  
  Dim yearly_change As Double
  
  Dim percen_change As Double

  For i = 2 To EndRow
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker_Name = Cells(i, 1).Value

      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
      
      yearly_change = val_last_time - val_first_time
      percen_change = ((val_last_time / val_first_time) * 100) - 100


      Range("I" & Summary_Table_Row).Value = Ticker_Name

      Range("L" & Summary_Table_Row).Value = Ticker_Volume
      
      Range("J" & Summary_Table_Row).Value = yearly_change
      Range("K" & Summary_Table_Row).Value = percen_change

      Summary_Table_Row = Summary_Table_Row + 1
      
      Total_Volume = 0
      
      fisrt_time = Cells(i + 1, 2).Value
      last_time = Cells(i + 1, 2).Value
      val_first_time = Cells(i + 1, 3).Value
      val_last_time = Cells(i + 1, 6).Value

    Else

      Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
      
      If Cells(i + 1, 2).Value > last_time Then
        last_time = Cells(i + 1, 2).Value
        val_last_time = Cells(i + 1, 6).Value
      Else
        fist_time = Cells(i + 1, 2).Value
        val_first_time = Cells(i + 1, 3).Value
      End If

    End If

  Next i
  
    Dim rng As Range
    Dim cell As Range
    
    Set rng = Range("J2:J" & Cells(Rows.Count, "J").End(xlUp).Row)
    
  
    For Each cell In rng
       
        If cell.Value > 0 Then
            cell.Interior.Color = RGB(0, 255, 0) ' Green color
        ElseIf cell.Value < 0 Then
            cell.Interior.Color = RGB(255, 0, 0) ' Red color
        End If
    Next cell


Dim Rng_1 As Range
Dim maxPercentChange As Double
Dim maxTicker As String
Dim minPercentChange As Double
Dim minTickerChange As String
Dim Rng_2 As Range
Dim maxVolume As Double
Dim maxVolumeTicker As String

Set Rng_1 = Range("K2:K" & Cells(Rows.Count, "K").End(xlUp).Row)
Set Rng_2 = Range("L2:L" & Cells(Rows.Count, "L").End(xlUp).Row)
    
   
    maxPercentChange = WorksheetFunction.Max(Rng_1)
    maxTicker = WorksheetFunction.Index(Range("I2:I" & Cells(Rows.Count, "I").End(xlUp).Row), WorksheetFunction.Match(maxPercentChange, Rng_1, 0))
    
    minPercentChange = WorksheetFunction.Min(Rng_1)
    minTicker = WorksheetFunction.Index(Range("I2:I" & Cells(Rows.Count, "I").End(xlUp).Row), WorksheetFunction.Match(minPercentChange, Rng_1, 0))
    
    maxVolume = WorksheetFunction.Max(Rng_2)
    maxVolumeTicker = WorksheetFunction.Index(Range("I2:I" & Cells(Rows.Count, "I").End(xlUp).Row), WorksheetFunction.Match(maxVolume, Rng_2, 0))
    
    Cells(2, 16).Value = maxTicker
    Cells(2, 17).Value = maxPercentChange
    Cells(3, 16).Value = minTicker
    Cells(3, 17).Value = minPercentChange
    Cells(4, 16).Value = maxVolumeTicker
    Cells(4, 17).Value = maxVolume

End Sub



