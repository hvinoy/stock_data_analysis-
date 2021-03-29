Attribute VB_Name = "Stockdata"
Sub stock_data()


'loop thru all worksheets


 Dim WS_Count As Integer
         'Dim w As Integer

       
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For w = 1 To WS_Count

            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
   
Worksheets(ActiveWorkbook.Worksheets(w).Name).Activate

Dim lastrow As Long
Dim sum_row As Integer
sum_row = 2
Dim ticker_name As String
Dim beg_price As Double
beg_price = Cells(2, 3).Value

Dim end_price As Double
Dim price_change As Double
Dim percent As Double
Range("I1").Value = "Ticker"
Dim volume As Double
volume = 0
Range("J1").Value = "Yearly Change"
Range("k1").Value = "Percent Change"
Range("l1").Value = "Total Volume"

Dim lastrow2 As Long
Dim start As Double
start = 0
Dim ticker_nm As String
ticker_nm = ""
Dim start2 As Double
start2 = 0
Dim increase As String
increase_ticker = ""
Dim start3 As Double
start3 = 0
Dim decrease_ticker As String
decrease_ticker = ""




''Ticker name

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Range("K2:K" & lastrow).NumberFormat = ".##%" 'decimals and percenatge conversion

For I = 2 To lastrow

If Cells(I + 1, 1).Value <> Cells(I, 1) Then
    ticker_name = Cells(I, 1).Value
    Cells(sum_row, 9).Value = ticker_name


'Yearly price change
    end_price = Cells(I, 6).Value
    price_change = end_price - beg_price
    Cells(sum_row, 10).Value = price_change
    
 'Percentage change
    If price_change = 0 Or beg_price = 0 Then
    percent = 0
    
    Else
    
    percent = (price_change / beg_price)
    Cells(sum_row, 11).Value = percent
    
'color index

    If percent >= 0 Then
        Cells(sum_row, 10).Interior.ColorIndex = 4
    Else
        Cells(sum_row, 10).Interior.ColorIndex = 3
    End If
    End If
    
'calculates total volume
    volume = volume + Cells(I, 7).Value
    Cells(sum_row, 12).Value = volume


    sum_row = sum_row + 1
    beg_price = Cells(I + 1, 3).Value
    price_change = 0
    volume = 0
    
Else
    volume = volume + Cells(I, 7).Value
    

End If



Next I


'bonus

Range("p1").Value = "Ticker"
Range("q1").Value = "Value"

lastrow2 = Cells(Rows.Count, 12).End(xlUp).Row

For j = 2 To lastrow2

'greatest increase

Range("o2").Value = " Greatest % increase"
If (Cells(j, 11) >= start) Then
    start = Cells(j, 11).Value
    Cells(2, 17).Value = start
    Range("Q2").NumberFormat = ".##%"   ''''percent sign
    increase_ticker = Cells(j, 9).Value
  
    Cells(2, 16).Value = increase_ticker

End If

'greatest decrease

Range("o3").Value = " Greatest % decrease"
If (Cells(j, 11).Value <= start2) Then
    start2 = Cells(j, 11).Value
    Cells(3, 17).Value = start2
    Range("Q3").NumberFormat = ".##%"    ''''percent sign
    decrease_ticker = Cells(j, 9).Value
    Cells(3, 16).Value = decrease_ticker
 
End If


'highest volume

Range("o4").Value = " Greatest Total Volume"
If (Cells(j, 12).Value >= start3) Then
    start3 = Cells(j, 12).Value
   Cells(4, 17).Value = start3
   ticker_nm = Cells(j, 9).Value
   Cells(4, 16).Value = ticker_nm

End If


Next j

'AutoFit Every Worksheet Column in a Workbook
  For Each sht In ThisWorkbook.Worksheets
    sht.Cells.EntireColumn.AutoFit
  Next sht
  
       'MsgBox ActiveWorkbook.Worksheets(w).Name

         Next w
  

End Sub
