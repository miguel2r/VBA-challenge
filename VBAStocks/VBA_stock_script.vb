'created by Miguel Rojas on 7/10/2020
'sub routing that let us to execute the macro on each sheet

Sub main()
Dim ws As Worksheet
 For Each ws In Sheets
    ws.Activate
    Call labels
    Call submit
 Next ws
End Sub


Sub labels()
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Open"
Cells(1, 11).Value = "Close"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Stock Volumen"
Cells(1, 18).Value = "Ticker"
Cells(1, 19).Value = "Value"
Cells(2, 17).Value = "Greatest % Increase"
Cells(3, 17).Value = "Greatest % Decrease"
Cells(4, 17).Value = "Greatest Total Volumen"


End Sub

Sub submit()


Dim counter, open_price, closing_price, price_change, old_price, per_change, volumen, tot_vol, old_volumen As Double
Dim old_p_change, old_p_dchange As Long
Dim nextv, currenv As Long
Dim flag As String



counter = 2
nextv = 1
currentv = 0
price_change = 0
rowresult = 2
tot_vol = 0
flag = "y"
old_p_change = 0
old_p_dchange = 0
old_volumen = 0


'Do Until ThisWorkbook.Sheets("sheet1").Cells(counter, 1).Value = ""
Do Until Cells(counter, 1).Value = ""

   currentv = Cells(counter, 1).Value
   nextv = Cells(counter + 1, 1).Value
   'do this for initial row
   If counter = 2 Then
      open_price = Cells(counter, 3).Value
      Cells(rowresult, 10).Value = open_price
      old_price = open_price
   End If
   'perform a couple of tasks once you find the next stock row
   If currentv <> nextv Then
     Cells(rowresult, 9).Value = currentv
     If counter <> 3 Then
        open_price = Cells(counter + 1, 3).Value
        Cells(rowresult + 1, 10).Value = open_price 'J column
        close_price = Cells(counter, 6).Value
        Cells(rowresult, 11).Value = close_price 'K column
        'price change calculation
        price_change = close_price - old_price
        Cells(rowresult, 12).Value = price_change 'L column
        '% change calculation
        Debug.Print price_change; old_price; currentv
        per_change = 0
        If price_change <> 0 And old_price <> 0 Then
         per_change = (price_change * 100) / old_price
        End If
        Debug.Print per_change; price_change; old_price
        
        Cells(rowresult, 13).Value = per_change  'M column
        'identifying the greatets % increase
        If per_change > old_p_change Then
           Cells(2, 19).Value = per_change
           Cells(2, 18).Value = currentv
           old_p_change = per_change
        End If
        'identifying the greatets % decrease
        If per_change < old_p_dchange Then
           Cells(3, 19).Value = per_change
           Cells(3, 18).Value = currentv
           old_p_dchange = per_change
        End If
        ' change the color of the cell based on (+ or -)
        If per_change < 0 Then
           Cells(rowresult, 13).Interior.ColorIndex = 3
           Else
            Cells(rowresult, 13).Interior.ColorIndex = 43
        End If
        
        'Volumen calculation
        volumen = Cells(counter, 7).Value
        tot_vol = tot_vol + volumen
        Cells(rowresult, 14).Value = tot_vol   'N column volumen
        old_price = open_price
        'identify the greates volumen
        If tot_vol > old_volumen Then
           Cells(4, 19).Value = tot_vol
           Cells(4, 18).Value = currentv
           old_volumen = tot_vol
        End If
        
        tot_vol = 0
        flag = "n"
     End If
     rowresult = rowresult + 1
   End If
   If flag <> "n" Then
      volumen = Cells(counter, 7).Value
      tot_vol = tot_vol + volumen
   End If

   flag = "y"
   'Debug.Print tot_vol; volumen; flag; counter; counter
   counter = counter + 1
Loop
End Sub


