# VBA-challenges
Assignment 2 

The below is the script I wrote with VBA.

------------------------------
Sub multipleyears()

'declaring variables

Dim ticker As String
Dim total As Double
Dim yearlychange As Double
Dim openprice As Double
Dim summary_table_row As Integer

'naming the header of the summary table

Range("J1").Value = "Ticker"
Range("K1").Value = "YearlyChange"
Range("L1").Value = "YearlyChangePercent"
Range("M1").Value = "TotalStockVolume"

total = 0
summary_table_row = 2
openprice = Cells(2, 3)

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

       If Cells(i, 1) <> Cells(i + 1, 1) Then

                ticker = Cells(i, 1)
                Range("J" & summary_table_row) = ticker

                total = total + Cells(i, 7)
                Range("M" & summary_table_row) = total
               
                

                yearlychange = Cells(i, 6) - openprice
                Range("K" & summary_table_row) = yearlychange
                If yearlychange < 0 Then
                Range("K" & summary_table_row).Interior.ColorIndex = 3
                Else
                Range("K" & summary_table_row).Interior.ColorIndex = 4
                End If
                

                Range("L" & summary_table_row) = yearlychange / openprice
                Range("L" & summary_table_row).Select
                Selection.Style = "Percent"
                

                'moving to the next open price & next ticker
                
                 openprice = Cells(i + 1, 3)
                 summary_table_row = summary_table_row + 1

       Else

                 total = total + Cells(i, 7)

       End If

Next i

End Sub

Sub max_min_total()

Range("P2").Value = "Greatest % increase"
Range("P3").Value = "Greatest & descrease"
Range("P4").Value = "Greatest Total Volume"
Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

Dim max As Double
Dim maxindex As Integer

max = Cells(2, 12)
maxindex = 2

lastrow = Cells(Rows.Count, 10).End(xlUp).Row

For i = 3 To lastrow

      If Cells(i, 12) > max Then
         max = Cells(i, 12)
         maxindex = i
      
      End If

Next i
      
Range("Q2") = Cells(maxindex, 10)
Range("R2") = Cells(maxindex, 12)
Range("R2").Select
Selection.Style = "percent"

Dim min As Double
Dim minindex As Integer

min = Cells(2, 12)
minindex = 2

For i = 3 To lastrow

      If Cells(i, 12) < min Then
         min = Cells(i, 12)
         minindex = i
      
      

      End If

Next i
      
Range("Q3") = Cells(minindex, 10)
Range("R3") = Cells(minindex, 12)
Range("R3").Select
Selection.Style = "percent"

Dim maxtotal As Double
Dim maxtotalindex As Integer

maxtotal = Cells(2, 13)
maxtotalindex = 2

For i = 3 To lastrow
     
     If Cells(i, 13) > maxtotal Then
    maxtotal = Cells(i, 13)
    maxtotalindex = i
     
     End If
Next i

Range("Q4") = Cells(maxtotalindex, 10)
Range("R4") = Cells(maxtotalindex, 13)

End Sub



