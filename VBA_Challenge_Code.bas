Attribute VB_Name = "Module1"
Sub VBAHomework_Final()

' Repeating through Worksheets & Setting up Headers

For Each ws In Worksheets:
Worksheets(ws.Name).Activate

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'-----------------

' Variables for Ticker Summary
Dim Ticker_Name As String

Dim Ticker_Total As Double
Ticker_Total = 0

Dim Ticker_Open As Double
Dim Ticker_Close As Double

Dim Yearly_Change As Double
Dim Percent_Change As Double

' Summary Table Location
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Variables for Challenge Portion
Dim Percent_Column As range

Dim Percent_Max As Double
Dim Percent_Min As Double

Dim Total_Volume As range
Dim Max_Volume As Double

' Last Row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'-----------------

' Summarizing Tickers

For i = 2 To lastrow

    ' Find the Ticker Open Value
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
               
        Ticker_Open = Cells(i, 3).Value
       
    End If
         
    ' Find the Ticker Close Value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
           Ticker_Name = Cells(i, 1).Value
           
           Ticker_Total = Ticker_Total + Cells(i, 7).Value
           
            ' Put the Ticker in the Summary Table
            range("I" & Summary_Table_Row).Value = Ticker_Name
     
            ' Put the Ticker total in the Summary Table
            range("L" & Summary_Table_Row).Value = Ticker_Total
             
            ' Find Ticker CLose
            Ticker_Close = Cells(i, 6).Value
           
            ' Set Yearly Change & add to Summary Table
            Yearly_Change = Ticker_Close - Ticker_Open
           
           range("J" & Summary_Table_Row).Value = Yearly_Change
           
                If Yearly_Change < 0 Then
                    range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                   
                    Else
                    range("J" & Summary_Table_Row).Interior.ColorIndex = 43
                   
                End If
                             
                If Ticker_Close = 0 Or Ticker_Open = 0 Then
                    range("K" & Summary_Table_Row).Value = 0
                   
                    Else
                    range("K" & Summary_Table_Row).Value = (Ticker_Close / Ticker_Open) - 1
               
                End If
           
           ' Add one to Summary Table Row (for next Ticker)
           Summary_Table_Row = Summary_Table_Row + 1
           
            ' Reset Ticker Total
            Ticker_Total = 0
       
       Else
   
            Ticker_Total = Ticker_Total + Cells(i, 7).Value
                           
    End If
           
Next i

'-----------------

' Finding Max Percent
Set Percent_Column = range("K:K")
Percent_Max = Application.WorksheetFunction.Max(Percent_Column)
range("P2").Value = Percent_Max

' Finding Min Percent
Set Percent_Column = range("K:K")
Percent_Min = Application.WorksheetFunction.Min(Percent_Column)
range("P3").Value = Percent_Min

' Finding Max Volume
Set Total_Volume = range("L:L")
Max_Volume = Application.WorksheetFunction.Max(Total_Volume)
range("P4").Value = Max_Volume

' Finding Tickers for Max Volume, Max Percent and Min Percent
For i = 2 To lastrow
    If Cells(i, 12).Value = Max_Volume Then
        range("O4").Value = Cells(i, 9).Value
    End If
       
    If Cells(i, 11).Value = Percent_Min Then
        range("O3").Value = Cells(i, 9).Value
    End If
       
    If Cells(i, 11).Value = Percent_Max Then
        range("O2").Value = Cells(i, 9).Value
    End If

Next i

'-------------------

Next ws

'-------------------

End Sub
