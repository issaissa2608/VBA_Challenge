Attribute VB_Name = "Module2"


Sub StockMarketData_Analysis()

'Define all  variables
'--------------------------------------------------

Dim Ticker As String

Dim year_open As Double

Dim year_close As Double

Dim Yearly_Change As Double

Dim Total_Stock_Volume As Double

Dim Percent_Change As Double


'Define a variable to set up a row to start

Dim start_data As Integer

'Define variable of the worksheet to excute the code in all work sheet at once in the workbook

Dim ws As Worksheet


'loop in all worksheet to excute the code once
'--------------------------------------------------

For Each ws In Worksheets

    
    'Assign a column header for every task to perform

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

    
    'Assign intiger for the loop to start
        start_data = 2
        previous_i = 1
        Total_Stock_Volume = 0

    
    'Go to the last row of column A

        EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    
    'For each Ticker summrize and loop the yearly change, percent change, and total stock volume

        For i = 2 To EndRow

    'If Tickersymbol change or not equal to the previous one excute to record

             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Get the Tickersymbol

            Ticker = ws.Cells(i, 1).Value

    'Intiate the variable to go to the next Ticker Alphabet

            previous_i = previous_i + 1

    ' Get the value first day open form the column 3 or "C" and last day close of the year on column 6 or "F"

            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value

    ' A for loop to sum the total stock volume using vol which is found in column 7 or "G"

            For j = previous_i To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j

    'When the loop get the value zero open the data

                If year_open = 0 Then

                    Percent_Change = year_close

                Else
                    Yearly_Change = year_close - year_open

                    Percent_Change = Yearly_Change / year_open

                End If
    '--------------------------------------------------

    'Get the values in the worksheet summary table

            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change

    'Use percentage format

            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume

    'In the data summary when the first row task completed go to the next row

            start_data = start_data + 1

    'Get back the variable to zero

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

    'Move i number to variable previous_i
            previous_i = i

        End If

    Next i


'--------------------------------------------------
' Conditional formatting columns colors

'The end row for column J

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To jEndRow

            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

'Excute to next worksheet

Next ws

'--------------------------------------------------
End Sub


