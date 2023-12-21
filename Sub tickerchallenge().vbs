Sub tickerchallenge():

'To run the macro in each of the worksheets
For Each ws In Worksheets

'Give table a name
Dim WorksheetName As String
WorksheetName = ws.Name

'Assign column names that needs to calculated
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

'Decalare variables for ticker, counter and the volume calculation
'Dim Tick_Name As String

Dim Tick_Vol As Long
Tick_Vol = 0

Dim Tick_Count As Long
Tick_Count = 2

'Declare variables for calculating yearly change and percent change
Dim Close_Price As Double
Dim Open_Price As Double
Dim Per_Change As Double

Open_Price = ws.Cells(2, 3).Value

'Create a variable to store total number of rows in worksheet
Dim TotRow As Long
Tot_Row = Range("A1").End(xlDown).Row


For i = 2 To Tot_Row

    'Check till the next cell value is different from the current cell value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Tick_Vol = Tick_Vol + Cells(i, 7).Value

    'Get the ticker name and ticker volume
        Range("I" & Tick_Count).Value = ws.Cells(i, 1).Value
        Range("L" & Tick_Count).Value = Tick_Vol + ws.Cells(i, 7).Value

    'Calculate Yearly change price
        Close_Price = ws.Cells(i, 6).Value
        Range("J" & Tick_Count).Value = Close_Price - Open_Price
    
    'Calculate Percent Change
     If Open_Price = 0 Then
        Range("K" & Tick_Count).Value = 0
     Else
        Range("K" & Tick_Count).Value = Range("J" & Tick_Count).Value / Open_Price
        'Convert the value in percentage format
        Range("K" & Tick_Count).NumberFormat = "0.00%"
     End If

    'Change the ticker count and ticker volume
    Tick_Count = Tick_Count + 1
    Tick_Vol = 0
    
    'Resetting the opening price
    Open_Price = ws.Cells(i + 1, 3).Value

    'Else
   ' Tick_Vol = Tick_Vol + Cells(i, 7).Value

    End If
Next i

'Coloring positive and negative yearly change values
'Calculate rows in new ticker column
Dim RowCountNew As Long
RowCountNew = Range("I1").End(xlDown).Row

For i = 2 To RowCountNew
    If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = vbGreen
            Else
                ws.Cells(i, 10).Interior.Color = vbRed
            End If

'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'Declare 3 variables and store the value as in first entry(2nd row)

Dim GPI As Double   'Greatest percentage increase
Dim GPD As Double   'Greatest percentage decrease
Dim GTV As Double   'Greatest total volume

GPI = ws.Cells(2, 11).Value
GPD = ws.Cells(2, 11).Value
GTV = ws.Cells(2, 12).Value

'Assign cell name to enter the values

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Greatest percentage increase calculation
    If ws.Cells(i, 11).Value > GPI Then
        GPI = ws.Cells(i, 11).Value
        ws.Cells(2, 16) = ws.Cells(i, 9).Value
        
    Else
        GPI = GPI
        
    End If
    
'Greatest percent decrease calculation
If ws.Cells(i, 11).Value < GPD Then
        GPD = ws.Cells(i, 11).Value
        ws.Cells(3, 16) = ws.Cells(i, 9).Value
        
    Else
        GPD = GPD
        
    End If


'Greatest total vloume calculation
    If ws.Cells(i, 12).Value > GTV Then
    
        GTV = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
    Else
        
        GTV = GTV

    End If

     ws.Cells(2, 17).Value = Format(GPI, "Percent")
     ws.Cells(3, 17).Value = Format(GPD, "Percent")
     ws.Cells(4, 17).Value = Format(GTV, "Scientific")

 Next i

'Adjust the column width of all the calculations
 Worksheets(WorksheetName).Columns("A:Z").AutoFit
 
 
Next ws

End Sub















