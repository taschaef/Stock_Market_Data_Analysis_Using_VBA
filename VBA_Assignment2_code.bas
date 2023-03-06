Attribute VB_Name = "Module1"
Sub stocksloop()

'PART I - Setup'
'-----------------------------------------------------------------------------------------------'
'Loop through all sheets'
For Each ws In Worksheets
ws.Activate

'Read and store given variables - "Retrival of Data"'
Dim ticker_symbol As String
Dim volume_of_stock As Double
Dim open_price As Double
Dim close_price As Double
  
'Make variables and columns with headers for summary table'
Dim ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Vol As Double

'Set correct variables equal to zero'
Yearly_Change = 0
Percent_Change = 0
Total_Stock_Vol = 0
 
'Outline Summary Table'
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
  
'Add column and header for Ticker'
ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9).Value = "Ticker"
  
'Add column and header for Yearly Change'
ws.Range("J1").EntireColumn.Insert
ws.Cells(1, 10).Value = "Yearly Change"
ws.Range("J1").EntireColumn.Style = "Currency"
  
'Add column and header for Percent Change'
ws.Range("K1").EntireColumn.Insert
ws.Cells(1, 11).Value = "Percent Change"
ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
  
'Add column and header for Total Stock Volume'
ws.Range("L1").EntireColumn.Insert
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("L1").EntireColumn.Style = "Normal"

'Autofit columns'
ws.Columns("A:S").AutoFit

          
'Part II - Loop through stock information & create table'
'--------------------------------------------------------------------------------------------------------------'
    'Loop through all stock exchanges'
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For I = 2 To LastRow
  
        'Check to see if we are still within the same stock brand, if not. . .'
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
   
            'Set ticker name in Summary Table'
            ticker = ws.Cells(I, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = ticker
    
                'Figure out Yearly_Change and print in Summary Table'
                Yearly_Change = (ws.Cells(I, 6).Value) - (ws.Cells(I, 3).Value)
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                    'Figure out Percent_Change and print in Summary Table'
                    Percent_Change = (Yearly_Change / ws.Cells(I, 3))
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
                        'Figure out Total_Stock_Vol and print in Summary Table'
                        Total_Stock_Vol = Total_Stock_Vol + ws.Cells(I, 7).Value
                        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Vol
       
                    'Add a row to the Summary Table and reset all variables back to zero'
                    Summary_Table_Row = Summary_Table_Row + 1
                    Yearly_Change = 0
                    Percent_Change = 0
                    Total_Stock_Vol = 0
        
                'If the cell immediately following a row is the same stock brand'
                Else
    
            'Add to Total_Stock_Vol'
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(I, 7).Value
   
            End If
        Next I
Next ws

            
'PART III - conditional formatting'
'--------------------------------------------------------------------------------------------------------------'
'Conditionally format column J, Yearly Change.'
For Each ws In Worksheets
ws.Activate

    For j = 2 To LastRow
    If ws.Cells(j, 10).Value <= 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
        ws.Cells(j, 10).Interior.ColorIndex = 4
     
    End If
        Next j
            Next ws
                
'Conditionally format column K, Percent Change'
For Each ws In Worksheets
ws.Activate

    For k = 2 To LastRow
    If ws.Cells(k, 11).Value <= 0 Then
    ws.Cells(k, 11).Interior.ColorIndex = 3
        Else
        ws.Cells(k, 11).Interior.ColorIndex = 4
    
    End If
        Next k
            Next ws
            

'PART IV - Analysis Table'
'--------------------------------------------------------------------------------------------------------------'
'Loop through all sheets'
For Each ws In Worksheets
ws.Activate

'Insert column and header for Ticker_'
ws.Range("P1").EntireColumn.Insert
ws.Cells(1, 16).Value = "Ticker_"

'Insert column and header for Value'
ws.Range("Q1").EntireColumn.Insert
ws.Cells(1, 17).Value = "Value"
  
'Outline Analysis Table'
Dim Analysis_Table_Row As Integer
Analysis_Table_Row = 2
      
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Columns("A:S").AutoFit
      
'Create variables and set them to 0'
Dim ticker_value1 As String
Dim ticker_value2 As String
Dim ticker_value3 As String
Dim Inc As Double
Dim D As Double
Dim V As Double
       
Inc = 0
D = 0
V = 0
       
ws.Cells(2, 16).Value = ticker_value1
ws.Cells(3, 16).Value = ticker_value2
ws.Cells(4, 16).Value = ticker_value3

'Search Column K, Percent Change, for maximum and minimum'
Inc = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
IncPos = WorksheetFunction.Match(Inc, ws.Range("K2:K" & LastRow), 0)
ticker_value1 = ws.Cells(IncPos + 1, 9).Value
        
D = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
DPos = WorksheetFunction.Match(D, ws.Range("K2:K" & LastRow), 0)
ticker_value2 = ws.Cells(DPos + 1, 9).Value
    
'Search column L, Total Stock Volume, for maximum'
V = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
VPos = WorksheetFunction.Match(V, ws.Range("L2:L" & LastRow), 0)
ticker_value3 = ws.Cells(VPos + 1, 9).Value

'Deposit values in corresponding Analysis Table locations'
ws.Cells(2, 17).Value = Inc
ws.Range("Q2").NumberFormat = "0.00%"
      
ws.Cells(3, 17).Value = D
ws.Range("Q3").NumberFormat = "0.00%"
      
ws.Cells(4, 17).Value = V
ws.Range("Q4").Style = "Normal"
        
Next ws
                
End Sub

