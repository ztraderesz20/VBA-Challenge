Sub VBA_WallStreet()


Dim ws As Worksheet

'...Loop through ALL sheets

For Each ws In Worksheets:
Worksheets(ws.Name).Activate

    
'...define and initialize variables

    Dim ticker As String

    Dim daily_change As Double
    daily_change = 0

    Dim pct_change As Double
    Dim close_price As Double

    Dim volume_total As Double
    volume_total = 0

    Dim summary_table_row As Integer
    summary_table_row = 2

    Dim sum_tab_r As Long
    
    Dim max_vol As Double
    

'...determine last row/column of data in worksheet

    lastrow = ActiveSheet.UsedRange.Rows.Count
    lastcolumn = ActiveSheet.UsedRange.Columns.Count

'...begin loop condition to calculate for EACH ticker seperately

    For I = 2 To lastrow

'...conditional to determine when ticker row changes

'... if ticker changes then
        If (Cells(I + 1, 1).Value <> Cells(I, 1).Value) Then
        
            Range("I1").Value = "Ticker"
            ticker = Cells(I, 1).Value
            Range("I" & summary_table_row).Value = ticker
                                                    
            Range("L1").Value = "Total Stock Volume"
            volume_total = volume_total + Cells(I, 7).Value
            Range("L" & summary_table_row).Value = volume_total
            
'...find length of summary table and then apply column formatting options
            sum_tab_r = Range("L" & Rows.Count).End(xlUp).Row
            Range("L2 : L" & sum_tab_r).ColumnWidth = 20
        
            Range("J1").Value = "Yearly Change"
            Range("J" & summary_table_row).Value = daily_change
            
            Range("J2 : J" & sum_tab_r).ColumnWidth = 12
            Range("J2 : J" & sum_tab_r).NumberFormat = "$  ###,##0.00"
                     
'...calculate annual percent change by summing daily pct change for each ticker & formatting results in summary table
            Range("K1").Value = "Percent Change"
            close_price = Cells(I, 6).Value
            pct_change = daily_change / close_price
            Range("K" & summary_table_row).Value = pct_change

            Range("K2 : K" & sum_tab_r).ColumnWidth = 14
            Range("K2 : K" & sum_tab_r).NumberFormat = "0.00%"
            
'...challenge section results summary
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest %Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("O1 : O5").ColumnWidth = 20
            
        
'...reset variables when change of ticker occurs

            summary_table_row = summary_table_row + 1
            volume_total = 0
            daily_change = 0
        
        Else
        
'...calculate following while ticker is same value
            volume_total = volume_total + Cells(I, 7).Value
            daily_change = daily_change + (Cells(I + 1, 6).Value - Cells(I, 6).Value)
        
        End If
    
    Next I
    
    'max_vol = Application.WorksheetFunction.Max(Range("L2:L" & sum_tab_r))
     '       MsgBox (max_vol)

'...repeat code for next worksheet in workbook
        
Next ws


End Sub
