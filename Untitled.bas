Attribute VB_Name = "Module1"
Sub alpha_test():

Dim headers() As Variant
Dim mainws As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

'set header

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O1").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


For Each mainws In wb.Sheets
    With mainws
    .Rows(1).Value = " "
    For i = LBound(headers()) To UBound(headers())
    .Cells(1 + i).Value = headers(i)
    
        Next i
         .Rows(1).Font.Bold = True
         .Rows(1).VerticalAlignment = xlCenter
         End With
        
        Next mainws
        
        For Each mainws In Worksheets
        
        Dim ticker_name As String
        ticker_name = " "
        Dim total_ticker_volume As Double
        total_ticker_volume = 0
        Dim beg_price As Double
        beg_price = 0
        Dim end_price As Double
        end_price = 0
        Dim yearly_price_change As Double
        yearly_price_change = 0
        Dim yearly_price_change_percent As Double
        yearly_price_change_percent = 0
        Dim max_ticker_name As String
        max_ticker_name = " "
        Dim min_ticker_name As String
        min_ticker_name = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_volume_ticker_name As String
        max_volume_ticker_name = " "
        Dim max_volume As Double
        max_volume = 0
        
        
        'set location for variables
        
        Dim summary_table_row As Long
        summary_table_row = 2
        
        'set row count
        Dim lastrow As Long
        
        lastrow = mainws.Cells(Rows.Count, 1).End(xlUp).Row
        
        beg_price = mainws.Cells(2, 3).Value
        
        For i = 2 To lastrow
        
                If mainws.Cells(i + 1, 1).Value <> mainws.Cells(i, 1).Value Then
                
                ticker_name = mainws.Cells(i, 1).Value
                
                
                end_price = mainws.Cells(i, 6).Value
                yearly_price_change = end_price - beg_price
                
                If beg_price <> 0 Then
                    yearly_price_change_percent = (yearly_price_change / beg_price) * 100
                    
                    End If
                    
                    total_ticker_volume = total_ticker_volume + mainws.Cells(i, 7).Value
                    
                    mainws.Range("I" & summary_table_row).Value = ticker_name
                    
                    mainws.Range("J" & summary_table_row).Value = yearly_price_change
                    
                
                'colour fill
                If (yearly_price_change > 0) Then
                mainws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                
                ElseIf (yearly_price_change <= 0) Then
                
                        mainws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                        
                        End If
                        
                        'print yearly price chang
                    mainws.Range("K" & summary_table_row).Value = (CStr(yearly_price_change_percent) & "%")
                    
                    'print total stock
                    
                    mainws.Range("L" & summary_table_row).Value = total_ticker_volume
                    
                    'add 1
                    summary_table_row = summary_table_row + 1
                    
                    'get next beg price
                    beg_price = main.wscells(i + 1, 3).Value
                    
                    'calculations
                    
                    If (yearly_price_change_percent > max_percent) Then
                    max_percent = yearly_price_change_percent
                    max_ticker_name = ticker_name
                    
                    End If
                    
                    If (total_ticker_volume > max_volume) Then
                    max_volume = total_ticker_volume
                    max_volume_ticker_name = ticker_name
                    End If
                    
                    'reset value
                    yearly_price_change_percent = 0
                    total_ticker_volume = 0
                    
                    Else
                    
                    total_ticker_volume = total_ticker_volume + mainws.Cells(i, 7).Value
                    
                    End If
                    
                    Next i
                    
                    'print values in cell
                    mainws.Range("Q2").Value = (CStr(max_percent) & "%")
                    mainws.Range("Q3").Value = (CStr(min_percent) & "%")
                    mainws.Range("P2").Value = max_ticker_name
                    mainws.Range("P3").Value = min_ticker_name
                    mainws.Range("Q4").Value = max_volume
                    mainws.Range("O2").Value = "greatest % increase"
                    mainws.Range("O3").Value = "greatest % decrease"
                    mainws.Range("O4").Value = "greatest total volume"
                    
                    
                    
                    Next mainws
                    
               
                    
                    
                
        
End Sub



