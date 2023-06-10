Attribute VB_Name = "Module1"
Option Explicit

Sub SummaryMultipleStockYear()

Dim ws As Worksheet

    For Each ws In Worksheets

        'Declearing and naming variables
            Dim WorksheetName As String
            Dim cr As Long          'Current row
            Dim tb As Long          'Start row of ticker block
            Dim TickCount As Long   'Index counter to fill Ticker Row
            Dim LastRowA As Long    'Last row column A
            Dim LastRowI As Long    'Last row column I
            Dim PC As Double        'Variable for Percent Change Calculation
            Dim GRTIncr As Double   'Variable for Greatest Increase
            Dim GRTDecr As Double   'Variable for Greatest Decrease
            Dim GRTTOTVol As Double 'Variable for Greatest Total Volume
                                           
            WorksheetName = ws.Name 'Getting Worksheet Name
            
        'Headers For Data Report
                ws.[I1].Value = "Ticker"
                ws.[J1].Value = "Yearly Change"
                ws.[K1].Value = "Percent Change"
                ws.[L1].Value = "Total Stock Volume"
                ws.[O2].Value = "Greatest Percent Increase"
                ws.[O3].Value = "Greatest Percent Decrease"
                ws.[O4].Value = "Greatest Percent Total Volume"
                ws.[P1].Value = "Ticker"
                ws.[Q1].Value = "Value"
                
                
                
        'Formatting Headers
            ws.Activate
                Range("I:L").Select
            With Selection
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .Columns.AutoFit
            End With
                
                Range("O:Q").Select
            With Selection
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .Columns.AutoFit
            End With
            
            ws.Activate
                Columns("J:J").Select
                Selection.Style = "Currency"
            
            
        'Assigning Variables
            TickCount = 2                                   'Setting Ticker Counter to first row
            tb = 2                                          'Set Start to row 2
            LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row 'find the last non-blank cell in column A
           
            
        'For Loop all Rows
            For cr = 2 To LastRowA
                If ws.Cells(cr + 1, 1).Value <> ws.Cells(cr, 1).Value Then 'Check ticker name change
                
                ws.Cells(TickCount, 9).Value = ws.Cells(cr, 1).Value                            'Set ticker to column (I2)
                ws.Cells(TickCount, 10).Value = ws.Cells(cr, 6).Value - ws.Cells(tb, 3).Value 'Calculates Yearly Change in column (J2)
                
        'Condtional Formatting
            If ws.Cells(TickCount, 10).Value < 0 Then
                ws.Cells(TickCount, 10).Interior.ColorIndex = 3 'Cells Background to Red
            
            Else
                ws.Cells(TickCount, 10).Interior.ColorIndex = 4 'Cells Background to Green
            End If
            
        'Calculates percent change in column K
            If ws.Cells(tb, 3).Value <> 0 Then
                PC = ((ws.Cells(cr, 6).Value - ws.Cells(tb, 3).Value) / ws.Cells(tb, 3).Value)
            
        'Formatting PC
                ws.Cells(TickCount, 11).Value = Format(PC, "Percent")
            Else
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
            End If
            
        'Calculates total column in column L
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(tb, 7), ws.Cells(cr, 7)))
            
        'Increasing TickCount by 1
                TickCount = TickCount + 1
        
        'New Start row for ticker block
                tb = cr + 1
            End If
        Next cr
        
        'Last non-blank column I
                LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
             
                
                
        'Summary
                GRTTOTVol = ws.[L2].Value
                GRTIncr = ws.[K2].Value
                GRTDecr = ws.[K2].Value
            
        'Loop Summary
            For cr = 2 To LastRowI
                If ws.Cells(cr, 12).Value > GRTTOTVol Then
                GRTTOTVol = ws.Cells(cr, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(cr, 9).Value
                
            Else
                GRTTOTVol = GRTTOTVol
                
                
            End If
                If ws.Cells(cr, 11).Value > GRTIncr Then
                GRTIncr = ws.Cells(cr, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(cr, 9).Value
                
            Else
                GRTIncr = GRTIncr
            End If
            
            
            If ws.Cells(cr, 11).Value < GRTDecr Then
                GRTDecr = ws.Cells(cr, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(cr, 9).Value
            Else
                GRTDecr = GRTDecr
            End If
            
        'Summary for results
                ws.Cells(2, 17).Value = Format(GRTIncr, "Percent")
                ws.Cells(3, 17).Value = Format(GRTDecr, "Percent")
                ws.Cells(4, 17).Value = Format(GRTTOTVol, "Scientific")
                
            Next cr
            
    Next ws
    
         

End Sub
