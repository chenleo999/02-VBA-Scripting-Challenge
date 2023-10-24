Attribute VB_Name = "Module1"
Sub TickerSummary()

'get count of worksheets
Dim wscnt As Integer
wscnt = Worksheets.Count

'check if all sheets have same number of columns
For ws = 2 To wscnt
    If Worksheets(ws).Range("A1").End(xlToRight).Column _
    <> Worksheets(ws - 1).Range("A1").End(xlToRight).Column Then
        MsgBox "Sheets have different number of columns!"
        Exit For
        Exit Sub
    End If
Next ws

'check if all headers are same
Dim col, wscol As Integer
For ws = 2 To wscnt
    wscol = Worksheets(ws).Range("A1").End(xlToRight).Column
    For col = 1 To wscol
        If Worksheets(ws).Cells(1, col) _
        <> Worksheets(ws - 1).Cells(1, col) Then
            MsgBox "Sheets have different headers!"
            Exit For
            Exit Sub
        End If
    Next col
Next ws

'turn off screen update
Application.ScreenUpdating = False

'loop all sheets, run main process
Dim wsh As Worksheet
For Each wsh In Worksheets
    wsh.Select

    'add headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("M1") = "opendate"
    Range("N1") = "openprice"
    Range("O1") = "closedate"
    Range("P1") = "closeprice"
        
    'initiate summary1
    Cells(2, 9) = Cells(2, 1)       'ticker
    Cells(2, 12) = Cells(2, 7)      'total volume
    Cells(2, 13) = Cells(2, 2)      'open date
    Cells(2, 14) = Cells(2, 3)      'open price
    Cells(2, 15) = Cells(2, 2)      'close date
    Cells(2, 16) = Cells(2, 6)      'close price
    
    'loop and calculate summary1
    Dim row, sumrow, s As Integer, Got1 As Boolean
    row = 3         'start checking data at this row
    sumrow = 2      'last summary row
    Got1 = False    'ticker exists in summary
    
    While IsEmpty(Cells(row, 1)) = 0
        For s = 2 To sumrow
            'if match ticker, add total
            If Cells(row, 1) = Cells(s, 9) Then
                Got1 = True
                Cells(s, 12) = Cells(s, 12) + Cells(row, 7)
                'if date earlier then open, update open date and price
                If Cells(row, 2) < Cells(s, 13) Then
                    Cells(s, 13) = Cells(row, 2)
                    Cells(s, 14) = Cells(row, 3)
                'if date later then close, update close date and price
                ElseIf Cells(row, 2) > Cells(s, 15) Then
                    Cells(s, 15) = Cells(row, 2)
                    Cells(s, 16) = Cells(row, 6)
                End If
                Exit For
            End If
        Next s
        
        If Got1 = False Then 'find a new ticker, add a new summary row, and populate info
            sumrow = sumrow + 1
            Cells(sumrow, 9) = Cells(row, 1)
            Cells(sumrow, 12) = Cells(row, 7)
            Cells(sumrow, 13) = Cells(row, 2)
            Cells(sumrow, 14) = Cells(row, 3)
            Cells(sumrow, 15) = Cells(row, 2)
            Cells(sumrow, 16) = Cells(row, 6)
        Else
            Got1 = False
        End If
    
    row = row + 1
    Wend
    
    'finalize summary1
    s = 2
    While IsEmpty(Cells(s, 9)) = 0
        Cells(s, 10) = Cells(s, 16) - Cells(s, 14)
        If Cells(s, 10) > 0 Then
            Cells(s, 10).Interior.ColorIndex = 4
        ElseIf Cells(s, 10) < 0 Then
            Cells(s, 10).Interior.ColorIndex = 3
        End If
     
        Cells(s, 11) = (Cells(s, 16) - Cells(s, 14)) / Cells(s, 14)
        Cells(s, 11).Style = "Percent"
        Cells(s, 11).NumberFormat = "0.00%"
        
        s = s + 1
    Wend
    
    'clear helper columns
    Range("M:P").Delete (xlshitleft)
    
    'set summary2 headers
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Toal Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    'initiate summary2
    Range("P2") = Range("I2")
    Range("Q2") = Range("K2")
    Range("P3") = Range("I2")
    Range("Q3") = Range("K2")
    Range("P4") = Range("I2")
    Range("Q4") = Range("L2")
    
    'calculate summary2
    s = 3
    While IsEmpty(Cells(s, 9)) = 0
        Range("P2") = IIf(Cells(s, 11) > Range("Q2"), Cells(s, 9), Range("P2"))
        Range("Q2") = IIf(Cells(s, 11) > Range("Q2"), Cells(s, 11), Range("Q2"))
        Range("P3") = IIf(Cells(s, 11) < Range("Q3"), Cells(s, 9), Range("P3"))
        Range("Q3") = IIf(Cells(s, 11) < Range("Q3"), Cells(s, 11), Range("Q3"))
        Range("P4") = IIf(Cells(s, 12) > Range("Q4"), Cells(s, 9), Range("P4"))
        Range("Q4") = IIf(Cells(s, 12) > Range("Q4"), Cells(s, 12), Range("Q4"))
        s = s + 1
    Wend
    
    Range("Q2:Q3").Style = "Percent"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'final sheet format
    Range("I1:Q1").Font.Bold = 1
    Range("O2:O4").Font.Bold = 1
    Columns("A:Q").AutoFit

Next

'turn screen update back on
Application.ScreenUpdating = True

MsgBox "All set!!!"

End Sub


