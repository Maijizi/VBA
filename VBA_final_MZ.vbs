Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim b As Integer

    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For b = 1 To WS_Count
    
        Dim i As Double
        Dim j As Double
        Set MyRg1 = ActiveWorkbook.Worksheets(b).Range("A:A")
        Set MyRg2 = ActiveWorkbook.Worksheets(b).Range("G:G")
        
        'print headers for all exercises
        ActiveWorkbook.Worksheets(b).Range("J1") = "Tickers"
        ActiveWorkbook.Worksheets(b).Range("K1") = "Yearly Change"
        ActiveWorkbook.Worksheets(b).Range("L1") = "Percent Change"
        ActiveWorkbook.Worksheets(b).Range("M1") = "Total Stock Volume"
        ActiveWorkbook.Worksheets(b).Range("P2") = "Greatest % Increase"
        ActiveWorkbook.Worksheets(b).Range("P3") = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(b).Range("P4") = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(b).Range("Q1") = "Ticker"
        ActiveWorkbook.Worksheets(b).Range("R1") = "Value"
        
        'Copy tickers to a new column
        ActiveWorkbook.Worksheets(b).Range("A:A").Copy _
        Destination:=ActiveWorkbook.Worksheets(b).Range("J1")
        
        'dedupe tikers
        ActiveWorkbook.Worksheets(b).Range("J:J").RemoveDuplicates Columns:=1, Header:=xlYes
           
        'loop, total stock volume
        
        j = ActiveWorkbook.Worksheets(b).Range("J" & Rows.Count).End(xlUp).row
        For i = 2 To j
            ActiveWorkbook.Worksheets(b).Cells(i, "M") = WorksheetFunction.SumIf(MyRg1, ActiveWorkbook.Worksheets(b).Cells(i, "J").Value, MyRg2)
        Next i
        
        
        'Start of second exercise, calculate change

        Dim k As Double
        Dim a As Double
    
        j = ActiveWorkbook.Worksheets(b).Range("J" & Rows.Count).End(xlUp).row
        a = 0
        For i = 2 To j
            'count rows of each stickers
            k = WorksheetFunction.CountIf(MyRg1, ActiveWorkbook.Worksheets(b).Cells(i, "J").Value)
        
            'calculate yearly change and color code change
            ActiveWorkbook.Worksheets(b).Cells(i, "K") = ActiveWorkbook.Worksheets(b).Cells(k + 1 + a, "F") - ActiveWorkbook.Worksheets(b).Cells(a + 2, "C")
            If ActiveWorkbook.Worksheets(b).Cells(i, "K") > 0 Then
                ActiveWorkbook.Worksheets(b).Cells(i, "K").Interior.Color = RGB(124, 252, 0)
            ElseIf ActiveWorkbook.Worksheets(b).Cells(i, "K") < 0 Then
                ActiveWorkbook.Worksheets(b).Cells(i, "K").Interior.Color = RGB(255, 0, 0)
            End If
               
            'PLNT in 2014 and 2015 have a lot of 0s in data. To resolve the issue, do a if check if the dominator is 0
            If ActiveWorkbook.Worksheets(b).Cells(a + 2, "C") <> 0 Then
                ActiveWorkbook.Worksheets(b).Cells(i, "L") = ActiveWorkbook.Worksheets(b).Cells(i, "K") / ActiveWorkbook.Worksheets(b).Cells(a + 2, "C")
            Else:
                ActiveWorkbook.Worksheets(b).Cells(i, "L") = 0
            End If
            a = a + k
        Next i

        
        'start of the third exercise
        
        Dim change As Double
        Dim location As String
        Dim row As String
        Set MyRg1 = ActiveWorkbook.Worksheets(b).Range("L:L")
        Set MyRg2 = ActiveWorkbook.Worksheets(b).Range("M:M")

        change = Application.WorksheetFunction.Max(MyRg1)
        location = MyRg1.Find(What:=change, LookIn:=xlFormulas).Address
        row = Split(location, "$")(2)
        ActiveWorkbook.Worksheets(b).Range("Q2") = ActiveWorkbook.Worksheets(b).Cells(Int(row), "J")
        ActiveWorkbook.Worksheets(b).Range("R2") = change

        change = Application.WorksheetFunction.Min(MyRg1)
        location = MyRg1.Find(What:=change, LookIn:=xlFormulas).Address
        row = Split(location, "$")(2)
        ActiveWorkbook.Worksheets(b).Range("Q3") = ActiveWorkbook.Worksheets(b).Cells(Int(row), "J")
        ActiveWorkbook.Worksheets(b).Range("R3") = change

        change = Application.WorksheetFunction.Max(MyRg2)
        location = MyRg2.Find(What:=change, LookIn:=xlFormulas).Address
        row = Split(location, "$")(2)
        ActiveWorkbook.Worksheets(b).Range("Q4") = ActiveWorkbook.Worksheets(b).Cells(Int(row), "J")
        ActiveWorkbook.Worksheets(b).Range("R4") = change
        
        ActiveWorkbook.Worksheets(b).Range("L:L").NumberFormat = "#.##%"
        ActiveWorkbook.Worksheets(b).Range("R2:R3").NumberFormat = "#.##%"

    Next b
    

End Sub