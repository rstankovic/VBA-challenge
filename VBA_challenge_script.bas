Attribute VB_Name = "Module1"
Sub LetsRun()
    Dim b As Integer
    Dim isEmpty As String
    Dim currTicker As String
    Dim lastTicker As String
    Dim percChange As Double
    Dim yChange As Long
    Dim yStart As Long
    Dim yEnd As Long
    Dim totalVol As Long
    Dim tableB As Integer 'this is to manage the creation of the new table
    tableB = 2
    b = 2
    isEmpty = Cells(b, 1).Value
    lastTicker = Cells(b - 1, 1).Value
    yStart = Cells(b, 3).Value
    
    While isEmpty <> ""
        currTicker = Cells(b, 1).Value
        lastTicker = Cells(b - 1, 1).Value
        If lastTicker = "<ticker>" Then
            'skip because first row is just for headers
            'we don't need to worry about yStart because we established it outside of the while loop
        ElseIf currTicker <> lastTicker Then
            'these next few lines finish up the range and begin calculations
            yEnd = Cells(b - 1, 6).Value 'establish end of last ticker range
            percChange = yStart / yEnd - 1  'calculate percent change
            yChange = yStart - yEnd 'calculate actual change in price
            'the next four lines write a new table off to the right
            Cells(tableB, 9).Value = lastTicker '
            Cells(tableB, 10).Value = yChange
            'this if statement is to change color format of yChange column
            If yChange < 0 Then
                Cells(tableB, 10).Interior.Color = vbRed
            Else
                Cells(tableB, 10).Color = vbGreen
            End If
            Cells(tableB, 12).Value = totalVol
            Cells(tableB, 11).Value = percChange 'this comes before reset of yStart so that we can still perform calculations
            'next few lines reset values because we have hit a new ticker
            totalVol = Cells(b, 7).Value
            yStart = Cells(b, 3).Value
            totalVol = Cells(b, 7).Value 'new ticker new sum to calculate
            tableB = tableB + 1 'new ticker, new row
            'no need to set currTicker here because it gets set every iteration of the while loop
        Else
            totalVol = totalVol + Cells(b, 7).Value
        End If

        b = b + 1
        isEmpty = Cells(b, 1).Value 'this while loop will close before setting currTicker as empty string
    Wend
    
End Sub
