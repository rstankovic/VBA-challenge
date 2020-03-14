Attribute VB_Name = "Module1"
Sub LetsRun()
    'sub that will run through the active worksheet
    Dim b As Double
    Dim isEmpty As String
    Dim currTicker As String
    Dim lastTicker As String
    Dim percChange As Double
    Dim yChange As Double
    Dim yStart As Double
    Dim yEnd As Double
    Dim totalVol As Variant
    Dim tableB As Integer 'this is to manage the creation of the new table
    tableB = 2
    b = 3
    isEmpty = Cells(b, 1).Value
    yStart = Cells(b - 1, 3).Value
    
    While isEmpty <> ""
        currTicker = Cells(b, 1).Value
        lastTicker = Cells(b - 1, 1).Value
        If lastTicker = "<ticker>" Then
            'skip because first row is just for headers
            'we don't need to worry about yStart because we established it outside of the while loop
        ElseIf currTicker <> lastTicker Then
            'these next few lines finish up the range and begin calculations
            yEnd = Cells(b - 1, 6).Value 'establish end of last ticker range
            yChange = yEnd - yStart 'calculate actual change in price
            'the next four lines write a new table off to the right
            Cells(tableB, 9).Value = lastTicker '
            Cells(tableB, 10).Value = yChange
            'this if statement is to change color format of yChange column
            If yChange < 0 Then
                Cells(tableB, 10).Interior.Color = vbRed
            Else
                Cells(tableB, 10).Interior.Color = vbGreen
            End If
            Cells(tableB, 12).Value = totalVol
            If yStart <> 0 Then 'to avoid times when we might divide by zero
                Cells(tableB, 11).Value = Format(yEnd / yStart - 1)
            Else
                Cells(tableB, 11).Value = "N/A"
            End If
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
    b = 2
    Dim currMax As Double
    Dim currMin As Double
    Dim maxStock As String
    Dim minStock
    currMax = 0
    currMin = 0
    isEmpty = Cells(b, 9).Value
    While isEmpty <> ""
        If Cells(b, 11).Value <> "N/A" Then 'for divide by zero cases
            If currMax < CDbl(Cells(b, 11).Value) Then
                currMax = Cells(b, 11).Value
                maxStock = Cells(b, 9).Value
            End If
            If currMin > CDbl(Cells(b, 11).Value) Then
                currMin = Cells(b, 11).Value
                minStock = Cells(b, 9).Value
            End If
        End If
        b = b + 1
        isEmpty = Cells(b, 9).Value
        
    Wend
        
    Cells(1, 13).Value = "Stock with greatest positive change is " & maxStock
    Cells(2, 13).Value = Format(currMax, "Percent")
    Cells(1, 14).Value = "Stock with greatest negative change is " & minStock
    Cells(2, 14).Value = Format(currMin, "Percent")
    
End Sub

Sub Umbrella()
    'sub that will run previous sub through all worksheets
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ' do whatever you need
        LetsRun
        ws.Cells(1, 1) = 1 'this sets cell A1 of each sheet to "1"
    Next
End Sub
