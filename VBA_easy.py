Sub easy()

' Loop through Worksheet
For Each ws In Worksheets
Dim WorksheetName As String

    ' Define Variables
    Dim TickerSymbol As String
    Dim TickerTotal As Double

    ' Sum of ticker total volume should start at 0
    TickerTotal = 0

    ' Keep track of the location of each TickerSymbol in the Summary Table
    Dim Summary_TickerSymbol As Integer
    Summary_TickerSymbol = 2

    ' Determine the Last Row of the Data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   

    ' Loop through all Ticker Data
    For i = 2 To LastRow

        ' Check if we are still withing the same Ticker Symbol, if we are not ...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Then
        TickerTotal = TickerTotal + Cells(i, 7).Value
        TickerSymbol = Cells(i, 1).Value

        ' Print TickerSymbol in the Summary Table
        Range("I" & Summary_TickerSymbol).Value = TickerSymbol

        ' Print TickerTotal into the Summary Table
        Range("J" & Summary_TickerSymbol).Value = TickerTotal

        ' Add one to the Summary_TickerSymbol Table
        Summary_TickerSymbol = Summary_TickerSymbol + 1

        ' Resent TickerTotal
        TickerTotal = 0

        ' If the cell mmediately following a row is the same brand...
        Else
        ' Add to the TickerTotal
        TickerTotal = TickerTotal + Cells(i, 7).Value

        End If
        
        Next i
    
        ' Add Ticker and Total Stock Volume Headers to Summary Table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Total Stock Volume"
        
Next ws

End Sub

