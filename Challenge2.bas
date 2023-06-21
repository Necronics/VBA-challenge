Attribute VB_Name = "Module1"
Sub challenge2():

'--------- Looping Through Sheets In The Workbook ---------'
Dim sheetCount As Integer
sheetCount = Application.Worksheets.Count
For k = 1 To sheetCount
Worksheets(k).Activate

'--------- Defining Variables ---------'
Dim ticker As String
Dim yearOpen As Double
Dim stockVolume As Variant
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Variant
Dim j As Integer

'--------- Resetting Variables For Each Sheet ---------'
ticker = ""
yearOpen = 0
greatestIncrease = 0
greatestDecrease = 0
greatestVolume = 0

'--------- Looping Through Original Table ---------'
j = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'--------- Recording And Outputting Data For A New Ticker ---------'
    If Cells(i, 1).Value <> ticker Then
        ticker = Cells(i, 1).Value
        Cells(j, 9).Value = ticker
        yearOpen = Cells(i, 3).Value
        stockVolume = Cells(i, 7).Value
        
'--------- Recording Data For Last Row Of A Ticker  ---------'
    ElseIf Cells(i + 1, 1).Value <> ticker Then
    
'--------- Outputting Yearly Change  ---------'
        Cells(j, 10).Value = Cells(i, 6).Value - yearOpen
        
'--------- Outputting Percent Change  ---------'
        Cells(j, 11).Value = Cells(j, 10).Value / yearOpen

'--------- Comparing and Outputting Greatest Percent Increase  ---------'
        If Cells(j, 11).Value > greatestIncrease Then
            greatestIncrease = Cells(j, 11).Value
            Cells(2, 17).Value = greatestIncrease
            Cells(2, 16).Value = Cells(j, 9).Value
        End If

'--------- Comparing and Outputting Greatest Percent Decrease  ---------'
        If Cells(j, 11).Value < greatestDecrease Then
            greatestDecrease = Cells(j, 11).Value
            Cells(3, 17).Value = greatestDecrease
            Cells(3, 16).Value = Cells(j, 9).Value
        End If

'--------- Comparing and Outputting Total Volume and Greatest Total Volume  ---------'
        Cells(j, 12) = stockVolume + Cells(i, 7).Value
        If Cells(j, 12).Value > greatestVolume Then
            greatestVolume = Cells(j, 12).Value
            Cells(4, 17).Value = greatestVolume
            Cells(4, 16).Value = Cells(j, 9).Value
        End If
        
'--------- Incrementing New Table Row Upon Completion Of A Ticker  ---------'
        j = j + 1
        
 '--------- Calculating Total Stock Volume  ---------'
    Else
        stockVolume = stockVolume + Cells(i, 7).Value
    End If
'--------- Incrementing Original Table Row  ---------'
    Next i
Next
End Sub
