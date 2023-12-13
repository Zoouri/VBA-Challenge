Attribute VB_Name = "Module1"
Sub Multiple_Stock_Data()
        For Each ws In Worksheets
        
 ' Keep track of the location
        Dim Summary_Stock As Integer
        Summary_Stock_Row = 2
  
' Set an initial variable for holding the Ticker
        Dim Total_Stock As Double
        Dim Brand_Name As String
        Dim Opening_Price As Double
        Dim Yearly_Change As Double
        Dim Closing_Price As Double
        Dim Percent_Change As Double
        
 'Setting Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
' Set an initial variable for holding the Vol/Price/Percentage
Opening_Price = Cells(2, 3).Value
Total_Stock = 0

'Inserting new column names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Looping the ticker
    For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Brand_Name = Cells(i, 1).Value
      Total_Stock = Total_Stock + Cells(i, 7).Value
      Range("I" & Summary_Stock_Row).Value = Brand_Name
      Range("L" & Summary_Stock_Row).Value = Total_Stock
      
      'Defining Closing price
      Closing_Price = Cells(i, 6).Value
      
      'Substracting opening price and closing price
      Yearly_Change = (Closing_Price - Opening_Price)
      
      'Write the change in stock in designated spot
      Range("J" & Summary_Stock_Row).Value = Yearly_Change
      
      'Percentages
      If Opening_Price = 0 Then
        Percent_Change = 0
    Else
        Percent_Change = Yearly_Change / Opening_Price
    End If
    'Write the percentage change in designated area
        Range("K" & Summary_Stock_Row).Value = Percent_Change
        
        'Change to Percent
        Range("K" & Summary_Stock_Row).NumberFormat = "0.00%"
        
        'Reset everything
      Summary_Stock_Row = Summary_Stock_Row + 1
      Total_Stock = 0
      Opening_Price = Cells(i + 1, 3)
      
    Else
    
     Total_Stock = Total_Stock + Cells(i, 7).Value
    
    End If

    Next i
    'Highlighting changes using conditional formatting
        LastRow_Summary_Table = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To LastRow_Summary_Table
  
        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 50
        Else
        Cells(i, 10).Interior.ColorIndex = 3
  End If
  Next i
  
  'BONUS
  'Inserting Column/Row names
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
  'Determining min and max values sourcing ticker name and max volume
  For i = 2 To LastRow_Summary_Table
  'Finding max percentage
   If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & LastRow_Summary_Table)) Then
    Cells(2, 16).Value = Cells(i, 9).Value
    Cells(2, 17).Value = Cells(i, 11).Value
    Cells(2, 17).NumberFormat = "0.00%"
    'Finding min percentage
    ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & LastRow_Summary_Table)) Then
    Cells(3, 16).Value = Cells(i, 9).Value
    Cells(3, 17).Value = Cells(i, 11).Value
    Cells(3, 17).NumberFormat = "0.00%"
    'Finding max vol
    ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & LastRow_Summary_Table)) Then
    Cells(4, 16).Value = Cells(i, 9).Value
    Cells(4, 17).Value = Cells(i, 12).Value
    
    End If
    
        Next i
  Next ws
  
End Sub

