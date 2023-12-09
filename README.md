![image](https://github.com/Zoouri/VBA-Challenge/assets/151431300/2324e58e-34e3-4060-8843-aad0b62eb8bf)# VBA-Challenge
Homework Week 2 - VBA

Sub Multiple_Stock_Data()
        For Each ws In Worksheets
        
 ' Keep track of the location
        Dim Summary_Stock As Integer
        Summary_Stock_Row = 2
  
' Set an initial variable for holding the Ticker
        Dim Total_Stock As Double
        Dim Brand_Name As String
        
'Setting Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
' Set an initial variable for holding the Vol
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
      Summary_Stock_Row = Summary_Stock_Row + 1
      
      Total_Stock = 0
    Else
    
     Total_Stock = Total_Stock + Cells(i, 7).Value
    
    End If
    
  Next i
  Next ws
  
End Sub
