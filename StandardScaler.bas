Attribute VB_Name = "Module1"
Sub StandardScaler()
    Dim inputRange As Range
    Dim outputRange As Range
    Dim cell As Range
    Dim i As Long
    Dim mean As Double
    Dim std As Double
    
    ' Define the input range
    Set inputRange = Application.InputBox("Define the input area", Type:=8)
    
    ' Define the output range
    Set outputRange = Application.InputBox("Select the cell in the upper left corner of the output range", Type:=8)
    
    ' Calculate the average value for each input column
    For i = 1 To inputRange.Columns.Count
        mean = WorksheetFunction.Average(inputRange.Columns(i))
        
        ' Calculate the standard deviation for each input column
        std = WorksheetFunction.StDev(inputRange.Columns(i))
        
        ' Scale each value in the column and store the result in the appropriate cell
        For Each cell In inputRange.Columns(i).Cells
            outputRange.Offset(cell.Row - inputRange.Row, i - 1).Value = (cell.Value - mean) / std
        Next cell
    Next i
    
    MsgBox "Data scaling is complete"
End Sub

