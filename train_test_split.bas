Attribute VB_Name = "Module2"
Sub Train_test_split()
    Dim Data_range As Range
    Dim Total_rows As Long
    Dim Training_rows As Long
    Dim Test_rows As Long
    Dim i As Long
    Dim Random_number As Double
    Dim Target_col As Long
    
    'set data range
    Set Data_range = Range("A1").CurrentRegion
    
    'amount of row in a range
    Total_rows = Data_range.Rows.Count - 1
    
    'create train set (80%)
    Training_rows = Application.RoundUp(Total_rows * 0.8, 0)
    
    'create test set (20%)
    Test_rows = Total_rows - Training_rows
    
    'find first empty column on the right side
    Target_col = Data_range.Columns.Count + 1
    Do While Application.CountA(Data_range.Resize(, Target_col).Columns(Target_col)) > 0
        Target_col = Target_col + 1
    Loop
    
    'random set test or training in the last right column
    For i = 2 To Total_rows + 1
        Random_number = Rnd()
        If Random_number < 0.8 And Training_rows > 0 Then
            Cells(i, Target_col).Value = "Train"
            Training_rows = Training_rows - 1
        Else
            Cells(i, Target_col).Value = "Test"
            Test_rows = Test_rows - 1
        End If
    Next i
End Sub
