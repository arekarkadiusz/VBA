Attribute VB_Name = "Module2"
Sub Train_test_split()
    Dim DataRange As Range
    Dim TotalRows As Long
    Dim TrainingRows As Long
    Dim TestRows As Long
    Dim i As Long
    Dim RandomNumber As Double
    Dim TargetCol As Long
    
    'set data range
    Set DataRange = Range("A1").CurrentRegion
    
    'amount of row in a range
    TotalRows = DataRange.Rows.Count - 1
    
    'create train set (80%)
    TrainingRows = Application.RoundUp(TotalRows * 0.8, 0)
    
    'create test set (20%)
    TestRows = TotalRows - TrainingRows
    
    'find first empty column on the right side
    TargetCol = DataRange.Columns.Count + 1
    Do While Application.CountA(DataRange.Resize(, TargetCol).Columns(TargetCol)) > 0
        TargetCol = TargetCol + 1
    Loop
    
    'random set test or training in the last right column
    For i = 2 To TotalRows + 1
        RandomNumber = Rnd()
        If RandomNumber < 0.7 And TrainingRows > 0 Then
            Cells(i, TargetCol).Value = "Training"
            TrainingRows = TrainingRows - 1
        Else
            Cells(i, TargetCol).Value = "Test"
            TestRows = TestRows - 1
        End If
    Next i
End Sub
