Attribute VB_Name = "SkipStopUse"
Sub SkipStopUse()

Dim RngCurrentCell As Range
Dim RngSelected As Range
Dim lCurrentRow As Long
Dim lCurrentFeatureName As String

Set RngSelected = Application.Selection
lCurrentRow = 0

For Each RngCurrentCell In RngSelected.Cells
    
    If RngCurrentCell.Row <> lCurrentRow Then
        lCurrentRow = RngCurrentCell.Row
        lCurrentFeatureName = Cells(lCurrentRow, 2).Value
        
        'Check user selected a good range
        If lCurrentRow < 8 Then
            MsgBox "Your selection must be lower than row 7"
            End
        End If
        
        'If the row contains a name for the feature in column 2, shuffle through the prefixes
        If Len(lCurrentFeatureName) >= 4 Then '(only checking up to 4 digits because some features are only 4 characters long - even though the prefix is 5-characters long)
        
            'If prefixed with "SKIP_", switch to "STOP_"
            If Left(lCurrentFeatureName, 4) = "SKIP" Then
                lCurrentFeatureName = "STOP_" & Right(lCurrentFeatureName, Len(lCurrentFeatureName) - 5)
                Cells(lCurrentRow, 2).Font.Color = -16776961
                Cells(lCurrentRow, 2).Font.Bold = True
            'Else if prefixed with "STOP_", switch to unprefixed
            ElseIf Left(lCurrentFeatureName, 4) = "STOP" Then
                lCurrentFeatureName = Right(lCurrentFeatureName, Len(lCurrentFeatureName) - 5)
                Cells(lCurrentRow, 2).Font.Color = -16777216
                Cells(lCurrentRow, 2).Font.Bold = False
             'else if not yet prefixed, add "SKIP_"
            Else
                lCurrentFeatureName = "SKIP_" & lCurrentFeatureName
                Cells(lCurrentRow, 2).Font.Color = -6279056
                Cells(lCurrentRow, 2).Font.Bold = True
            End If
            
            Cells(lCurrentRow, 2).Value = lCurrentFeatureName
            
        Else
            'There isn't a long-enough name written in column 2. Ignore this row.
        End If
    End If
    
Next RngCurrentCell


End Sub



