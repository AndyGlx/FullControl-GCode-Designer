Attribute VB_Name = "AddFeature"

Sub AddFeature()

Dim arrFeatureTypes As Variant
Dim strFeatureName As String
Dim lFeatureRow As Long

lFeatureRow = Selection.Row

If lFeatureRow < 8 Then

    MsgBox "Choose a row greater than row 7"
    End

End If

If Not WorksheetFunction.CountA(Range(Cells(lFeatureRow, 2), Cells(lFeatureRow, 16))) = 0 Then
    If (MsgBox("There is already data in row " & lFeatureRow & vbNewLine & "Do you want to continue? Click ""No"" to end the program", vbYesNo)) = 7 Then
        End
    End If
End If

'Ask the user which feature to add
arrFeatureTypes = Sheets("FeatParams").Range("B2:B42").Value
'Change the userform textand add all the designs to the userform list.
Load LoadDesignForm
LoadDesignForm.Caption = "New feature"
LoadDesignForm.Label1.Caption = "Choose a feature to add:"
For i = 1 To UBound(arrFeatureTypes, 1)
    If arrFeatureTypes(i, 1) <> "" Then
        LoadDesignForm.ListBox1.AddItem (arrFeatureTypes(i, 1))
    End If
Next i
LoadDesignForm.Show
strFeatureName = LoadDesignForm.ListBox1.Value
Unload LoadDesignForm

'Find the feature from the list at the top of the sheet.
For i = 1 To UBound(arrFeatureTypes, 1)
    If arrFeatureTypes(i, 1) = strFeatureName Then
        Sheets("FeatParams").Activate
        Cells(1 + i, 2).Select
    End If
Next i

Range(Selection, Selection.Offset(0, 14)).Select
Selection.Copy
Sheets("Main sheet").Activate

Cells(lFeatureRow, 2).Select

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

If Cells(lFeatureRow, 2).Value = "Line (cartesian)" Then
    Cells(lFeatureRow, 2).Value = "Line"
End If
If Cells(lFeatureRow, 2).Value = "Line (polar)" Then
    Cells(lFeatureRow, 2).Value = "Line"
End If
If Cells(lFeatureRow, 2).Value = "Reflect (XY)" Then
    Cells(lFeatureRow, 2).Value = "Reflect"
End If
If Cells(lFeatureRow, 2).Value = "Reflect (polar)" Then
    Cells(lFeatureRow, 2).Value = "Reflect"
End If
If Cells(lFeatureRow, 2).Value = "Reflect (Z)" Then
    Cells(lFeatureRow, 2).Value = "Reflect"
End If
If Cells(lFeatureRow, 2).Value = "Concentric repeat (only for ""rectangle"")" Then
    Cells(lFeatureRow, 2).Value = "Concentric repeat"
End If


If Cells(lFeatureRow, 2).Value = "Repeat rule" Then


    'Ask the user which feature to add
    arrFeatureTypes = Sheets("FeatParams").Range("G23:G38").Value
    'Change the userform textand add all the designs to the userform list.
    Load LoadDesignForm
    LoadDesignForm.Caption = "Repeat rule"
    LoadDesignForm.Label1.Caption = "Choose the type of ""Repeat rule"" to add:"
    For i = 1 To UBound(arrFeatureTypes, 1)
        If arrFeatureTypes(i, 1) <> "" Then
            LoadDesignForm.ListBox1.AddItem (arrFeatureTypes(i, 1))
        End If
    Next i
    LoadDesignForm.Show
    strFeatureName = LoadDesignForm.ListBox1.Value
    Unload LoadDesignForm
    
    'Find the feature from the list at the top of the sheet.
    For i = 1 To UBound(arrFeatureTypes, 1)
        If arrFeatureTypes(i, 1) = strFeatureName Then
            Sheets("FeatParams").Activate
            Cells(22 + i, 7).Select
        End If
    Next i
    
    Range(Selection, Selection.Offset(0, 9)).Select
    Selection.Copy
    Sheets("Main sheet").Activate
    
    Cells(lFeatureRow, 7).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

End If


If Cells(lFeatureRow, 2).Value = "Postprocess" Then


    'Ask the user which feature to add
    arrFeatureTypes = Sheets("FeatParams").Range("D44:D47").Value
    'Change the userform textand add all the designs to the userform list.
    Load LoadDesignForm
    LoadDesignForm.Caption = "Postprocess"
    LoadDesignForm.Label1.Caption = "Choose the type of ""Postprocess"" feature to add:"
    For i = 1 To UBound(arrFeatureTypes, 1)
        If arrFeatureTypes(i, 1) <> "" Then
            LoadDesignForm.ListBox1.AddItem (arrFeatureTypes(i, 1))
        End If
    Next i
    LoadDesignForm.Show
    strFeatureName = LoadDesignForm.ListBox1.Value
    Unload LoadDesignForm
    
    'Find the feature from the list at the top of the sheet.
    For i = 1 To UBound(arrFeatureTypes, 1)
        If arrFeatureTypes(i, 1) = strFeatureName Then
            Sheets("FeatParams").Activate
            Cells(43 + i, 4).Select
        End If
    Next i
    
    Range(Selection, Selection.Offset(0, 12)).Select
    Selection.Copy
    Sheets("Main sheet").Activate
    
    Cells(lFeatureRow, 4).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

End If

Rows(lFeatureRow).Select



End Sub
