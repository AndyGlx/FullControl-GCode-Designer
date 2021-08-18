Attribute VB_Name = "SaveLoadDesign"
Sub SaveDesign()

Dim lLowestCell As Long
Dim lLowestCellTemp As Long
Dim strRangeToCopy As String
Dim wSaveSheet As Worksheet
Dim strSheetID As String
Dim strDesignID As String
Dim arrDesignIDs As Variant



lLowestCell = 1

Sheets("Main Sheet").Activate
strSheetID = Range("D6").Value
strDesignID = Range("B6").Value
Set wSaveSheet = Sheets(strSheetID)


'Check whether the design ID is unique
wSaveSheet.Activate
lLowestCellTemp = Cells(1000000, 44).End(xlUp).Row
arrDesignIDs = Range(Cells(1, 44), Cells(lLowestCellTemp, 44)).Value
'Only do this check if the sheet already has at least one design saved in it
If lLowestCellTemp > 1 Then
    For i = 1 To lLowestCellTemp
        If arrDesignIDs(i, 1) = strDesignID Then
            Sheets("Main Sheet").Activate
            MsgBox "WARNING" & vbNewLine & "WARNING" & vbNewLine & "WARNING" & vbNewLine & "WARNING" & vbNewLine & "WARNING" & vbNewLine & "There is already a design recorded in the """ & strSheetID & """ worksheet with the design ID """ & strDesignID & """, please choose a unique Design ID and try again."
            End
        End If
    Next i
End If

'Find the range of data to copy
Sheets("Main Sheet").Activate
For i = 1 To 40
    lLowestCellTemp = Cells(1000000, i).End(xlUp).Row
    If lLowestCellTemp > lLowestCell Then
        lLowestCell = lLowestCellTemp
    End If
Next i
ststrRangeToCopy = Range(Cells(6, 1), Cells(lLowestCell, 40)).Address

'Find the row number to save the data in the save worksheet
wSaveSheet.Activate
lLowestCell = 1
For i = 1 To 40
    lLowestCellTemp = Cells(1000000, i).End(xlUp).Row
    If lLowestCellTemp > lLowestCell Then
        lLowestCell = lLowestCellTemp
    End If
Next i

'Copy the data across
Sheets("Main Sheet").Activate
Range(ststrRangeToCopy).Select
Selection.Copy
wSaveSheet.Activate
Cells(lLowestCell + 2, 1).Select
ActiveSheet.Paste

Cells(lLowestCell + 2, 41).Value = "Save data (DO NOT DELETE):"
Cells(lLowestCell + 2, 42).Value = "Save time: " & Now
For i = 1 To 4
    Cells(lLowestCell + 2, 42 + i).Value = Cells(lLowestCell + 2, i).Value
Next i

'Notify the user of a successful save (unless this is a temporary save to the recycle bin)
If Not strSheetID = "RecycleBin" Then
    MsgBox "The design called """ & strDesignID & """ has been successfully saved in the """ & strSheetID & """ worksheet."
End If

Sheets("Main Sheet").Activate

End Sub


Sub LoadDesign()

'First save the old design to the "RecycleBin" worksheet

Sheets("Main Sheet").Activate
strSheetID = "RecycleBinSave"
Range("A6:D6").Select
Selection.Copy
Range("G6").Select
ActiveSheet.Paste
Range("D6").Value = "RecycleBin"
Range("B6").Value = "RecycleBinSave at " & Now
Call SaveDesign
Sheets("Main Sheet").Activate


Dim load_lLowestCell As Long
Dim load_lLowestCellTemp As Long
Dim load_strRangeToCopy As String
Dim load_wSaveSheet As Worksheet
Dim load_strSheetID As String
Dim load_strDesignID As String
Dim load_arrDesignIDs As Variant

Dim load_topRow As Long
Dim load_bottomRow As Long

'Clear the old data
Sheets("Main Sheet").Activate
For i = 1 To 40
    load_lLowestCellTemp = Cells(1000000, i).End(xlUp).Row
    If load_lLowestCellTemp > load_lLowestCell Then
        load_lLowestCell = load_lLowestCellTemp
    End If
Next i
'Delete the existing design ONLY IF THERE IS A DESIGN THERE!
If load_lLowestCell > 6 Then
    Range(Cells(6, 1), Cells(load_lLowestCell, 40)).Delete
End If



'Ask the user which sheet to load from
Load LoadDesignForm
LoadDesignForm.Label1.Caption = "Choose a folder (worksheet):"
For i = 1 To ThisWorkbook.Sheets.Count
    If Not ThisWorkbook.Sheets(i).Name = "Main Sheet" _
    And Not ThisWorkbook.Sheets(i).Name = "FeatParams" _
    And Not ThisWorkbook.Sheets(i).Name = "Printpath" _
    And Not ThisWorkbook.Sheets(i).Name = "StartGCODE" _
    And Not ThisWorkbook.Sheets(i).Name = "EndGCODE" _
    And Not ThisWorkbook.Sheets(i).Name = "GCODE" _
    And Not ThisWorkbook.Sheets(i).Name = "ToolGCODE" _
    And Not ThisWorkbook.Sheets(i).Name = "RepFeatList" _
    Then
        LoadDesignForm.ListBox1.AddItem (ThisWorkbook.Sheets(i).Name)
    End If
Next i
LoadDesignForm.Show
load_strSheetID = LoadDesignForm.ListBox1.Value
Set load_wSaveSheet = Sheets(load_strSheetID)
Unload LoadDesignForm

'Ask the user which design to load
'Activate the sheet and find the lowest cell
load_wSaveSheet.Activate
load_lLowestCellTemp = Cells(1000000, 44).End(xlUp).Row
load_arrDesignIDs = Range(Cells(1, 44), Cells(load_lLowestCellTemp, 44)).Value
'Change the userform textand add all the designs to the userform list.
Load LoadDesignForm
LoadDesignForm.Label1.Caption = "Choose a design (worksheet):"
For i = 1 To load_lLowestCellTemp
    If load_arrDesignIDs(i, 1) <> "" Then
        LoadDesignForm.ListBox1.AddItem (load_arrDesignIDs(i, 1))
    End If
Next i
LoadDesignForm.Show
load_strDesignID = LoadDesignForm.ListBox1.Value
Unload LoadDesignForm

'load_strSheetID = "Sheet1" '''''''UPDATE THIS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'load_strDesignID = "Test4" '''''''UPDATE THIS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set load_wSaveSheet = Sheets(load_strSheetID)


'Find the range of data to copy
load_wSaveSheet.Activate
load_lLowestCellTemp = Cells(1000000, 44).End(xlUp).Row
load_arrDesignIDs = Range(Cells(1, 44), Cells(load_lLowestCellTemp, 44)).Value

'First, find the design and check if it's the last in the list.
For i = 1 To load_lLowestCellTemp
    If load_arrDesignIDs(i, 1) = load_strDesignID Then
        load_topRow = i
    End If
Next i
    
'If it's NOT the last design in the list, copy to the row above the design below.
If load_topRow <> load_lLowestCellTemp Then
    load_bottomRow = Cells(load_topRow, 44).End(xlDown).Row - 2

'If it IS the last design in the list, find the bottom row of sheet and copy to row below.
Else
    For i = 1 To 40
        load_lLowestCellTemp = Cells(1000000, i).End(xlUp).Row
        If load_lLowestCellTemp > load_lLowestCell Then
            load_lLowestCell = load_lLowestCellTemp
        End If
    Next i
    load_bottomRow = load_lLowestCell
End If



'Copy the data across
Range(Cells(load_topRow, 1), Cells(load_bottomRow, 40)).Select
Selection.Copy
Sheets("Main Sheet").Activate
Cells(6, 1).Select
ActiveSheet.Paste


Call SetParameterCellNames



End Sub


