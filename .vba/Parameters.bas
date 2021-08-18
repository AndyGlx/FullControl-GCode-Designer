Attribute VB_Name = "Parameters"
Sub SetParameterCellNames()

Application.DisplayAlerts = False

'First delete current names
Dim NmCellName As Name
For Each NmCellName In ActiveWorkbook.Names
    If Not Left(NmCellName.Name, 5) = "_xlfn" Then
        NmCellName.Delete
    End If
Next NmCellName

'Then assign new names
Range("S37:T1000").Select
Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:=False

Worksheets("Main Sheet").Range("T15").Name = "BedTemp"
Worksheets("Main Sheet").Range("T16").Name = "NozzleFeedratePrinting"
Worksheets("Main Sheet").Range("T17").Name = "NozzleFeedrateTravelling"
Worksheets("Main Sheet").Range("T18").Name = "NozzleTemp"
Worksheets("Main Sheet").Range("T26").Name = "FanSpeed"

Range("A6").Select

Application.DisplayAlerts = True

End Sub
