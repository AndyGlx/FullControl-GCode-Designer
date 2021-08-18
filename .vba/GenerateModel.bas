Attribute VB_Name = "GenerateModel"

Public Pi As Double
Dim cCommandType As Integer
'For cCommandType = "Print" or "Travel":
Public cX1 As Integer
Public cY1 As Integer
Public cZ1 As Integer
Public cX2 As Integer
Public cY2 As Integer
Public cZ2 As Integer
Public cW As Integer
Public cH As Integer
Public cE As Integer
Public cF As Integer
Public cT As Integer
'For cCommandType = "Retraction":
Public cRetractE As Integer
Public cRetractSpeed As Integer
Public cRetractZhop As Integer
Public cRetractZhopSpeed As Integer
'For all cCommandTypes:
Public cID As Integer
Public cIDtree As Integer
Public cNotes As Integer
Public cGCODE As Integer


Sub GenerateModel()

Application.ScreenUpdating = False

Dim bCheckErrors As Boolean
bCheckErrors = True

If bCheckErrors Then
    On Error GoTo ErrorHandlerPre
End If
Dim arrFeatures As Variant
Dim arrFeaturesTemp As Variant
Dim arrCommands As Variant
Dim arrTravelResponses As Variant
Dim arrFeatureList As Variant
Dim arrFeatureAdditonalParams As Variant
Dim arrCommandsBeingRepeated As Variant
Dim arrCommandsBeingModified As Variant
Dim arrRepeatRules As Variant
Dim arrCustomGCODEtemp As Variant
Dim arrStartGCODE As Variant
Dim arrEndGCODE As Variant
Dim arrToolChangeGCODE As Variant
Dim arrFullGCODE As Variant


Dim lCurrentCommand As Long
Dim lCurrentCommandFeature As Long
Dim lPreviousLineCommand As Long
Dim lAddedCommandsCounter As Long
Dim lCurrentRetractionFeature As Long
Dim lCurrentCustomGCODEFeature As Long
Dim lCurrentRectangleFeature As Long
Dim lCurrentCircleArcFeature As Long
Dim lCurrentPolygonFeature As Long
Dim lCurrentCartesianRepFeature As Long
Dim lCurrentReflectRepFeature As Long
Dim lCurrentPolarRepFeature As Long
Dim lCurrentLinearRepFeature As Long
Dim lCurrentCircRepFeature As Long
Dim lCurrentLayerRepFeature As Long
Dim lCurrentLineEqFeature As Long
Dim lCurrentLineEqPolarFeature As Long
Dim lCurrentConcentricRepeatFeature As Long
Dim lNumberOfCommandsBeingRepeated As Long
Dim lNumberOfCommandsBeingModified As Long
Dim lNumberOfRepeatRules As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim k1 As Long
Dim iParamVaried As Integer
Dim iOriginal As Long
Dim iRepeatRules As Integer
Dim iRepeatRulesFeatNumber_BEING_REPEATED As Integer
Dim iFeatureNumber As Integer
Dim iNumberOfFeatures As Integer
Dim iInputCols As Integer

Dim dblXoffset As Double
Dim dblYoffset As Double
Dim dblZoffset As Double
Dim dblLayer1EMultiplier As Double
Dim dblLayer1SpeedMultiplier As Double
Dim dblPrintSpeed As Double
Dim dblTravelSpeed As Double
Dim dblCurrentSpeed As Double

Dim strAutoRetractYesNo As String
Dim dblAutoRetractThreshold As Double
Dim dblAutoRetractE As Double
Dim dblAutoRetractEspeed As Double
Dim dblAutoUnretractE As Double
Dim dblAutoUnretractEspeed As Double
Dim dblAutoRetractZhop As Double
Dim dblAutoRetractZhopSpeed As Double


Dim dblXdisp As Double
Dim dblYdisp As Double
Dim dblZdisp As Double
Dim dblXcentre As Double
Dim dblYcentre As Double
Dim dblRotationAngle As Double
Dim dblLayerHeight As Double
Dim dblOffset As Double
Dim iDirection As Integer
Dim iNumberOfRepeats As Integer

Dim dblXreflect1 As Double
Dim dblYreflect1 As Double
Dim dblXreflect2 As Double
Dim dblYreflect2 As Double
Dim dblZreflect As Double

Dim strFeatIDrenumbered As String
Dim strFeatIDtree As String
Dim strInsideOutside As String

Dim strAdditionalParams As String

Dim strFeatureList As String

Dim dblXold1 As Double
Dim dblYold1 As Double
Dim dblZold1 As Double
Dim dblXold2 As Double
Dim dblYold2 As Double
Dim dblZold2 As Double
Dim dblXnew1 As Double
Dim dblYnew1 As Double
Dim dblZnew1 As Double
Dim dblXnew1TEMP As Double
Dim dblYnew1TEMP As Double
Dim dblZnew1TEMP As Double
Dim dblXnew2 As Double
Dim dblYnew2 As Double
Dim dblZnew2 As Double
Dim dblWidthOld As Double
Dim dblHeightOld As Double
Dim dblE As Double
Dim dblFspeed As Double
Dim iToolNumber As Integer
Dim strPrintTravelOld As String
Dim dblWidthNew As Double
Dim dblHeightNew As Double
Dim strPrintTravelNew As String

Dim dblRetractE As Double
Dim dblRetractSpeed As Double
Dim dblRetractZhop As Double
Dim dblRetractZhopSpeed As Double

Dim strCustomGCODE As String
Dim strToolChange_ID As String
Dim strStartGCODE_ID As String
Dim strEndGCODE_ID As String

Dim dblStartCornerX As Double
Dim dblStartCornerY As Double
Dim dblRecSizeX As Double
Dim dblRecSizeY As Double

Dim dblRadius As Double
Dim dblStartAngle As Double
Dim dblArcAngle As Double
Dim dblAngleSeg As Double
Dim lNumberOfSegs As Long

Dim dblRadius1 As Double
Dim dblRadius2 As Double
Dim dblAngle1 As Double
Dim dblAngle2 As Double

Dim strXequation As String
Dim strXequationTemp As String
Dim strYequation As String
Dim strYequationTemp As String
Dim strZequation As String
Dim strZequationTemp As String
Dim strAequation As String
Dim strAequationTemp As String
Dim strRequation As String
Dim strRequationTemp As String
Dim strWidthEquation As String
Dim strWidthEquationTemp As String
Dim strHeightEquation As String
Dim strHeightEquationTemp As String
Dim strEequation As String
Dim strEequationTemp As String
Dim strFspeedEquation As String
Dim strFspeedEquationTemp As String
Dim strRadiusEquation As String
Dim strRadiusEquationTemp As String
Dim strAngleEquation As String
Dim strAngleEquationTemp As String

Dim dblTstart As Double
Dim dblTend As Double
Dim dblTstep As Double
Dim dblT1 As Double
Dim dblT2 As Double


Dim dblLinearMathsOffsetX As Double
Dim dblLinearMathsOffsetY As Double
Dim dblLinearMathsOffsetZ As Double

Dim dblCurrentX As Double
Dim dblCurrentY As Double
Dim dblCurrentZ As Double
Dim iCurrentToolNumber As Integer
Dim dblInitialX As Double
Dim dblInitialY As Double
Dim dblInitialZ As Double
Dim iInitialTool As Integer

Dim dblPreviousX As Double
Dim dblPreviousY As Double
Dim dblPreviousZ As Double
Dim dblNextX As Double
Dim dblNextY As Double
Dim dblNextZ As Double

Dim dblStartTime As Double
Dim dblElapsedTime As Double
dblStartTime = Timer


Dim dblFeedstockFilamentDiameter As Double: dblFeedstockFilamentDiameter = Sheets("Main Sheet").Range("T19").Value
Dim strExtrusionUnits As String: strExtrusionUnits = Sheets("Main Sheet").Range("T20").Value

cCommandType = 1
'For cCommandType = "Print" or "Travel":
cX1 = 2
cY1 = 3
cZ1 = 4
cX2 = 5
cY2 = 6
cZ2 = 7
cW = 8
cH = 9
cE = 10
cF = 11
cT = 12
'For cCommandType = "Retraction":
cRetractE = 2
cRetractSpeed = 3
cRetractZhop = 2
cRetractZhopSpeed = 3
'For all cCommandTypes:
cID = 14
cIDtree = 15
cNotes = 16
cGCODE = 17


Progress_Box.Show vbModeless
DoEvents

Progress_Box.Label1.Caption = "Beginning program"
DoEvents

'VBA_edit_002 start
'CHECK FOR USE OF COMMA DECIMAL POINTS
If (InStr(1, CStr(9.9), ",") > 0) Then
    MsgBox "Your system appears to be using number formatting with "","" for decimal places instead of "".""" & vbNewLine & vbNewLine & "FullControl currently doesn't work with this formatting. Please change settings as described here: bit.ly/3dteSmc" & vbNewLine & vbNewLine & "(It's quick and easy, and you can change the settings back after using FullControl)"
End If
'VBA_edit_002 end


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''''''TAKE IN DATA FOR SETTINGS AND FEATURES
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@


Progress_Box.Label1.Caption = "Taking in data"
DoEvents

iInputCols = 14 'This is how many columns are read into VBA from the Excel spreadsheet

Sheets("Main Sheet").Activate
Range("B8").Select
If Range("B9").Value <> "" Then
    Range(Selection, Selection.End(xlDown)).Select
End If
'Range(Selection, Selection.Offset(0, (iInputCols + (iAdditionalParams - 1)) - 1)).Select
Range(Selection, Selection.Offset(0, iInputCols - 1)).Select
arrFeatures = Selection.Value

'To allow for a formula in column B giving a value of "", the iNumberOfFeatures is set so that the program ignores rows after the first blank one.
'This allows formulas to be used in Excel to dictate whether the features are included
'The reason I'm doing this to is allow the Progress box to indicate the correct number of features
iNumberOfFeatures = UBound(arrFeatures, 1)
For i = 1 To UBound(arrFeatures, 1)
    If arrFeatures(i, 1) = "" Then
        iNumberOfFeatures = i - 1
        Exit For
    End If
Next i

'Check for fomulas that currently return an error and change the array element to be the string of the formula instead of the (error) result
For i = 1 To iNumberOfFeatures
    For j = 1 To UBound(arrFeatures, 2)
        If IsError(arrFeatures(i, j)) Then
        arrFeatures(i, j) = Sheets("Main Sheet").Cells(7, 1).Offset(i, j).Formula
        End If
    Next j
Next i


''''''///////////////////////////////////////////
''''''ASSIST WITH VERSION CHANGES
'Assist with designs being translated from previous version of FullControl with different formats
For i = 1 To iNumberOfFeatures
    'I changed the format of "Line Equation" to be "Line equation" for consistency with other features' names.
    If arrFeatures(i, 1) = "Line Equation" Then
        arrFeatures(i, 1) = "Line equation"
    End If
    'I changed the way the user writes E and F override values. In the old format the cells identified below would be numeric
    Select Case arrFeatures(i, 1)
    Case "Line"
        Select Case arrFeatures(i, 2)
        Case "Cartesian"
            If arrFeatures(i, 12) <> "" And IsNumeric(arrFeatures(i, 12)) Then: MsgBox "Feature " & i & " has a numeric value for E/F/T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
        Case "Polar"
            If arrFeatures(i, 14) <> "" And IsNumeric(arrFeatures(i, 14)) Then: MsgBox "Feature " & i & " has a numeric value for E/F/T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
        End Select
    Case "Line equation"
        If arrFeatures(i, 12) <> "" And IsNumeric(arrFeatures(i, 12)) Then: MsgBox "Feature " & i & " has a numeric value for T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
    Case "Line equation polar"
        If arrFeatures(i, 14) <> "" And IsNumeric(arrFeatures(i, 14)) Then: MsgBox "Feature " & i & " has a numeric value for T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
    Case "Circle/arc"
        If arrFeatures(i, 12) <> "" And IsNumeric(arrFeatures(i, 12)) Then: MsgBox "Feature " & i & " has a numeric value for E/F/T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
    Case "Rectangle"
        If arrFeatures(i, 10) <> "" And IsNumeric(arrFeatures(i, 10)) Then: MsgBox "Feature " & i & " has a numeric value for E/F/T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
    Case "Polygon"
        If arrFeatures(i, 11) <> "" And IsNumeric(arrFeatures(i, 11)) Then: MsgBox "Feature " & i & " has a numeric value for E/F/T, which is not permitted and indicates the design was copied from an old version of FullControl - please correct this error": End
    End Select
    'I added a Z term to the repeat rule for linear offsets
    If arrFeatures(i, 1) = "Repeat rule" Then
        If arrFeatures(i, 6) = "OffsetLinear" Or arrFeatures(i, 6) = "OffsetLinearIncrement" Or arrFeatures(i, 6) = "OffsetLinearIncrementGraded" Then
            If arrFeatures(i, 9) = "YES" Or arrFeatures(i, 9) = "NO" Then: MsgBox "Feature " & i & " has a value of ""YES"" or ""NO"" for Zoffset, which indicates the design was copied from an old version of FullControl - please correct this error": End
        End If
    End If
    'I added autoretraction options and toolchange options
    If Sheets("Main Sheet").Range("S12").Value <> "ToolChange ID" Then: MsgBox "Cell ""S12"" on the ""Main Sheet"" does not have the value ""ToolChange ID"" which indicates you're using a format for the settings section from a previous version of FullControl - please change to use the latest format for settings": End
    If Sheets("Main Sheet").Range("S27").Value <> "AutoTravelRetraction?" Then: MsgBox "Cell ""S27"" on the ""Main Sheet"" does not have the value ""AutoTravelRetraction?"" which indicates you're using a format for the settings section from a previous version of FullControl - please change to use the latest format for settings": End
Next i
''''''///////////////////////////////////////////

dblXoffset = 0
dblYoffset = 0
dblZoffset = Range("T9").Value
dblLayer1EMultiplier = Range("T10").Value
dblLayer1SpeedMultiplier = Range("T11").Value
strToolChange_ID = Range("T12").Value
iInitialTool = Range("T13").Value
iCurrentToolNumber = iInitialTool
dblPrintSpeed = Range("T16").Value
dblTravelSpeed = Range("T17").Value
strStartGCODE_ID = Range("T21").Value
strEndGCODE_ID = Range("T22").Value
dblInitialX = Range("T23").Value
dblInitialY = Range("T24").Value
dblInitialZ = Range("T25").Value
dblCurrentX = dblInitialX
'VBA_edit_003 start
dblCurrentY = dblInitialY
dblCurrentZ = dblInitialZ
'VBA_edit_003 end
dblPreviousX = dblCurrentX
dblPreviousY = dblCurrentY
dblPreviousZ = dblCurrentZ


strAutoRetractYesNo = Range("T27").Value
dblAutoRetractThreshold = Range("T28").Value
dblAutoRetractE = Range("T29").Value
dblAutoRetractEspeed = Range("T30").Value
dblAutoUnretractE = Range("T31").Value
dblAutoUnretractEspeed = Range("T32").Value
dblAutoRetractZhop = Range("T33").Value
dblAutoRetractZhopSpeed = Range("T34").Value


Sheets("StartGCODE").Activate
For i = 1 To 1000
    If Range("A1").Offset(0, i).Value = strStartGCODE_ID Then
        ActiveSheet.Range(Cells(2, i + 1), Cells(10000, i + 1).End(xlUp)).Select
        arrStartGCODE = Selection.Value
        Exit For
    End If
Next i

Sheets("EndGCODE").Activate
For i = 1 To 1000
    If Range("A1").Offset(0, i).Value = strEndGCODE_ID Then
        ActiveSheet.Range(Cells(2, i + 1), Cells(10000, i + 1).End(xlUp)).Select
        arrEndGCODE = Selection.Value
        Exit For
    End If
Next i

Sheets("ToolGCODE").Activate
For i = 1 To 1000
    If Range("A1").Offset(0, i).Value = strToolChange_ID Then
        ActiveSheet.Range(Cells(3, i + 1), Cells(3, i + 1).Offset(22 * 4 - 1, 0)).Select
        arrToolChangeGCODE = Selection.Value
        Exit For
    End If
Next i


Sheets("Main Sheet").Activate
        
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''''''SETUP VARIABLES
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Pi = 3.14159265358979
ReDim arrCommands(1 To cGCODE, 1 To 1)
ReDim arrCustomGCODEtemp(1 To 10)
ReDim arrTravelResponses(1 To UBound(arrFeatures, 1), 1 To 1)
ReDim arrRepeatRules(1 To 13, 1 To 1)
'iNumberOfFeatures = UBound(arrFeatures, 1)
lCurrentCommand = 1
lCurrentCommandFeature = 1
lCurrentRetractionFeature = 1
lCurrentCustomGCODEFeature = 1
lCurrentRectangleFeature = 1
lCurrentCircleArcFeature = 1
lCurrentPolygonFeature = 1
lCurrentLinearRepFeature = 1
lCurrentCartesianRepFeature = 1
lCurrentReflectRepFeature = 1
lCurrentPolarRepFeature = 1
lCurrentCircRepFeature = 1
lCurrentLayerRepFeature = 1
lCurrentLineEqFeature = 1
lCurrentLineEqPolarFeature = 1
lCurrentConcentricRepeatFeature = 1
lNumberOfRepeatRules = 0
lNumberOfCommandsBeingRepeated = 0
lNumberOfCommandsBeingModified = 0



dblCurrentSpeed = 0



If bCheckErrors Then
    On Error GoTo ErrorHandlerRepRec
End If

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''''''FIND ALL "REPRODUCE AND RECALCULATE" FEATURES
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

ReDim arrFeaturesTemp(1 To 100000, 1 To iInputCols)

Dim bCheckIfReproduceFeatureUsed As Boolean
'Dim iNumberOfFeaturesTemp As Integer
Dim iNumberOfFeaturesBeingReproduced As Integer
'Dim iNumberOfFeaturesAdded As Integer
Dim iCurrentFeatureInNewLists As Integer
Dim iNumberOfFeaturesBeforeCurrentReproduceFeature As Integer
Dim arrOriginalFeaturesNumbers As Variant
Dim arrNewFeaturesNumbers As Variant
Dim arrFeatureReproduced As Variant
Dim arrReproducedFeaturesMappedToUserWrittenOnes As Variant

ReDim arrOriginalFeaturesNumbers(1 To 1)
ReDim arrFeatureReproduced(1 To 1)
ReDim arrNewFeaturesNumbers(1 To 1)
'iNumberOfFeaturesAdded = 0
iCurrentFeatureInNewLists = 1

bCheckIfReproduceFeatureUsed = False
For i = 1 To iNumberOfFeatures
    If arrFeatures(i, 1) = "Reproduce and recalculate" Then
        bCheckIfReproduceFeatureUsed = True
    End If
Next i

If bCheckIfReproduceFeatureUsed = True Then
    For i = 1 To iNumberOfFeatures
    
    
        If arrFeatures(i, 1) <> "Reproduce and recalculate" Then
    
            'Populate the two arrays that keep track of original and new feature numbers
            ReDim Preserve arrOriginalFeaturesNumbers(1 To iCurrentFeatureInNewLists)
            ReDim Preserve arrFeatureReproduced(1 To iCurrentFeatureInNewLists)
            ReDim Preserve arrNewFeaturesNumbers(1 To iCurrentFeatureInNewLists)
            arrOriginalFeaturesNumbers(iCurrentFeatureInNewLists) = i
            arrFeatureReproduced(iCurrentFeatureInNewLists) = iCurrentFeatureInNewLists
            arrNewFeaturesNumbers(iCurrentFeatureInNewLists) = iCurrentFeatureInNewLists
    
            'If this is not a "Reproduce and recalculate" feature, update the arrays of feature numbers (original versus new), copy the feature details to the temporary array of features, update the coordinates and update any list of user-written-feature-numbers.
            For j = 1 To iInputCols
                arrFeaturesTemp(iCurrentFeatureInNewLists, j) = arrFeatures(i, j)
            Next j
    
            'update any list of user-written-feature-numbers
            Call UpdateFeatNumbersOf_NOT_ReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, arrOriginalFeaturesNumbers)
    
            iCurrentFeatureInNewLists = iCurrentFeatureInNewLists + 1
    
    
        Else ' arrFeatures(i, 1) = "Reproduce and recalculate"
    
            'This if function will be removed once I'd added functions for polar offset as well as cartesian
            If arrFeatures(i, 7) <> "" Or arrFeatures(i, 8) <> "" Or arrFeatures(i, 9) <> "" Or arrFeatures(i, 10) <> "" Then
                MsgBox "Error for feature " & i & " - polar functions are not yet programmed"
                End
                'I don't think the X-centre or Y-centre terms are required. The radial and angular terms are simply added to radial or angular parameters. They are not used to generate X/Y displacements. They are set with the cartesian displacement parameters.
                'I'm not sure, would polar offsets affect cartesian-described features too?
            End If
    
            'Take in the data
            iNumberOfRepeats = arrFeatures(i, 3)
            dblXdisp = arrFeatures(i, 4)
            dblYdisp = arrFeatures(i, 5)
            dblZdisp = arrFeatures(i, 6)
    
            'Find which features to reproduce
            strFeatureList = arrFeatures(i, 2)
            Call ReplaceDashesInString(strFeatureList)
            arrFeaturesWrittenByUser = Split(strFeatureList, ",")
            iNumberOfFeaturesBeingReproduced = UBound(arrFeaturesWrittenByUser) + 1
    
            iNumberOfFeaturesBeforeCurrentReproduceFeature = iCurrentFeatureInNewLists - 1 ' the "-1" is because the "current" feature has not yet been written
    '        iNumberOfFeaturesBeforeCurrentReproduceFeature = i
    
            For iRep = 1 To iNumberOfRepeats
    
                iFirstFeatureForTheseRepeats = iCurrentFeatureInNewLists
                ReDim arrReproducedFeaturesMappedToUserWrittenOnes(0 To UBound(arrFeaturesWrittenByUser))
    
    
                'Run through each feature (in the new list of features - in case we have multiple "Reproduce and recalculate" features) and check if it is one of those included in the "Reproduce and recalculate" feature.
                For ii = 1 To iNumberOfFeaturesBeforeCurrentReproduceFeature
    
                    'See if the ii'th feature is supposed to be reproduced (written by the user?)
                    For k = 0 To iNumberOfFeaturesBeingReproduced - 1
                        'check if feature ii is being reproduced
                        If arrOriginalFeaturesNumbers(ii) = CInt(arrFeaturesWrittenByUser(k)) Then
    
                            'Reproduce this feature
                            'Populate the two arrays that keep track of original and new feature numbers
                            ReDim Preserve arrOriginalFeaturesNumbers(1 To iCurrentFeatureInNewLists)
                            ReDim Preserve arrFeatureReproduced(1 To iCurrentFeatureInNewLists)
                            ReDim Preserve arrNewFeaturesNumbers(1 To iCurrentFeatureInNewLists)
                            arrOriginalFeaturesNumbers(iCurrentFeatureInNewLists) = i
                            arrFeatureReproduced(iCurrentFeatureInNewLists) = ii
                            arrNewFeaturesNumbers(iCurrentFeatureInNewLists) = iCurrentFeatureInNewLists
    
                            'If this is not a "Reproduce and recalculate" feature, update the arrays of feature numbers (original versus new), copy the feature details to the temporary array of features, update the coordinates and update any list of user-written-feature-numbers.
                            For j = 1 To iInputCols
                                arrFeaturesTemp(iCurrentFeatureInNewLists, j) = arrFeaturesTemp(ii, j)
                            Next j
    
                            'update the coordinates
                            Call UpdateCoordsOfReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, dblXdisp, dblYdisp, dblZdisp, iRep)
    
                            'update any list of user-written-feature-numbers
                            'Set the feature number is this array so that the user-written-feature-numbers (e.g. in Cartesian repeat) can be updated to the correct numbers
                            arrReproducedFeaturesMappedToUserWrittenOnes(k) = iCurrentFeatureInNewLists
                            Call UpdateFeatNumbersOfReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, iFirstFeatureForTheseRepeats, arrFeatureReproduced, arrNewFeaturesNumbers, iRep)
    
    
    
    
    
    
    
                            iCurrentFeatureInNewLists = iCurrentFeatureInNewLists + 1
    
                        End If
    
                    Next k
    
                Next ii
            Next iRep
    
        End If
    Next i
    
    Sheets("RepFeatList").Activate
    ActiveSheet.Range("A2:T100000").ClearContents
    
    Dim arrFeatNumsToPaste As Variant
    ReDim arrFeatNumsToPaste(1 To iCurrentFeatureInNewLists - 1, 1 To 1)
    For i = 1 To iCurrentFeatureInNewLists - 1
        arrFeatNumsToPaste(i, 1) = arrOriginalFeaturesNumbers(i)
    Next i
    ActiveSheet.Range(Cells(2, 1), Cells(2, 1).Offset(UBound(arrFeatNumsToPaste, 1) - 1, 1)).Value = arrFeatNumsToPaste
    For i = 1 To iCurrentFeatureInNewLists - 1
        arrFeatNumsToPaste(i, 1) = arrFeatureReproduced(i)
    Next i
    ActiveSheet.Range(Cells(2, 2), Cells(2, 2).Offset(UBound(arrFeatNumsToPaste, 1) - 1, 1)).Value = arrFeatNumsToPaste
    For i = 1 To iCurrentFeatureInNewLists - 1
        arrFeatNumsToPaste(i, 1) = arrNewFeaturesNumbers(i)
    Next i
    ActiveSheet.Range(Cells(2, 3), Cells(2, 3).Offset(UBound(arrFeatNumsToPaste, 1) - 1, 1)).Value = arrFeatNumsToPaste
    
    ReDim arrFeatures(1 To iCurrentFeatureInNewLists - 1, 1 To iInputCols)
    For i = 1 To iCurrentFeatureInNewLists - 1
        For j = 1 To iInputCols
            arrFeatures(i, j) = arrFeaturesTemp(i, j)
        Next j
    Next i
    ActiveSheet.Range(Cells(2, 4), Cells(2, 4).Offset(UBound(arrFeatures, 1) - 1, (UBound(arrFeatures, 2) - 1))).Value = arrFeatures

End If


'The following code is directly copied from earlier in the program.
'To allow for a formula in column B giving a value of "", the iNumberOfFeatures is set so that the program ignores rows after the first blank one.
'This allows formulas to be used in Excel to dictate whether the features are included
'The reason I'm doing this to is allow the Progress box to indicate the correct number of features
iNumberOfFeatures = UBound(arrFeatures, 1)
For i = 1 To UBound(arrFeatures, 1)
    If arrFeatures(i, 1) = "" Then
        iNumberOfFeatures = i - 1
        Exit For
    End If
Next i
'Check for fomulas that currently return an error and change the array element to be the string of the formula instead of the (error) result
For i = 1 To iNumberOfFeatures
    For j = 1 To UBound(arrFeatures, 2)
        If IsError(arrFeatures(i, j)) Then
        arrFeatures(i, j) = Sheets("Main Sheet").Cells(7, 1).Offset(i, j).Formula
        End If
    Next j
Next i
Sheets("Main Sheet").Activate



If bCheckErrors Then
    On Error GoTo ErrorHandlerRepRule
End If

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''''''FIND ALL REPEAT RULES
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

For i = 1 To iNumberOfFeatures
    If arrFeatures(i, 1) = "Repeat rule" Then
        
        If arrFeatures(i, 6) = "OffsetLinearMaths" Then
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "xval", "Xval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "yval", "Yval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "zval", "Zval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "repval", "REPval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "xval", "Xval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "yval", "Yval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "zval", "Zval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "repval", "REPval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "xval", "Xval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "yval", "Yval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "zval", "Zval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "repval", "REPval")
        End If
        If arrFeatures(i, 6) = "OffsetPolarMaths" Then
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "xval", "Xval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "yval", "Yval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "zval", "Zval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "aval", "Aval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "rval", "Rval")
            arrFeatures(i, 9) = Replace(arrFeatures(i, 9), "repval", "REPval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "xval", "Xval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "yval", "Yval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "zval", "Zval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "aval", "Aval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "rval", "Rval")
            arrFeatures(i, 10) = Replace(arrFeatures(i, 10), "repval", "REPval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "xval", "Xval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "yval", "Yval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "zval", "Zval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "aval", "Aval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "rval", "Rval")
            arrFeatures(i, 11) = Replace(arrFeatures(i, 11), "repval", "REPval")
        End If
        
        lNumberOfRepeatRules = lNumberOfRepeatRules + 1
        ReDim Preserve arrRepeatRules(1 To 13, 1 To lNumberOfRepeatRules) 'Feature number (repeat feature), feature numbers (features being repeated), start layer, end layer, modifier type, modifier value
        arrRepeatRules(1, lNumberOfRepeatRules) = arrFeatures(i, 2) 'Feature number for the repeat feature
        arrRepeatRules(2, lNumberOfRepeatRules) = arrFeatures(i, 3) 'Feature numbers for the features being repeated
        arrRepeatRules(3, lNumberOfRepeatRules) = arrFeatures(i, 4) 'start layer
        arrRepeatRules(4, lNumberOfRepeatRules) = arrFeatures(i, 5) 'end layer
        arrRepeatRules(5, lNumberOfRepeatRules) = arrFeatures(i, 6) 'modifier name
        arrRepeatRules(6, lNumberOfRepeatRules) = arrFeatures(i, 7) 'modifier value
        arrRepeatRules(7, lNumberOfRepeatRules) = arrFeatures(i, 8) 'modifier value
        arrRepeatRules(8, lNumberOfRepeatRules) = arrFeatures(i, 9) 'modifier value
        arrRepeatRules(9, lNumberOfRepeatRules) = arrFeatures(i, 10) 'modifier value
        arrRepeatRules(10, lNumberOfRepeatRules) = arrFeatures(i, 11) 'modifier value
        arrRepeatRules(11, lNumberOfRepeatRules) = arrFeatures(i, 12) 'modifier value
        arrRepeatRules(12, lNumberOfRepeatRules) = arrFeatures(i, 13) 'modifier value
        arrRepeatRules(13, lNumberOfRepeatRules) = i 'the number actual feature number of this repeat rule (this comes at the end because I added it later in the development of this software.
        
    End If
Next i






If bCheckErrors Then
    On Error GoTo ErrorHandlerMain
End If

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''''''GO THROUGH EACH FEATURE
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

For i = 1 To iNumberOfFeatures
    
'    Application.StatusBar = "Working on feature " & i & " of " & iNumberOfFeatures
    
    dblElapsedTime = Timer - dblStartTime
    Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
    DoEvents
    
    Call ConvertRelativeToAbsoluteCoordintes(arrFeatures, i, dblCurrentX, dblCurrentY, dblCurrentZ)
    
    
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''LINE FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    If arrFeatures(i, 1) = "Line" Then
    
        If arrFeatures(i, 2) = "Cartesian" Then
            dblXnew1 = arrFeatures(i, 3)
            dblYnew1 = arrFeatures(i, 4)
            dblZnew1 = arrFeatures(i, 5)
            dblXnew2 = arrFeatures(i, 6)
            dblYnew2 = arrFeatures(i, 7)
            dblZnew2 = arrFeatures(i, 8)
            strPrintTravelNew = arrFeatures(i, 9)
            dblWidthNew = arrFeatures(i, 10)
            dblHeightNew = arrFeatures(i, 11)
            
            'Convert the list of additional parameters into individual values for dblE, dblFspeed and iToolNumber
            strAdditionalParams = arrFeatures(i, 12)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)
        
        ElseIf arrFeatures(i, 2) = "Polar" Then
        
            dblXcentre = arrFeatures(i, 3)
            dblYcentre = arrFeatures(i, 4)
            
            'Find the start of the line:
            dblRadius = arrFeatures(i, 5)
            dblStartAngle = arrFeatures(i, 6)
            'Temporarily set the old X Y to be at the point for angle=0
            dblXold = dblXcentre + dblRadius
            dblYold = dblYcentre
            'Find the real start position of the line (dblXnew1 and dblYnew1)
            Call RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblStartAngle * Pi / 180, dblXnew1, dblYnew1)
            
            'Find the end of the line:
            dblRadius = arrFeatures(i, 8)
            dblEndAngle = arrFeatures(i, 9)
            'Temporarily set the old X Y to be at the point for angle=0
            dblXold = dblXcentre + dblRadius
            dblYold = dblYcentre
            'Find the real end position of the line (dblXnew2 and dblYnew2)
            Call RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblEndAngle * Pi / 180, dblXnew2, dblYnew2)
            
            dblZnew1 = arrFeatures(i, 7)
            dblZnew2 = arrFeatures(i, 10)
            strPrintTravelNew = arrFeatures(i, 11)
            dblWidthNew = arrFeatures(i, 12)
            dblHeightNew = arrFeatures(i, 13)
            
            'Convert the list of additional parameters into individual values for dblE, dblFspeed and iToolNumber
            strAdditionalParams = arrFeatures(i, 14)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)
        
        End If
        
        
        strFeatIDrenumbered = "LineFeat" & lCurrentCommandFeature
        strFeatIDtree = i
        
        '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
        dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)

        
        lCurrentCommandFeature = lCurrentCommandFeature + 1
        
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''RETRACTION FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Retraction" Then
        
        'Do the material retraction bit first
        dblRetractE = -arrFeatures(i, 2)
        dblRetractSpeed = arrFeatures(i, 3)
        dblRetractZhop = arrFeatures(i, 4)
        dblRetractZhopSpeed = arrFeatures(i, 5)
        strFeatIDrenumbered = "RetractionFeat" & lCurrentRetractionFeature
        strFeatIDtree = i
        
        Call AddRetraction(dblRetractE, dblRetractSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand)
        'The do the Zhop bit if Zhop is bigger than 0
        If dblRetractZhop <> 0 Then
            Call AddRetractionZhop(dblRetractZhop, dblRetractZhopSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblCurrentZ)
        End If
        
        
        lCurrentRetractionFeature = lCurrentRetractionFeature + 1
        

    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''CUSTOM GCODE FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Custom GCODE" Then
        
        strFeatIDrenumbered = "CustomGCODEFeat" & lCurrentCustomGCODEFeature
        strFeatIDtree = i
        For j = 1 To UBound(arrCustomGCODEtemp)
            arrCustomGCODEtemp(j) = arrFeatures(i, j + 1)
        Next j
        
        Call AddCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, arrCustomGCODEtemp)
        lCurrentCustomGCODEFeature = lCurrentCustomGCODEFeature + 1
        

    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''RECTANGLE FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Rectangle" Then
    
        dblStartCornerX = arrFeatures(i, 2)
        dblStartCornerY = arrFeatures(i, 3)
        dblRecSizeX = arrFeatures(i, 4) - dblStartCornerX
        dblRecSizeY = arrFeatures(i, 5) - dblStartCornerY
        dblZnew1 = arrFeatures(i, 6)
        dblZnew2 = dblZnew1
        strPrintTravelNew = "Print"
        dblWidthNew = arrFeatures(i, 8)
        dblHeightNew = arrFeatures(i, 9)
        
        'Convert the list of additional parameters into individual values for dblE, dblFspeed and iToolNumber
        strAdditionalParams = arrFeatures(i, 10)
        Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)
        
        strFeatIDrenumbered = "RectangleFeat" & lCurrentRectangleFeature
        strFeatIDtree = i
        
        '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblStartCornerX, dblStartCornerY, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
        dblCurrentX = dblStartCornerX: dblCurrentY = dblStartCornerY: dblCurrentZ = dblZnew1
        
        If arrFeatures(i, 7) = "CW" And (dblRecSizeX >= 0 And dblRecSizeY >= 0) _
        Or arrFeatures(i, 7) = "CW" And (dblRecSizeX < 0 And dblRecSizeY < 0) _
        Or arrFeatures(i, 7) = "anti-CW" And (dblRecSizeX < 0 And dblRecSizeY >= 0) _
        Or arrFeatures(i, 7) = "anti-CW" And (dblRecSizeX >= 0 And dblRecSizeY < 0) _
        Then

            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY + dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX + dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY - dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX - dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
    
        Else
            
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX + dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY + dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX - dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY - dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            
        End If
        
        lCurrentRectangleFeature = lCurrentRectangleFeature + 1
        

    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''CIRCLE/ARC OR POLYGON FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Circle/arc" Or arrFeatures(i, 1) = "Polygon" Then
    
        dblXcentre = arrFeatures(i, 2)
        dblYcentre = arrFeatures(i, 3)
        dblZnew1 = arrFeatures(i, 4)
        dblZnew2 = dblZnew1
        dblRadius = arrFeatures(i, 5)
        dblStartAngle = arrFeatures(i, 6)
        
        If arrFeatures(i, 1) = "Circle/arc" Then
            If arrFeatures(i, 8) = "CW" Then
                dblArcAngle = -arrFeatures(i, 7)
            Else
                dblArcAngle = arrFeatures(i, 7)
            End If
            lNumberOfSegs = arrFeatures(i, 9)
            dblWidthNew = arrFeatures(i, 10)
            dblHeightNew = arrFeatures(i, 11)
            
            'Convert the list of additional parameters into individual values for dblE, dblFspeed and iToolNumber
            strAdditionalParams = arrFeatures(i, 12)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)
            
            strFeatIDrenumbered = "CircleArcFeat" & lCurrentCircleArcFeature
            strFeatIDtree = i
        Else
            If arrFeatures(i, 8 - 1) = "CW" Then
                dblArcAngle = -360
            Else
                dblArcAngle = 360
            End If
            lNumberOfSegs = arrFeatures(i, 9 - 1)
            dblWidthNew = arrFeatures(i, 10 - 1)
            dblHeightNew = arrFeatures(i, 11 - 1)
            
            'Convert the list of additional parameters into individual values for dblE, dblFspeed and iToolNumber
            strAdditionalParams = arrFeatures(i, 12 - 1)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)

            strFeatIDrenumbered = "PolygonFeat" & lCurrentPolygonFeature
            strFeatIDtree = i
            
        End If
        
        dblAngleSeg = dblArcAngle / lNumberOfSegs
        
        strPrintTravelNew = "Print"
        
        
        'Temporarily set the old X Y to be at the point for angle=0
        dblXold = dblXcentre + dblRadius
        dblYold = dblYcentre
        'Find the realy start position of the circle/arc (dblXnew1 and dblYnew1)
        Call RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblStartAngle * Pi / 180, dblXnew1, dblYnew1)
        'Find the X Y of the end of the first segment
        Call RotatePoint(dblXnew1, dblYnew1, dblXcentre, dblYcentre, dblAngleSeg * Pi / 180, dblXnew2, dblYnew2)
        'Add the first segment to the list of commands
        '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
        dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
        
        
        
        For j = 1 To lNumberOfSegs - 1
            
            dblXold1 = dblXnew1
            dblYold1 = dblYnew1
            dblXold2 = dblXnew2
            dblYold2 = dblYnew2
            Call RotateLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXcentre, dblYcentre, dblAngleSeg * Pi / 180, dblXnew1, dblYnew1, dblXnew2, dblYnew2)
            'Add the first segment to the list of commands
            '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
            dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
            
        Next j
        
        If arrFeatures(i, 1) = "Circle/arc" Then
            lCurrentCircleArcFeature = lCurrentCircleArcFeature + 1
        Else
            lCurrentPolygonFeature = lCurrentPolygonFeature + 1
        End If
    
    
    
    

    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''LINE EQUATION FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Line equation" Then
    
        strXequation = arrFeatures(i, 2)
        strYequation = arrFeatures(i, 3)
        strZequation = arrFeatures(i, 4)
        dblTstart = arrFeatures(i, 5)
        dblTend = arrFeatures(i, 6)
        lNumberOfSegs = arrFeatures(i, 7)
        
        
        dblTstep = (dblTend - dblTstart) / lNumberOfSegs
        
        strPrintTravelNew = "Print"
        
        If IsNumeric(arrFeatures(i, 8)) Then: dblWidthNew = arrFeatures(i, 8): strWidthEquation = "": Else strWidthEquation = arrFeatures(i, 8)
        If IsNumeric(arrFeatures(i, 9)) Then: dblHeightNew = arrFeatures(i, 9): strHeightEquation = "": Else strHeightEquation = arrFeatures(i, 9)
        If IsNumeric(arrFeatures(i, 10)) Then: dblE = arrFeatures(i, 10): strEequation = "": Else strEequation = arrFeatures(i, 10)
        If IsNumeric(arrFeatures(i, 11)) Then: dblFspeed = arrFeatures(i, 11): strFspeedEquation = "": Else strFspeedEquation = arrFeatures(i, 11)
        
        strFeatIDrenumbered = "LineEquationFeat" & lCurrentLineEqFeature
        strFeatIDtree = i
        
        'Replace the terms in equations with Uppercase first letters if they don't have them already
        strXequation = Replace(strXequation, "tval", "Tval")
        strXequation = Replace(strXequation, "zval", "Zval")
        strYequation = Replace(strYequation, "tval", "Tval")
        strYequation = Replace(strYequation, "zval", "Zval")
        strZequation = Replace(strZequation, "tval", "Tval")
        strZequation = Replace(strZequation, "zval", "Zval")
        strZequation = Replace(strZequation, "xval", "Xval")
        strZequation = Replace(strZequation, "yval", "Yval")
        
        For j = 1 To lNumberOfSegs
        'For each segment:
            
            If lNumberOfSegs > 5000 Then
                If j Mod 1000 = 0 Then
                    dblElapsedTime = Timer - dblStartTime
                    Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(line equation segment " & Format(j, "#,###") & " of " & Format(lNumberOfSegs, "#,###") & ")" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
                    DoEvents
                End If
            End If
        
            'Calculate the value of T at the start of this segment
            dblT1 = dblTstart + dblTstep * (j - 1)
            
            'Calculate the value of T at the end of this segment
            dblT2 = dblTstart + dblTstep * j
            
            strXequationTemp = Replace(strXequation, "Tval", CStr(dblT1))
            strXequationTemp = Replace(strXequationTemp, "Zval", CStr(dblCurrentZ))
            dblXnew1 = Evaluate(strXequationTemp)
            strXequationTemp = Replace(strXequation, "Tval", CStr(dblT2))
            strXequationTemp = Replace(strXequationTemp, "Zval", CStr(dblCurrentZ))
            dblXnew2 = Evaluate(strXequationTemp)
            strYequationTemp = Replace(strYequation, "Tval", CStr(dblT1))
            strYequationTemp = Replace(strYequationTemp, "Zval", CStr(dblCurrentZ))
            dblYnew1 = Evaluate(strYequationTemp)
            strYequationTemp = Replace(strYequation, "Tval", CStr(dblT2))
            strYequationTemp = Replace(strYequationTemp, "Zval", CStr(dblCurrentZ))
            dblYnew2 = Evaluate(strYequationTemp)
            strZequationTemp = Replace(strZequation, "Tval", CStr(dblT1))
            'VBA_edit_004 (replaced "Xavl" with "Xval"):
            strZequationTemp = Replace(strZequationTemp, "Xval", CStr(dblCurrentX))
            strZequationTemp = Replace(strZequationTemp, "Yval", CStr(dblCurrentY))
            strZequationTemp = Replace(strZequationTemp, "Zval", CStr(dblCurrentZ))
            dblZnew1 = Evaluate(strZequationTemp)
            strZequationTemp = Replace(strZequation, "Tval", CStr(dblT2))
            strZequationTemp = Replace(strZequationTemp, "Xval", CStr(dblCurrentX))
            strZequationTemp = Replace(strZequationTemp, "Yval", CStr(dblCurrentY))
            strZequationTemp = Replace(strZequationTemp, "Zval", CStr(dblCurrentZ))
            dblZnew2 = Evaluate(strZequationTemp)
            
            
            'solve the formulas for speed, extrusion, etc.
            'Replace "Xstart" terms with dblXnew1
            'Replace "Xend" terms with dblXnew2
            'Replace "Xmid" terms with (dblXnew1 + dblXnew2)/2
            'Similar for Y, Z and T
            If strWidthEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strWidthEquationTemp = strWidthEquation
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Xstart", CStr(dblXnew1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Xend", CStr(dblXnew2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Xmid", CStr((dblXnew1 + dblXnew2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Ystart", CStr(dblYnew1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Yend", CStr(dblYnew2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Ymid", CStr((dblYnew1 + dblYnew2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Zstart", CStr(dblZnew1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Zend", CStr(dblZnew2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Tstart", CStr(dblT1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Tend", CStr(dblT2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblWidthNew = Evaluate(strWidthEquationTemp)
            End If
            If strHeightEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strHeightEquationTemp = strHeightEquation
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Xstart", CStr(dblXnew1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Xend", CStr(dblXnew2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Xmid", CStr((dblXnew1 + dblXnew2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Ystart", CStr(dblYnew1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Yend", CStr(dblYnew2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Ymid", CStr((dblYnew1 + dblYnew2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Zstart", CStr(dblZnew1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Zend", CStr(dblZnew2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Tstart", CStr(dblT1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Tend", CStr(dblT2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblHeightNew = Evaluate(strHeightEquationTemp)
            End If
            If strEequation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strEequationTemp = strEequation
                strEequationTemp = Replace(strEequationTemp, "Xstart", CStr(dblXnew1)): strEequationTemp = Replace(strEequationTemp, "Xend", CStr(dblXnew2)): strEequationTemp = Replace(strEequationTemp, "Xmid", CStr((dblXnew1 + dblXnew2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Ystart", CStr(dblYnew1)): strEequationTemp = Replace(strEequationTemp, "Yend", CStr(dblYnew2)): strEequationTemp = Replace(strEequationTemp, "Ymid", CStr((dblYnew1 + dblYnew2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Zstart", CStr(dblZnew1)): strEequationTemp = Replace(strEequationTemp, "Zend", CStr(dblZnew2)): strEequationTemp = Replace(strEequationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Tstart", CStr(dblT1)): strEequationTemp = Replace(strEequationTemp, "Tend", CStr(dblT2)): strEequationTemp = Replace(strEequationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblE = Evaluate(strEequationTemp)
            End If
            If strFspeedEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strFspeedEquationTemp = strFspeedEquation
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Xstart", CStr(dblXnew1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Xend", CStr(dblXnew2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Xmid", CStr((dblXnew1 + dblXnew2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Ystart", CStr(dblYnew1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Yend", CStr(dblYnew2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Ymid", CStr((dblYnew1 + dblYnew2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zstart", CStr(dblZnew1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zend", CStr(dblZnew2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tstart", CStr(dblT1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tend", CStr(dblT2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblFspeed = Evaluate(strFspeedEquationTemp)
            End If
            
            'Convert the list of additional parameters into individual values. Note that dblE and dblFspeed are giving false parameters because they are definted above. iToolNumber is correct though.
            strAdditionalParams = arrFeatures(i, 12)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE_NOT_USED, dblFspeed_NOT_USED, iToolNumber)
            
            'Add travel and check for toolchange for the first line only
            If j = 1 Then
                '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
            End If
            
            'Add this line segment
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)

        Next j
        
        
        lCurrentLineEqFeature = lCurrentLineEqFeature + 1
        
    
    
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''LINE EQUATION POLAR FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    ElseIf arrFeatures(i, 1) = "Line equation polar" Then
        
        dblXcentre = arrFeatures(i, 2)
        dblYcentre = arrFeatures(i, 3)
            
        strAequation = arrFeatures(i, 4)
        strRequation = arrFeatures(i, 5)
        strZequation = arrFeatures(i, 6)
        dblTstart = arrFeatures(i, 7)
        dblTend = arrFeatures(i, 8)
        lNumberOfSegs = arrFeatures(i, 9)
        
        dblTstep = (dblTend - dblTstart) / lNumberOfSegs
        
        strPrintTravelNew = "Print"
        
        If IsNumeric(arrFeatures(i, 10)) Then: dblWidthNew = arrFeatures(i, 10): strWidthEquation = "": Else strWidthEquation = arrFeatures(i, 10)
        If IsNumeric(arrFeatures(i, 11)) Then: dblHeightNew = arrFeatures(i, 11): strHeightEquation = "": Else strHeightEquation = arrFeatures(i, 11)
        If IsNumeric(arrFeatures(i, 12)) Then: dblE = arrFeatures(i, 12): strEequation = "": Else strEequation = arrFeatures(i, 12)
        If IsNumeric(arrFeatures(i, 13)) Then: dblFspeed = arrFeatures(i, 13): strFspeedEquation = "": Else strFspeedEquation = arrFeatures(i, 13)
        
        strFeatIDrenumbered = "LineEquationPolarFeat" & lCurrentLineEqPolarFeature
        strFeatIDtree = i
        
        'Replace the terms in equations with Uppercase first letters if they don't have them already
        strAequation = Replace(strAequation, "tval", "Tval")
        strAequation = Replace(strAequation, "zval", "Zval")
        strRequation = Replace(strRequation, "tval", "Tval")
        strRequation = Replace(strRequation, "zval", "Zval")
        strZequation = Replace(strZequation, "tval", "Tval")
        strZequation = Replace(strZequation, "aval", "Aval")
        strZequation = Replace(strZequation, "rval", "Rval")
        strZequation = Replace(strZequation, "zval", "Zval")
        
        For j = 1 To lNumberOfSegs
        'For each segment:
                    
            If lNumberOfSegs > 5000 Then
                If j Mod 1000 = 0 Then
                    dblElapsedTime = Timer - dblStartTime
                    Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(line equation polar segment " & Format(j, "#,###") & " of " & Format(lNumberOfSegs, "#,###") & ")" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
                    DoEvents
                End If
            End If
        
            'Calculate the value of T at the start of this segment
            dblT1 = dblTstart + dblTstep * (j - 1)
            
            'Calculate the value of T at the end of this segment
            dblT2 = dblTstart + dblTstep * j
            
            
            strAequationTemp = Replace(strAequation, "Tval", CStr(dblT1))
            strAequationTemp = Replace(strAequationTemp, "Zval", CStr(dblCurrentZ))
            dblAngle1 = Evaluate(strAequationTemp)
            strAequationTemp = Replace(strAequation, "Tval", CStr(dblT2))
            strAequationTemp = Replace(strAequationTemp, "Zval", CStr(dblCurrentZ))
            dblAngle2 = Evaluate(strAequationTemp)
            strRequationTemp = Replace(strRequation, "Tval", Format(CDbl(dblT1), "#####0.0#####"))
            strRequationTemp = Replace(strRequationTemp, "Zval", Format(CDbl(dblCurrentZ), "#####0.0#####"))
'            strRequationTemp = Replace(strRequation, "Tval", CStr(dblT1))
'            strRequationTemp = Replace(strRequationTemp, "Zval", CStr(dblCurrentZ))
            dblRadius1 = Evaluate(strRequationTemp)
            strRequationTemp = Replace(strRequation, "Tval", Format(CDbl(dblT2), "#####0.0#####"))
            strRequationTemp = Replace(strRequationTemp, "Zval", Format(CDbl(dblCurrentZ), "#####0.0#####"))
'            strRequationTemp = Replace(strRequation, "Tval", CStr(dblT2))
'            strRequationTemp = Replace(strRequationTemp, "Zval", CStr(dblCurrentZ))
            dblRadius2 = Evaluate(strRequationTemp)
            strZequationTemp = Replace(strZequation, "Tval", CStr(dblT1))
            strZequationTemp = Replace(strZequationTemp, "Aval", CStr(dblAngle1))
            strZequationTemp = Replace(strZequationTemp, "Rval", CStr(dblRadius1))
            strZequationTemp = Replace(strZequationTemp, "Zval", CStr(dblCurrentZ))
            dblZnew1 = Evaluate(strZequationTemp)
            strZequationTemp = Replace(strZequation, "Tval", CStr(dblT2))
            strZequationTemp = Replace(strZequationTemp, "Aval", CStr(dblAngle2))
            strZequationTemp = Replace(strZequationTemp, "Rval", CStr(dblRadius2))
            strZequationTemp = Replace(strZequationTemp, "Zval", CStr(dblCurrentZ))
            dblZnew2 = Evaluate(strZequationTemp)
            
            
            'solve the formulas for speed, extrusion, etc.
            'Replace "Xstart" terms with dblXnew1
            'Replace "Xend" terms with dblXnew2
            'Replace "Xmid" terms with (dblXnew1 + dblXnew2)/2
            'Similar for Y, Z and T
            If strWidthEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strWidthEquationTemp = strWidthEquation
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Astart", CStr(dblAngle1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Aend", CStr(dblAngle2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Amid", CStr((dblAngle1 + dblAngle2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Rstart", CStr(dblRadius1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Yend", CStr(dblRadius2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Rmid", CStr((dblRadius1 + dblRadius2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Zstart", CStr(dblZnew1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Zend", CStr(dblZnew2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strWidthEquationTemp = Replace(strWidthEquationTemp, "Tstart", CStr(dblT1)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Tend", CStr(dblT2)): strWidthEquationTemp = Replace(strWidthEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblWidthNew = Evaluate(strWidthEquationTemp)
            End If
            If strHeightEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strHeightEquationTemp = strHeightEquation
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Astart", CStr(dblAngle1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Aend", CStr(dblAngle2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Amid", CStr((dblAngle1 + dblAngle2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Rstart", CStr(dblRadius1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Yend", CStr(dblRadius2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Rmid", CStr((dblRadius1 + dblRadius2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Zstart", CStr(dblZnew1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Zend", CStr(dblZnew2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strHeightEquationTemp = Replace(strHeightEquationTemp, "Tstart", CStr(dblT1)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Tend", CStr(dblT2)): strHeightEquationTemp = Replace(strHeightEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblHeightNew = Evaluate(strHeightEquationTemp)
            End If
            If strEequation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strEequationTemp = strEequation
                strEequationTemp = Replace(strEequationTemp, "Astart", CStr(dblAngle1)): strEequationTemp = Replace(strEequationTemp, "Aend", CStr(dblAngle2)): strEequationTemp = Replace(strEequationTemp, "Amid", CStr((dblAngle1 + dblAngle2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Rstart", CStr(dblRadius1)): strEequationTemp = Replace(strEequationTemp, "Yend", CStr(dblRadius2)): strEequationTemp = Replace(strEequationTemp, "Rmid", CStr((dblRadius1 + dblRadius2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Zstart", CStr(dblZnew1)): strEequationTemp = Replace(strEequationTemp, "Zend", CStr(dblZnew2)): strEequationTemp = Replace(strEequationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strEequationTemp = Replace(strEequationTemp, "Tstart", CStr(dblT1)): strEequationTemp = Replace(strEequationTemp, "Tend", CStr(dblT2)): strEequationTemp = Replace(strEequationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblE = Evaluate(strEequationTemp)
            End If
            If strFspeedEquation <> "" Then  'This condition is only satisfied if the variable was populated earlier in this section (see the IsNumeric check)
                strFspeedEquationTemp = strFspeedEquation
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Astart", CStr(dblAngle1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Aend", CStr(dblAngle2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Amid", CStr((dblAngle1 + dblAngle2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Rstart", CStr(dblRadius1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Yend", CStr(dblRadius2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Rmid", CStr((dblRadius1 + dblRadius2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zstart", CStr(dblZnew1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zend", CStr(dblZnew2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Zmid", CStr((dblZnew1 + dblZnew2) / 2))
                strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tstart", CStr(dblT1)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tend", CStr(dblT2)): strFspeedEquationTemp = Replace(strFspeedEquationTemp, "Tmid", CStr((dblT1 + dblT2) / 2))
                dblFspeed = Evaluate(strFspeedEquationTemp)
            End If
            
            
            'Convert the list of additional parameters into individual values. Note that dblE and dblFspeed are giving false parameters because they are definted above. iToolNumber is correct though.
            strAdditionalParams = arrFeatures(i, 14)
            Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE_NOT_USED, dblFspeed_NOT_USED, iToolNumber)
            
            'Find the start of the line:
            'Temporarily set the old X Y to be at the point for angle=0
            dblXold = dblXcentre + dblRadius1
            dblYold = dblYcentre
            'Find the real start position of the line (dblXnew1 and dblYnew1)
            Call RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblAngle1, dblXnew1, dblYnew1)
            'Why not just use sin and cos with dblAngle1 and dblRadius1, added to x/y centre?
            
            'Find the end of the line:
            'Temporarily set the old X Y to be at the point for angle=0
            dblXold = dblXcentre + dblRadius2
            dblYold = dblYcentre
            'Find the real end position of the line (dblXnew2 and dblYnew2)
            Call RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblAngle2, dblXnew2, dblYnew2)
            
            
            'Add travel and check for toolchange for the first line only
            If j = 1 Then
                '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
            End If
            
            'Add this line segment
            Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)

        Next j
        
        
        lCurrentLineEqPolarFeature = lCurrentLineEqPolarFeature + 1
        
        
        
        
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''CARTESIAN OR POLAR REPEAT FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    
    ElseIf arrFeatures(i, 1) = "Cartesian repeat" Or arrFeatures(i, 1) = "Polar repeat" Then
    
        
        If arrFeatures(i, 1) = "Cartesian repeat" Then
            'Find X displacement, Y displacement and number of repeats
            dblXdisp = arrFeatures(i, 3)
            dblYdisp = arrFeatures(i, 4)
            dblZdisp = arrFeatures(i, 5)
            iNumberOfRepeats = arrFeatures(i, 6)
            
        ElseIf arrFeatures(i, 1) = "Polar repeat" Then
            'Find centre point, angle and direction
            dblXcentre = arrFeatures(i, 3)
            dblYcentre = arrFeatures(i, 4)
            dblRotationAngle = arrFeatures(i, 5)
            dblRadialDisplacement = arrFeatures(i, 6)
            iNumberOfRepeats = arrFeatures(i, 7)
        End If
        
        
        'Find which commands/lines to repeat
        strFeatureList = arrFeatures(i, 2)
        For k = 1 To UBound(arrCommands, 2)
            If CheckIfCurrentCommandSatisfiesInclusionCriteria(arrCommands, k, strFeatureList, CStr(arrCommands(cIDtree, k))) = True Then
                lNumberOfCommandsBeingRepeated = lNumberOfCommandsBeingRepeated + 1
                ReDim Preserve arrCommandsBeingRepeated(1 To lNumberOfCommandsBeingRepeated)
                arrCommandsBeingRepeated(lNumberOfCommandsBeingRepeated) = k
            End If
        Next k
                              
        
        For j = 1 To iNumberOfRepeats
'            Application.StatusBar = "Working on feature " & i & " of " & iNumberOfFeatures & " ... repeat " & j & " of " & iNumberOfRepeats
            
            dblElapsedTime = Timer - dblStartTime
            Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(repeat " & j & " of " & iNumberOfRepeats & ")" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
            DoEvents
    
            For k = 1 To lNumberOfCommandsBeingRepeated
                
                
                If lNumberOfCommandsBeingRepeated > 5000 Then
                    If k Mod 1000 = 0 Then
                        dblElapsedTime = Timer - dblStartTime
                        Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(repeat " & j & " of " & iNumberOfRepeats & " - command " & Format(k, "#,###") & " of " & Format(lNumberOfCommandsBeingRepeated, "#,###") & ")" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
                        DoEvents
                    End If
                End If
            
                    
                lCurrentCommandBeingRepeated = arrCommandsBeingRepeated(k)
                                        
                If arrFeatures(i, 1) = "Cartesian repeat" Then
                    strFeatIDrenumbered = "CartesianRepFeat" & lCurrentCartesianRepFeature & " of " & arrCommands(cNotes, lCurrentCommandBeingRepeated)
                    strFeatIDtree = i & "." & j & "-" & arrCommands(cIDtree, lCurrentCommandBeingRepeated)
                Else
                    strFeatIDrenumbered = "PolarRepFeat" & lCurrentPolarRepFeature & " of " & arrCommands(cNotes, lCurrentCommandBeingRepeated)
                    strFeatIDtree = i & "." & j & "-" & arrCommands(cIDtree, lCurrentCommandBeingRepeated)
                End If
            
                If arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Retraction" Then
                
                    dblRetractE = arrCommands(cRetractE, lCurrentCommandBeingRepeated)
                    dblRetractSpeed = arrCommands(cRetractSpeed, lCurrentCommandBeingRepeated)
                
                ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "RetractionZhop" Then
                    dblRetractZhop = arrCommands(cRetractZhop, lCurrentCommandBeingRepeated)
                    dblRetractZhopSpeed = arrCommands(cRetractZhopSpeed, lCurrentCommandBeingRepeated)
                
                ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Custom GCODE" Then
                
                    strCustomGCODE = arrCommands(cGCODE, lCurrentCommandBeingRepeated)
                    For k1 = 1 To UBound(arrCustomGCODEtemp)
                        arrCustomGCODEtemp(k1) = arrCommands(k1 + 1, lCurrentCommandBeingRepeated)
                    Next k1
                    
                Else
                                        
                    'Find the X and Y coords for the line being repeated
                    dblXold1 = arrCommands(cX1, lCurrentCommandBeingRepeated)
                    dblYold1 = arrCommands(cY1, lCurrentCommandBeingRepeated)
                    dblZold1 = arrCommands(cZ1, lCurrentCommandBeingRepeated)
                    dblXold2 = arrCommands(cX2, lCurrentCommandBeingRepeated)
                    dblYold2 = arrCommands(cY2, lCurrentCommandBeingRepeated)
                    dblZold2 = arrCommands(cZ2, lCurrentCommandBeingRepeated)
                    strPrintTravelNew = arrCommands(cCommandType, lCurrentCommandBeingRepeated)
                    dblWidthNew = arrCommands(cW, lCurrentCommandBeingRepeated)
                    dblHeightNew = arrCommands(cH, lCurrentCommandBeingRepeated)
                    dblE = arrCommands(cE, lCurrentCommandBeingRepeated)
                    dblFspeed = arrCommands(cF, lCurrentCommandBeingRepeated)
                    iToolNumber = arrCommands(cT, lCurrentCommandBeingRepeated)
                    
                    
                    If arrFeatures(i, 1) = "Cartesian repeat" Then
                    
                        dblXnew1 = dblXold1 + dblXdisp * j
                        dblYnew1 = dblYold1 + dblYdisp * j
                        dblZnew1 = dblZold1 + dblZdisp * j
                        dblXnew2 = dblXold2 + dblXdisp * j
                        dblYnew2 = dblYold2 + dblYdisp * j
                        dblZnew2 = dblZold2 + dblZdisp * j
                        
                    
                    ElseIf arrFeatures(i, 1) = "Polar repeat" Then
                    
                        'Rotate the line
                        Call RotateLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXcentre, dblYcentre, dblRotationAngle * j * Pi / 180, dblXnew1, dblYnew1, dblXnew2, dblYnew2)
                        Call RadiallyDisplaceLine(dblXnew1, dblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement * j, dblXnew1, dblYnew1, dblXnew2, dblYnew2)
                        
                        dblZnew1 = dblZold1
                        dblZnew2 = dblZold2
                    
                    End If
                End If
                
                    
                
                'Check for repeat rules
                
                'If there are repeat rules
                If lNumberOfRepeatRules > 0 Then
                    'For each repeat rule
                    For iRepeatRules = 1 To lNumberOfRepeatRules
                        'Find which repeat-features the rule applies to (the features that are doing the repeating, not being repeated)
                        strFeatureList = arrRepeatRules(1, iRepeatRules)
                        Call ReplaceDashesInString(strFeatureList)
                        arrFeatureListForRepeatRule_DOING_REPEATING_ARRAY = Split(strFeatureList, ",")
                        'For each repeat-feature (the feature doing the repeating, not being repeated) affected by the current repeat rule
                        For iRepeatRulesRepeatFeatNumber_DOING_REPEATING = 0 To UBound(arrFeatureListForRepeatRule_DOING_REPEATING_ARRAY)
                            'If the current repeat rule applies to this repeat-feature (the feature doing the repeating, not being repeated)
                            If i = CInt(arrFeatureListForRepeatRule_DOING_REPEATING_ARRAY(iRepeatRulesRepeatFeatNumber_DOING_REPEATING)) Then
                                'If the current repeat rule applies to this repeat number
                                If j >= arrRepeatRules(3, iRepeatRules) And j <= arrRepeatRules(4, iRepeatRules) Then
                                    'Find which feature numbers the current repeat rule appli1es to
                                    strFeatureList = arrRepeatRules(2, iRepeatRules)
                                    
                                    If CheckIfCurrentCommandSatisfiesInclusionCriteria(arrCommands, lCurrentCommandBeingRepeated, strFeatureList, strFeatIDtree) = True Then
                                
                                        'Do the modification
                                        
                                        If arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Retraction" _
                                        Or arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "RetractionZhop" _
                                        Then
                                            
                                            If arrRepeatRules(5, iRepeatRules) = "GenericParameter" Then
                                                If arrRepeatRules(6, iRepeatRules) = 2 Then
                                                    'Minus sign because a positive number means retraction (negative E)
                                                    dblRetractE = -arrRepeatRules(7, iRepeatRules)
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 3 Then
                                                    dblRetractSpeed = arrRepeatRules(7, iRepeatRules)
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 4 Then
                                                    dblRetractZhop = arrRepeatRules(7, iRepeatRules)
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 5 Then
                                                    dblRetractZhopSpeed = arrRepeatRules(7, iRepeatRules)
                                                End If
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "GenericParameterIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                If arrRepeatRules(6, iRepeatRules) = 2 Then
                                                    'Minus sign because a positive number means retraction (negative E)
                                                    dblRetractE = dblRetractE - arrRepeatRules(7, iRepeatRules) * LayerCounterForCurrentRule
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 3 Then
                                                    dblRetractSpeed = dblRetractSpeed + arrRepeatRules(7, iRepeatRules) * LayerCounterForCurrentRule
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 4 Then
                                                    dblRetractZhop = dblRetractZhop + arrRepeatRules(7, iRepeatRules) * LayerCounterForCurrentRule
                                                ElseIf arrRepeatRules(6, iRepeatRules) = 5 Then
                                                    dblRetractZhopSpeed = dblRetractZhopSpeed + arrRepeatRules(7, iRepeatRules) * LayerCounterForCurrentRule
                                                End If
                                            End If
                                            
                                        ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Custom GCODE" Then
                                            
                                            If arrRepeatRules(5, iRepeatRules) = "GenericParameter" Then
                                                iParamVaried = arrRepeatRules(6, iRepeatRules)
                                                arrCustomGCODEtemp(iParamVaried - 1) = arrRepeatRules(7, iRepeatRules)
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "GenericParameterIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                iParamVaried = arrRepeatRules(6, iRepeatRules)
                                                arrCustomGCODEtemp(iParamVaried - 1) = arrCustomGCODEtemp(iParamVaried - 1) + arrRepeatRules(7, iRepeatRules) * LayerCounterForCurrentRule
                                            End If

                                        Else
                                            
                                            
                                            If arrRepeatRules(5, iRepeatRules) = "GenericParameter" Or arrRepeatRules(5, iRepeatRules) = "GenericParameterIncrement" Then
                                                Progress_Box.Show vbModeless: DoEvents: Progress_Box.Label1.Caption = "This GenericParameter and GenericParameterIncrement repeat rules are only allowed for ""Restraction"" and ""Custom GCODE"" features": DoEvents
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "NomWidth" Then
                                                dblWidthNew = arrRepeatRules(6, iRepeatRules)
                                                'reset dblE so that it is recalculated
                                                dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "NomWidthIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                dblWidthNew = dblWidthNew + arrRepeatRules(6, iRepeatRules) * LayerCounterForCurrentRule
                                                'reset dblE so that it is recalculated
                                                dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "NomHeight" Then
                                                dblHeightNew = arrRepeatRules(6, iRepeatRules)
                                                'reset dblE so that it is recalculated
                                                dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "NomHeightIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                dblHeightNew = dblHeightNew + arrRepeatRules(6, iRepeatRules) * LayerCounterForCurrentRule
                                                'reset dblE so that it is recalculated
                                                dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "Fspeed" Then
                                                dblFspeed = arrRepeatRules(6, iRepeatRules)
                                                'reset dblE so that it is recalculated
                                                dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "FspeedIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                dblFspeed = dblFspeed + arrRepeatRules(6, iRepeatRules) * LayerCounterForCurrentRule
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetPolar" Then
                                                dblXcentre = arrRepeatRules(6, iRepeatRules)
                                                dblYcentre = arrRepeatRules(7, iRepeatRules)
                                                dblRotationAngle = arrRepeatRules(8, iRepeatRules)
                                                dblRadialDisplacement = arrRepeatRules(9, iRepeatRules)
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    Call RotateLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                                                    Call RadiallyDisplaceLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(10, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(11, iRepeatRules) = "YES" Then
                                                    Call RotateLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                                                    Call RadiallyDisplaceLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                                                ElseIf arrRepeatRules(11, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(11, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetPolarIncrement" Then
                                                dblXcentre = arrRepeatRules(6, iRepeatRules)
                                                dblYcentre = arrRepeatRules(7, iRepeatRules)
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                dblRotationAngle = arrRepeatRules(8, iRepeatRules) * LayerCounterForCurrentRule
                                                dblRadialDisplacement = arrRepeatRules(9, iRepeatRules) * LayerCounterForCurrentRule
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    Call RotateLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                                                    Call RadiallyDisplaceLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(10, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(11, iRepeatRules) = "YES" Then
                                                    Call RotateLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                                                    Call RadiallyDisplaceLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                                                ElseIf arrRepeatRules(11, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(11, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetPolarMaths" Then
                                                dblXcentre = arrRepeatRules(6, iRepeatRules)
                                                dblYcentre = arrRepeatRules(7, iRepeatRules)
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                strAngleEquationRadians = arrRepeatRules(8, iRepeatRules)
                                                strRadiusEquation = arrRepeatRules(9, iRepeatRules)
                                                strZequation = arrRepeatRules(10, iRepeatRules)
                                                Call DetermineMathsOffsetPolar(dblXcentre, dblYcentre, strAngleEquationRadians, strRadiusEquation, strZequation, dblXnew1, dblYnew1, dblZnew1, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(11, iRepeatRules) = "YES" Then
                                                    dblXnew1 = dblXnew1 + dblLinearMathsOffsetX
                                                    dblYnew1 = dblYnew1 + dblLinearMathsOffsetY
                                                    dblZnew1 = dblZnew1 + dblLinearMathsOffsetZ
                                                ElseIf arrRepeatRules(11, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(8, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the end of the line, move the end position
                                                strAngleEquationRadians = arrRepeatRules(8, iRepeatRules)
                                                strRadiusEquation = arrRepeatRules(9, iRepeatRules)
                                                strZequation = arrRepeatRules(10, iRepeatRules)
                                                Call DetermineMathsOffsetPolar(dblXcentre, dblYcentre, strAngleEquationRadians, strRadiusEquation, strZequation, dblXnew2, dblYnew2, dblZnew2, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                                                If arrRepeatRules(12, iRepeatRules) = "YES" Then
                                                    dblXnew2 = dblXnew2 + dblLinearMathsOffsetX
                                                    dblYnew2 = dblYnew2 + dblLinearMathsOffsetY
                                                    dblZnew2 = dblZnew2 + dblLinearMathsOffsetZ
                                                ElseIf arrRepeatRules(12, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the current Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(9, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetLinear" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(9, iRepeatRules) = "YES" Then
                                                    dblXnew1 = dblXnew1 + arrRepeatRules(6, iRepeatRules)
                                                    dblYnew1 = dblYnew1 + arrRepeatRules(7, iRepeatRules)
                                                    dblZnew1 = dblZnew1 + arrRepeatRules(8, iRepeatRules)
                                                ElseIf arrRepeatRules(9, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(8, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the end of the line, move the end position
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    dblXnew2 = dblXnew2 + arrRepeatRules(6, iRepeatRules)
                                                    dblYnew2 = dblYnew2 + arrRepeatRules(7, iRepeatRules)
                                                    dblZnew2 = dblZnew2 + arrRepeatRules(8, iRepeatRules)
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(9, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetLinearIncrement" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(9, iRepeatRules) = "YES" Then
                                                    dblXnew1 = dblXnew1 + CDbl(arrRepeatRules(6, iRepeatRules)) * LayerCounterForCurrentRule
                                                    dblYnew1 = dblYnew1 + CDbl(arrRepeatRules(7, iRepeatRules)) * LayerCounterForCurrentRule
                                                    dblZnew1 = dblZnew1 + CDbl(arrRepeatRules(8, iRepeatRules)) * LayerCounterForCurrentRule
                                                ElseIf arrRepeatRules(9, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(8, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the end of the line, move the end position
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    dblXnew2 = dblXnew2 + CDbl(arrRepeatRules(6, iRepeatRules)) * LayerCounterForCurrentRule
                                                    dblYnew2 = dblYnew2 + CDbl(arrRepeatRules(7, iRepeatRules)) * LayerCounterForCurrentRule
                                                    dblZnew2 = dblZnew2 + CDbl(arrRepeatRules(8, iRepeatRules)) * LayerCounterForCurrentRule
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(9, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetLinearIncrementGraded" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(9, iRepeatRules) = "YES" Then
                                                    dblXnew1 = dblXnew1 + CDbl(arrRepeatRules(6, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                    dblYnew1 = dblYnew1 + CDbl(arrRepeatRules(7, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                    dblZnew1 = dblZnew1 + CDbl(arrRepeatRules(8, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                ElseIf arrRepeatRules(9, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(8, iRepeatRules) & """": End
                                                End If
                                                'If this rule applies to the end of the line, move the end position
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    dblXnew2 = dblXnew2 + CDbl(arrRepeatRules(6, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                    dblYnew2 = dblYnew2 + CDbl(arrRepeatRules(7, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                    dblZnew2 = dblZnew2 + CDbl(arrRepeatRules(8, iRepeatRules)) * (LayerCounterForCurrentRule * (LayerCounterForCurrentRule + 1) / 2)
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(9, iRepeatRules) & """": End
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "OffsetLinearMaths" Then
                                                LayerCounterForCurrentRule = (j - arrRepeatRules(3, iRepeatRules)) + 1
                                                strXequation = arrRepeatRules(6, iRepeatRules)
                                                strYequation = arrRepeatRules(7, iRepeatRules)
                                                strZequation = arrRepeatRules(8, iRepeatRules)
                                                Call DetermineMathsOffset(strXequation, strYequation, strZequation, dblXnew1, dblYnew1, dblZnew1, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                                                'If this rule applies to the start of the line, move the start position
                                                If arrRepeatRules(9, iRepeatRules) = "YES" Then
                                                    dblXnew1 = dblXnew1 + dblLinearMathsOffsetX
                                                    dblYnew1 = dblYnew1 + dblLinearMathsOffsetY
                                                    dblZnew1 = dblZnew1 + dblLinearMathsOffsetZ
                                                ElseIf arrRepeatRules(9, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(8, iRepeatRules) & """"
                                                End If
                                                'If this rule applies to the end of the line, move the end position
                                                strXequation = arrRepeatRules(6, iRepeatRules)
                                                strYequation = arrRepeatRules(7, iRepeatRules)
                                                strZequation = arrRepeatRules(8, iRepeatRules)
                                                Call DetermineMathsOffset(strXequation, strYequation, strZequation, dblXnew2, dblYnew2, dblZnew2, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                                                If arrRepeatRules(10, iRepeatRules) = "YES" Then
                                                    dblXnew2 = dblXnew2 + dblLinearMathsOffsetX
                                                    dblYnew2 = dblYnew2 + dblLinearMathsOffsetY
                                                    dblZnew2 = dblZnew2 + dblLinearMathsOffsetZ
                                                ElseIf arrRepeatRules(10, iRepeatRules) = "NO" Then 'Do nothing
                                                Else: MsgBox "Error: the value entered for the Repeat Rule written as Feature " & arrRepeatRules(13, iRepeatRules) & " should be ""YES"" or ""NO"". Current value is """ & arrRepeatRules(9, iRepeatRules) & """"
                                                End If
                                                'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                                                If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                                            End If
                                            If arrRepeatRules(5, iRepeatRules) = "ChangeTool" Then
                                                iToolNumber = arrRepeatRules(6, iRepeatRules)
                                            End If
                                            
                                        End If
                                        
                                    End If
                                End If
                            End If
                        Next iRepeatRulesRepeatFeatNumber_DOING_REPEATING
                    Next iRepeatRules
                End If
                
                If arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Retraction" Then
                    Call AddRetraction(dblRetractE, dblRetractSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand)
                    
                ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "RetractionZhop" Then
                    Call AddRetractionZhop(dblRetractZhop, dblRetractZhopSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblCurrentZ)
                    
                ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Custom GCODE" Then
                    Call AddCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, arrCustomGCODEtemp)
                    
                Else
                    '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                    dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
                    Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                    
                End If
                
            Next k
        Next j
        
        If arrFeatures(i, 1) = "Cartesian repeat" Then
            lCurrentCartesianRepFeature = lCurrentCartesianRepFeature + 1
        ElseIf arrFeatures(i, 1) = "Polar repeat" Then
            lCurrentPolarRepFeature = lCurrentPolarRepFeature + 1
        Else
            MsgBox "This bit of code should not be reached"
        End If
    
    
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''REFLECT FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@



    ElseIf arrFeatures(i, 1) = "Reflect" Then
        
        If arrFeatures(i, 3) = "XY" Then
            dblXreflect1 = arrFeatures(i, 4)
            dblYreflect1 = arrFeatures(i, 5)
            dblXreflect2 = arrFeatures(i, 6)
            dblYreflect2 = arrFeatures(i, 7)
        ElseIf arrFeatures(i, 3) = "Polar" Then
            dblXreflect1 = arrFeatures(i, 4)
            dblYreflect1 = arrFeatures(i, 5)
            dblXreflect2 = dblXreflect1 + Cos(CDbl(arrFeatures(i, 6)) * Pi / 180)
            dblYreflect2 = dblYreflect1 + Sin(CDbl(arrFeatures(i, 6)) * Pi / 180)
        ElseIf arrFeatures(i, 3) = "Z" Then
            dblZreflect = arrFeatures(i, 4)
        End If
        
        

        'Find which commands/lines to repeat
        'This naturally excluded automatically added travel and toolchange commands because they have a cID of 0
        strFeatureList = arrFeatures(i, 2)
        Call ReplaceDashesInString(strFeatureList)
        arrFeatureList = Split(strFeatureList, ",")
        For j = 0 To UBound(arrFeatureList)
            For k = 1 To UBound(arrCommands, 2)
                If arrCommands(cID, k) = CInt(arrFeatureList(j)) Then
                    lNumberOfCommandsBeingRepeated = lNumberOfCommandsBeingRepeated + 1
                    ReDim Preserve arrCommandsBeingRepeated(1 To lNumberOfCommandsBeingRepeated)
                    arrCommandsBeingRepeated(lNumberOfCommandsBeingRepeated) = k
                End If
            Next k
        Next j
        
        'Run through all commands being reflected in REVERSE order
        For j = lNumberOfCommandsBeingRepeated To 1 Step -1
        
            Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(this is a ""reflect"" feature)" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
            DoEvents
            
            lCurrentCommandBeingRepeated = arrCommandsBeingRepeated(j)
                        
            If arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Retraction" Then
                
                'Since lines are printed in reverse order for reflections, the retraction needs to be the opposite (retraction/unretraction) of what is originally was. Hence the minus sign.
                dblRetractE = -arrCommands(cRetractE, lCurrentCommandBeingRepeated)
                dblRetractSpeed = arrCommands(cRetractSpeed, lCurrentCommandBeingRepeated)
                
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "RetractionZhop" Then
            
                'Since lines are printed in reverse order for reflections, the retraction needs to be the opposite (retraction/unretraction) of what is originally was. Hence the minus sign.
                dblRetractZhop = -arrCommands(cRetractZhop, lCurrentCommandBeingRepeated)
                dblRetractZhopSpeed = arrCommands(cRetractZhopSpeed, lCurrentCommandBeingRepeated)

            ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Custom GCODE" Then

                strCustomGCODE = arrCommands(cGCODE, lCurrentCommandBeingRepeated)
                For k1 = 1 To UBound(arrCustomGCODEtemp)
                    arrCustomGCODEtemp(k1) = arrCommands(k1 + 1, lCurrentCommandBeingRepeated)
                Next k1
                
            Else
                    
                'Find the X and Y coords for the line being repeated
                dblXold1 = arrCommands(cX1, lCurrentCommandBeingRepeated)
                dblYold1 = arrCommands(cY1, lCurrentCommandBeingRepeated)
                dblZold1 = arrCommands(cZ1, lCurrentCommandBeingRepeated)
                dblXold2 = arrCommands(cX2, lCurrentCommandBeingRepeated)
                dblYold2 = arrCommands(cY2, lCurrentCommandBeingRepeated)
                dblZold2 = arrCommands(cZ2, lCurrentCommandBeingRepeated)
                strPrintTravelNew = arrCommands(cCommandType, lCurrentCommandBeingRepeated)
                dblWidthNew = arrCommands(cW, lCurrentCommandBeingRepeated)
                dblHeightNew = arrCommands(cH, lCurrentCommandBeingRepeated)
                dblE = arrCommands(cE, lCurrentCommandBeingRepeated)
                dblFspeed = arrCommands(cF, lCurrentCommandBeingRepeated)
                iToolNumber = arrCommands(cT, lCurrentCommandBeingRepeated)
                
                strFeatIDrenumbered = "ReflectFeat" & lCurrentReflectRepFeature & " of " & arrCommands(cNotes, lCurrentCommandBeingRepeated)
                strFeatIDtree = arrCommands(cIDtree, lCurrentCommandBeingRepeated) & "-" & i & ".1"
    
                
                If arrFeatures(i, 3) = "XY" Or arrFeatures(i, 3) = "Polar" Then
                
                    'reflect line
                    Call ReflectLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXreflect1, dblYreflect1, dblXreflect2, dblYreflect2, dblXnew1, dblYnew1, dblXnew2, dblYnew2)
                    
                    dblZnew1 = dblZold1
                    dblZnew2 = dblZold2
                    
                    'Swap direction of line
                    dblXnew1TEMP = dblXnew1
                    dblYnew1TEMP = dblYnew1
                    dblZnew1TEMP = dblZnew1
                    dblXnew1 = dblXnew2
                    dblYnew1 = dblYnew2
                    dblZnew1 = dblZnew2
                    dblXnew2 = dblXnew1TEMP
                    dblYnew2 = dblYnew1TEMP
                    dblZnew2 = dblZnew1TEMP
                
                ElseIf arrFeatures(i, 3) = "Z" Then
                
                    dblXnew1 = dblXold2
                    dblXnew2 = dblXold1
                    dblYnew1 = dblYold2
                    dblYnew2 = dblYold1
                    
                    'Increase Z by double the distance to the written value
                    dblZnew1 = dblZold1 + ((dblZreflect - dblZold1) * 2)
                    dblZnew2 = dblZold2 + ((dblZreflect - dblZold2) * 2)
                
                End If
                
            End If
            
            If arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Retraction" Then
                Call AddRetraction(dblRetractE, dblRetractSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand)
                
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "RetractionZhop" Then
                Call AddRetractionZhop(dblRetractZhop, dblRetractZhopSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblCurrentZ)
                
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingRepeated) = "Custom GCODE" Then
                Call AddCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, arrCustomGCODEtemp)
                    
            Else
                '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew1, dblYnew1, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                dblCurrentX = dblXnew1: dblCurrentY = dblYnew1: dblCurrentZ = dblZnew1
                Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                
            End If

        Next j
        
        lCurrentReflectRepFeature = lCurrentReflectRepFeature + 1



    
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''CONCENTRIC REPEAT FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    ElseIf arrFeatures(i, 1) = "Concentric repeat" Then
    
        
        
        'Find layer height and number of repeats
        strInsideOutside = arrFeatures(i, 3)
        iNumberOfRepeats = arrFeatures(i, 4)
        dblOffset = arrFeatures(i, 5)
        If strInsideOutside = "Outside" Then
            dblOffset = -dblOffset
        End If
        
        'Find which features to repeat
        arrFeatureList = Split(arrFeatures(i, 2), ",")
        
        For j = 0 To UBound(arrFeatureList)
            
            iFeatureNumber = arrFeatureList(j)
        
            For k = 1 To iNumberOfRepeats
            
                If arrFeatures(iFeatureNumber, 1) = "Rectangle" Then
                    
                    dblStartCornerX = arrFeatures(iFeatureNumber, 2)
                    dblStartCornerY = arrFeatures(iFeatureNumber, 3)
                    dblRecSizeX = arrFeatures(iFeatureNumber, 4) - dblStartCornerX
                    dblRecSizeY = arrFeatures(iFeatureNumber, 5) - dblStartCornerY
                    
                    If dblRecSizeX > 0 Then
                        dblOffsetX = dblOffset
                    Else
                        dblOffsetX = -dblOffset
                    End If
                    
                    If dblRecSizeY > 0 Then
                        dblOffsetY = dblOffset
                    Else
                        dblOffsetY = -dblOffset
                    End If
                    
                    dblStartCornerX = dblStartCornerX + dblOffsetX * k
                    dblStartCornerY = dblStartCornerY + dblOffsetY * k
                    dblRecSizeX = dblRecSizeX - 2 * dblOffsetX * k
                    dblRecSizeY = dblRecSizeY - 2 * dblOffsetY * k
                    
                    
                    dblZnew1 = arrFeatures(iFeatureNumber, 6)
                    dblZnew2 = dblZnew1
                    strPrintTravelNew = "Print"
                    dblWidthNew = arrFeatures(iFeatureNumber, 8)
                    dblHeightNew = arrFeatures(iFeatureNumber, 9)
                    
                    'reset dblE so that it is recalculated for the new size (changes due to concentric repeat) - similar for Fspeed
                    dblE = 0: dblFspeed = 0
                    'These values will be changed in the next pair of commands if the rectangle had an EvaluePerSegment or Fspeed defined by the user
                                        
                    'Convert the list of additional parameters (from the original rectangle feature) into individual values.
                    strAdditionalParams = arrFeatures(iFeatureNumber, 10)
                    Call DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)
                    
                    strFeatIDrenumbered = "ConcentricRepeat" & lCurrentConcentricRepeatFeature
                    'The next line doesn't work because there is no "lCurrentCommandBeingRepeated"
                    'strFeatIDtree = i & "." & k & "-" & arrCommands(cIDtree, lCurrentCommandBeingRepeated)
                    'Just refer to the rectangle feature number, because the concentric infill can only refer to a rectangle and therefore it won't be a derivative/repeated feature
                    strFeatIDtree = i & "." & k & "-" & iFeatureNumber
                    
                    '''''(THE ADD TRAVEL/TOOLCHANGE COMMAND IS NOW DONE AT THE END OF THE PROGRAM, BEFORE GCODE GENERATION)'''''Call AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX, dblCurrentY, dblCurrentZ, dblStartCornerX, dblStartCornerY, dblZnew1, arrCommands, lCurrentCommand, i, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                    dblCurrentX = dblStartCornerX: dblCurrentY = dblStartCornerY: dblCurrentZ = dblZnew1
                    
                    If arrFeatures(iFeatureNumber, 7) = "CW" Then
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY + dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX + dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY - dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX - dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                
                    Else
                        
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX + dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY + dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX - dblRecSizeX, dblCurrentY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        Call AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblCurrentX, dblCurrentY - dblRecSizeY, dblCurrentZ, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)
                        
                    End If
                    
                End If
            Next k
        Next j
        
        lCurrentConcentricRepeatFeature = lCurrentConcentricRepeatFeature + 1
        
    
    
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''POSTPROCESS FEATURE
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    ElseIf arrFeatures(i, 1) = "Postprocess" Then
        'Go through all the arrcommands and edit them if they're included in the feature list
        
        'Fix any lowercase terms in equations
        If arrFeatures(i, 3) = "OffsetLinearMaths" Then
            arrFeatures(i, 4) = Replace(arrFeatures(i, 4), "xval", "Xval")
            arrFeatures(i, 4) = Replace(arrFeatures(i, 4), "yval", "Yval")
            arrFeatures(i, 4) = Replace(arrFeatures(i, 4), "zval", "Zval")
            arrFeatures(i, 4) = Replace(arrFeatures(i, 4), "repval", "REPval")
            arrFeatures(i, 5) = Replace(arrFeatures(i, 5), "xval", "Xval")
            arrFeatures(i, 5) = Replace(arrFeatures(i, 5), "yval", "Yval")
            arrFeatures(i, 5) = Replace(arrFeatures(i, 5), "zval", "Zval")
            arrFeatures(i, 5) = Replace(arrFeatures(i, 5), "repval", "REPval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "xval", "Xval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "yval", "Yval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "zval", "Zval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "repval", "REPval")
        End If
        If arrFeatures(i, 3) = "OffsetPolarMaths" Then
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "xval", "Xval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "yval", "Yval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "zval", "Zval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "aval", "Aval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "rval", "Rval")
            arrFeatures(i, 6) = Replace(arrFeatures(i, 6), "repval", "REPval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "xval", "Xval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "yval", "Yval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "zval", "Zval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "aval", "Aval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "rval", "Rval")
            arrFeatures(i, 7) = Replace(arrFeatures(i, 7), "repval", "REPval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "xval", "Xval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "yval", "Yval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "zval", "Zval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "aval", "Aval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "rval", "Rval")
            arrFeatures(i, 8) = Replace(arrFeatures(i, 8), "repval", "REPval")
        End If
        
        
        dblElapsedTime = Timer - dblStartTime
        Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(postprocessing)" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
        DoEvents
        
        'Find which commands/lines to repeat
        strFeatureList = arrFeatures(i, 2)
        
        'This check allows me to call the QUICK version of "CheckIfCurrentCommandSatisfiesInclusionCriteria" since there are no "Y" or "N" criteria
        'This check is only done once instead of every time the function "CheckIfCurrentCommandSatisfiesInclusionCriteria" is ran
        If strFeatureList = "All" Or strFeatureList = "all" Then
            For k = 1 To UBound(arrCommands, 2)
                lNumberOfCommandsBeingModified = lNumberOfCommandsBeingModified + 1
                ReDim Preserve arrCommandsBeingModified(1 To lNumberOfCommandsBeingModified)
                arrCommandsBeingModified(lNumberOfCommandsBeingModified) = k
            Next k
        ElseIf strFeatureList Like "*Y*" = False And strFeatureList Like "*N*" = False Then
            'Still not fast...
            For k = 1 To UBound(arrCommands, 2)
                If CheckIfCurrentCommandSatisfiesInclusionCriteria_QUICK(arrCommands, k, strFeatureList, CStr(arrCommands(cIDtree, k))) = True Then
                    lNumberOfCommandsBeingModified = lNumberOfCommandsBeingModified + 1
                    ReDim Preserve arrCommandsBeingModified(1 To lNumberOfCommandsBeingModified)
                    arrCommandsBeingModified(lNumberOfCommandsBeingModified) = k
                End If
            Next k
        Else
            For k = 1 To UBound(arrCommands, 2)
                If CheckIfCurrentCommandSatisfiesInclusionCriteria(arrCommands, k, strFeatureList, CStr(arrCommands(cIDtree, k))) = True Then
                    lNumberOfCommandsBeingModified = lNumberOfCommandsBeingModified + 1
                    ReDim Preserve arrCommandsBeingModified(1 To lNumberOfCommandsBeingModified)
                    arrCommandsBeingModified(lNumberOfCommandsBeingModified) = k
                End If
            Next k
        End If
        
        

        
        dblElapsedTime = Timer - dblStartTime
        Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(postprocessing )" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
        DoEvents

        
        'Go through each command that is included (and therefore is being modified)
        For k = 1 To lNumberOfCommandsBeingModified
            
            
            If lNumberOfCommandsBeingModified > 1000 Then
                If k Mod 1000 = 0 Then
                    dblElapsedTime = Timer - dblStartTime
                    Progress_Box.Label1.Caption = "Working on feature " & i & " of " & iNumberOfFeatures & vbNewLine & "(postprocessing command " & Format(k, "#,###") & " of " & Format(lNumberOfCommandsBeingModified, "#,###") & ")" & vbNewLine & vbNewLine & "Time elapsed = " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec"
                    DoEvents
                End If
            End If
            
            lCurrentCommandBeingModified = arrCommandsBeingModified(k)
            
            'Copy the old parameters
            strFeatIDrenumbered = "Postprocessed version of " & arrCommands(cNotes, lCurrentCommandBeingModified)
            strFeatIDtree = i & "-" & arrCommands(cIDtree, lCurrentCommandBeingModified)
            
            If arrCommands(cCommandType, lCurrentCommandBeingModified) = "Retraction" Then
                dblRetractE = arrCommands(cRetractE, lCurrentCommandBeingModified)
                dblRetractSpeed = arrCommands(cRetractSpeed, lCurrentCommandBeingModified)
            
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingModified) = "RetractionZhop" Then
                dblRetractZhop = arrCommands(cRetractZhop, lCurrentCommandBeingModified)
                dblRetractZhopSpeed = arrCommands(cRetractZhopSpeed, lCurrentCommandBeingModified)
            
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingModified) = "Custom GCODE" Then
                strCustomGCODE = arrCommands(cGCODE, lCurrentCommandBeingModified)
                For k1 = 1 To UBound(arrCustomGCODEtemp)
                    arrCustomGCODEtemp(k1) = arrCommands(k1 + 1, lCurrentCommandBeingModified)
                Next k1
                
            Else
                dblXnew1 = arrCommands(cX1, lCurrentCommandBeingModified)
                dblYnew1 = arrCommands(cY1, lCurrentCommandBeingModified)
                dblZnew1 = arrCommands(cZ1, lCurrentCommandBeingModified)
                dblXnew2 = arrCommands(cX2, lCurrentCommandBeingModified)
                dblYnew2 = arrCommands(cY2, lCurrentCommandBeingModified)
                dblZnew2 = arrCommands(cZ2, lCurrentCommandBeingModified)
                strPrintTravelNew = arrCommands(cCommandType, lCurrentCommandBeingModified)
                dblWidthNew = arrCommands(cW, lCurrentCommandBeingModified)
                dblHeightNew = arrCommands(cH, lCurrentCommandBeingModified)
                dblE = arrCommands(cE, lCurrentCommandBeingModified)
                dblFspeed = arrCommands(cF, lCurrentCommandBeingModified)
                iToolNumber = arrCommands(cT, lCurrentCommandBeingModified)
                
            End If
            
            
            
            
            If arrCommands(cCommandType, lCurrentCommandBeingModified) = "Retraction" _
            Or arrCommands(cCommandType, lCurrentCommandBeingModified) = "RetractionZhop" _
            Then
                'Do nothing, although code could be here (as it is for repeat rules)
            ElseIf arrCommands(cCommandType, lCurrentCommandBeingModified) = "Custom GCODE" Then
                'Do nothing, although code could be here (as it is for repeat rules)
            Else
                'This is a printed line, so modify it accordingly
'                If arrFeatures(i, 3) = "GenericParameter" Or arrFeatures(i, 3) = "GenericParameterIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "NomWidth" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "NomWidthIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "NomHeight" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "NomHeightIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "Fspeed" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "FspeedIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
                If arrFeatures(i, 3) = "OffsetLinear" Then
                    'If this rule applies to the start of the line, move the start position
                    If arrFeatures(i, 7) = "YES" Then
                        dblXnew1 = dblXnew1 + arrFeatures(i, 4)
                        dblYnew1 = dblYnew1 + arrFeatures(i, 5)
                        dblZnew1 = dblZnew1 + arrFeatures(i, 6)
                    ElseIf arrFeatures(i, 7) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 7) & """": End
                    End If
                    'If this rule applies to the end of the line, move the end position
                    If arrFeatures(i, 8) = "YES" Then
                        dblXnew2 = dblXnew2 + arrFeatures(i, 4)
                        dblYnew2 = dblYnew2 + arrFeatures(i, 5)
                        dblZnew2 = dblZnew2 + arrFeatures(i, 6)
                    ElseIf arrFeatures(i, 8) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 8) & """": End
                    End If
                    'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                    If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                End If
'                If arrFeatures(i, 3) = "OffsetLinearIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
'                If arrFeatures(i, 3) = "OffsetLinearIncrementGraded" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
                If arrFeatures(i, 3) = "OffsetLinearMaths" Then
                    strXequation = arrFeatures(i, 4)
                    strYequation = arrFeatures(i, 5)
                    strZequation = arrFeatures(i, 6)
                    Call DetermineMathsOffset(strXequation, strYequation, strZequation, dblXnew1, dblYnew1, dblZnew1, LayerCounterForCurrentRule_NOT_USED, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                    'If this rule applies to the start of the line, move the start position
                    If arrFeatures(i, 7) = "YES" Then
                        dblXnew1 = dblXnew1 + dblLinearMathsOffsetX
                        dblYnew1 = dblYnew1 + dblLinearMathsOffsetY
                        dblZnew1 = dblZnew1 + dblLinearMathsOffsetZ
                    ElseIf arrFeatures(i, 7) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 7) & """": End
                    End If
                    'If this rule applies to the end of the line, move the end position
                    strXequation = arrFeatures(i, 4)
                    strYequation = arrFeatures(i, 5)
                    strZequation = arrFeatures(i, 6)
                    Call DetermineMathsOffset(strXequation, strYequation, strZequation, dblXnew2, dblYnew2, dblZnew2, LayerCounterForCurrentRule_NOT_USED, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                    If arrFeatures(i, 8) = "YES" Then
                        dblXnew2 = dblXnew2 + dblLinearMathsOffsetX
                        dblYnew2 = dblYnew2 + dblLinearMathsOffsetY
                        dblZnew2 = dblZnew2 + dblLinearMathsOffsetZ
                    ElseIf arrFeatures(i, 8) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 8) & """": End
                    End If
                    'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                    If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                End If
                If arrFeatures(i, 3) = "OffsetPolar" Then
                    dblXcentre = arrFeatures(i, 4)
                    dblYcentre = arrFeatures(i, 5)
                    dblRotationAngle = arrFeatures(i, 6)
                    dblRadialDisplacement = arrFeatures(i, 7)
                    'If this rule applies to the start of the line, move the start position
                    If arrFeatures(i, 8) = "YES" Then
                        Call RotateLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                        Call RadiallyDisplaceLine(dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, dblXnew1, dblYnew1, TEMPVARFORdblXnew2, TEMPVARFORdblYnew2)
                    ElseIf arrFeatures(i, 8) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 8) & """": End
                    End If
                    'If this rule applies to the start of the line, move the start position
                    If arrFeatures(i, 9) = "YES" Then
                        Call RotateLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRotationAngle * Pi / 180, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                        Call RadiallyDisplaceLine(TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2, dblXcentre, dblYcentre, dblRadialDisplacement, TEMPVARFORdblXnew1, TEMPVARFORdblYnew1, dblXnew2, dblYnew2)
                    ElseIf arrFeatures(i, 9) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 9) & """": End
                    End If
                    'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                    If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                End If
'                If arrFeatures(i, 3) = "OffsetPolarIncrement" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
                If arrFeatures(i, 3) = "OffsetPolarMaths" Then
                    dblXcentre = arrFeatures(i, 4)
                    dblYcentre = arrFeatures(i, 5)
                    strAngleEquationRadians = arrFeatures(i, 6)
                    strRadiusEquation = arrFeatures(i, 7)
                    strZequation = arrFeatures(i, 8)
                    Call DetermineMathsOffsetPolar(dblXcentre, dblYcentre, strAngleEquationRadians, strRadiusEquation, strZequation, dblXnew1, dblYnew1, dblZnew1, LayerCounterForCurrentRule_NOT_USED, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                    'If this rule applies to the start of the line, move the start position
                    If arrFeatures(i, 9) = "YES" Then
                        dblXnew1 = dblXnew1 + dblLinearMathsOffsetX
                        dblYnew1 = dblYnew1 + dblLinearMathsOffsetY
                        dblZnew1 = dblZnew1 + dblLinearMathsOffsetZ
                    ElseIf arrFeatures(i, 9) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 9) & """": End
                    End If
                    'If this rule applies to the end of the line, move the end position
                    strAngleEquationRadians = arrFeatures(i, 6)
                    strRadiusEquation = arrFeatures(i, 7)
                    strZequation = arrFeatures(i, 8)
                    Call DetermineMathsOffsetPolar(dblXcentre, dblYcentre, strAngleEquationRadians, strRadiusEquation, strZequation, dblXnew2, dblYnew2, dblZnew2, LayerCounterForCurrentRule_NOT_USED, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)
                    If arrFeatures(i, 10) = "YES" Then
                        dblXnew2 = dblXnew2 + dblLinearMathsOffsetX
                        dblYnew2 = dblYnew2 + dblLinearMathsOffsetY
                        dblZnew2 = dblZnew2 + dblLinearMathsOffsetZ
                    ElseIf arrFeatures(i, 10) = "NO" Then 'Do nothing
                    Else: MsgBox "Error: the value entered for Feature " & i & " should be ""YES"" or ""NO"". Current value is """ & arrFeatures(i, 10) & """": End
                    End If
                    'reset dblE so that it is recalculated for the new line's length (UNLESS either of the NomWidth or NomHeight values are zero, which indicates the E value was manually overridden and therefore should not be changed)
                    If dblWidthNew * dblHeightNew > 0 Then dblE = 0
                End If
'                If arrFeatures(i, 3) = "ChangeTool" Then
'                    'Do nothing, although code could be here (as it is for repeat rules)
'                End If
            End If
                   
        
        If arrCommands(cCommandType, lCurrentCommandBeingModified) = "Retraction" Then
            'CURRENTLY ONLY PROGRAMMED FOR PRINTED LINES/TRAVEL
'            Call AddRetraction(dblRetractE, dblRetractSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand)
            lCurrentCommandBeingModified = lCurrentCommandBeingModified + 1
            
        ElseIf arrCommands(cCommandType, lCurrentCommandBeingModified) = "RetractionZhop" Then
            'CURRENTLY ONLY PROGRAMMED FOR PRINTED LINES/TRAVEL
'            Call AddRetractionZhop(dblRetractZhop, dblRetractZhopSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblCurrentZ)
            lCurrentCommandBeingModified = lCurrentCommandBeingModified + 1
            
        ElseIf arrCommands(cCommandType, lCurrentCommandBeingModified) = "Custom GCODE" Then
            'CURRENTLY ONLY PROGRAMMED FOR PRINTED LINES/TRAVEL
'            Call AddCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, arrCustomGCODEtemp)
            lCurrentCommandBeingModified = lCurrentCommandBeingModified + 1
            
        Else
            'Use the "AddLine" subroutine to replace the original line with the updated one
            Call AddLine(dblXnew1, dblYnew1, dblZnew1, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, arrCommands(cID, lCurrentCommandBeingModified), strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommandBeingModified, dblFeedstockFilamentDiameter, strExtrusionUnits)
        End If
        
        Next k
    



    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ''''''IGNORED OR UNRECOGNISED FEATURES
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    ElseIf Len(arrFeatures(i, 1)) >= 4 Then
        If Left(arrFeatures(i, 1), 4) = "SKIP" Then
            'Do nothing
        ElseIf Left(arrFeatures(i, 1), 4) = "STOP" Then
            MsgBox "GCODE creation was stopped before feature " & i & " because its name was prefixed by ""STOP"""
            Exit For
        ElseIf arrFeatures(i, 1) = "Repeat rule" Or arrFeatures(i, 1) = "Reproduce and recalculate" Then
            'Do nothing - these feature-types are processed elsewhere
        Else
            MsgBox "Name of feature " & i & " is not recognised" & vbNewLine & "Current name is """ & arrFeatures(i, 1) & """" & vbNewLine & vbNewLine & "Feature will not be processed"
        End If
    Else
        MsgBox "Name of feature " & i & " is not recognised" & vbNewLine & "Current name is """ & arrFeatures(i, 1) & """" & vbNewLine & vbNewLine & "Feature will not be processed"
    
    
    
    
    End If
    
    
    
    
    
    
    lNumberOfCommandsBeingRepeated = 0
    lNumberOfCommandsBeingModified = 0
    ReDim arrCommandsBeingRepeated(1 To 1)
    ReDim arrCommandsBeingModified(1 To 1)
    
Next i





If bCheckErrors Then
    On Error GoTo ErrorHandlerPost
End If



Progress_Box.Label1.Caption = "Adding automaticlaly generated travel and toolchange commands if required"

Call AddTravelAndToolChangeCommandsToWholeArray(lAddedCommandsCounter, lCurrentCommand, arrCommands, arrToolChangeGCODE, dblInitialX, dblInitialY, dblInitialZ, iInitialTool, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
                 
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    

Progress_Box.Label1.Caption = "Formatting GCODE"

DoEvents

'Redim the array to remove empy columns (because the array is increased in chunks of 10,000 columns
'ReDim Preserve arrCommands(1 To UBound(arrCommands, 1), 1 To lCurrentCommand - 1)

ReDim Preserve arrCommands(1 To UBound(arrCommands, 1), 1 To lCurrentCommand - 1 + lAddedCommandsCounter)

Call GenerateGCODE(arrCommands, dblXoffset, dblYoffset, dblZoffset, dblLayer1EMultiplier, dblLayer1SpeedMultiplier, dblCurrentSpeed)

'Transpose the array arrcommands (for copying into the Printpath sheet)
Dim tempArray As Variant
ReDim tempArray(1 To UBound(arrCommands, 2), 1 To UBound(arrCommands, 1))
For i = 1 To UBound(arrCommands, 2)
    For j = 1 To UBound(arrCommands, 1)
        tempArray(i, j) = arrCommands(j, i)
    Next j
Next i


Progress_Box.Label1.Caption = "Copying GCODE into Excel ""GCODE"" worksheet"
DoEvents
    

'Output the whole printpath generation information
Sheets("Printpath").Activate
ActiveSheet.Range("A2:Q1000000").ClearContents
ActiveSheet.Range(Cells(2, 1), Cells(2, 1).Offset(UBound(tempArray, 1) - 1, (UBound(tempArray, 2) - 1))).Value = tempArray
ActiveSheet.Range("O1").Select
Range(Selection, Selection.End(xlDown)).Select


'Output the full GCODE
ReDim arrFullGCODE(1 To UBound(arrStartGCODE, 1) + UBound(tempArray, 1) + UBound(arrEndGCODE, 1), 1 To 1)
For i = 1 To UBound(arrStartGCODE, 1)
    arrFullGCODE(i, 1) = arrStartGCODE(i, 1)
Next i
For i = 1 To UBound(arrCommands, 2)
    arrFullGCODE(i + UBound(arrStartGCODE, 1), 1) = arrCommands(cGCODE, i)
Next i
For i = 1 To UBound(arrEndGCODE, 1)
    arrFullGCODE(i + UBound(arrStartGCODE, 1) + UBound(arrCommands, 2), 1) = arrEndGCODE(i, 1)
Next i
Sheets("GCODE").Activate
ActiveSheet.Range("A:A").ClearContents
ActiveSheet.Range(Cells(1, 1), Cells(UBound(arrFullGCODE, 1), 1)).Value = arrFullGCODE
ActiveSheet.Range("A1").Select
Range(Selection, Range("A1000000").End(xlUp)).Select

Selection.Copy
Sheets("Main Sheet").Activate


Application.ScreenUpdating = True

Progress_Box.Hide
Unload Progress_Box

dblElapsedTime = Timer - dblStartTime
MsgBox "The GCODE has been generated and copied to clipboard" _
& vbNewLine & "Paste it into Notepad (to save as a GCODE file) or into GCODE preview software" _
& vbNewLine & vbNewLine & "Total time taken:          " & Int(dblElapsedTime / 60) & " min " & CLng(dblElapsedTime - 60 * Int(dblElapsedTime / 60)) & " sec" _
& vbNewLine & "Total lines of GCODE:   " & Format(UBound(arrFullGCODE, 1), "#,###") _
& vbNewLine & vbNewLine & "Please reference the associate journal paper in articles/videos/presentations resulting from FullControl GCODE Designer"

Done:
    Exit Sub
ErrorHandlerPre:
MsgBox "Error encountered during initial stages of setting up paramaters" & vbNewLine & vbNewLine & "The program will now end"
Progress_Box.Hide
Unload Progress_Box
Application.ScreenUpdating = True
Exit Sub
ErrorHandlerRepRec:
MsgBox "Error encountered when analysing ""Reproduce and recalculate"" features. Check feature " & i & vbNewLine & vbNewLine & "The program will now end"
Progress_Box.Hide
Unload Progress_Box
Application.ScreenUpdating = True
Exit Sub
ErrorHandlerRepRule:
MsgBox "Error encountered when analysing ""Repeat rule"" features. Check feature " & i & vbNewLine & vbNewLine & "The program will now end"
Progress_Box.Hide
Unload Progress_Box
Application.ScreenUpdating = True
Exit Sub
ErrorHandlerMain:
MsgBox "Error encountered when working on feature " & i & vbNewLine & "For repeating features, this error may be caused by errors with ""Repeat Rule"" features. Similar associations exist for other features" & vbNewLine & vbNewLine & "The program will now end"
Progress_Box.Hide
Unload Progress_Box
Application.ScreenUpdating = True
Exit Sub
ErrorHandlerPost:
MsgBox "Error encountered after going through all features (e.g. during GCODE format conversion)"
Progress_Box.Hide
Unload Progress_Box
Application.ScreenUpdating = True
Exit Sub

End Sub


Sub AddLine(dblCurrentX, dblCurrentY, dblCurrentZ, dblXnew2, dblYnew2, dblZnew2, strPrintTravelNew, dblWidthNew, dblHeightNew, dblE, dblFspeed, iToolNumber, iCurrentToolNumber, dblPrintSpeed, dblTravelSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblFeedstockFilamentDiameter, strExtrusionUnits)

Dim dblLength As Double

If strExtrusionUnits = "mm3" Then
    Emultiplier = 1
Else
    'Convert from volume to length
    Emultiplier = 1 / (Pi * (dblFeedstockFilamentDiameter / 2) ^ 2)
End If

'Add data to arrCommands
Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)
arrCommands(cX1, lCurrentCommand) = dblCurrentX
arrCommands(cY1, lCurrentCommand) = dblCurrentY
arrCommands(cZ1, lCurrentCommand) = dblCurrentZ
arrCommands(cX2, lCurrentCommand) = dblXnew2
arrCommands(cY2, lCurrentCommand) = dblYnew2
arrCommands(cZ2, lCurrentCommand) = dblZnew2
arrCommands(cCommandType, lCurrentCommand) = strPrintTravelNew

If strPrintTravelNew = "Print" Then
    arrCommands(cW, lCurrentCommand) = dblWidthNew
    arrCommands(cH, lCurrentCommand) = dblHeightNew
    If dblE <> 0 Then
        arrCommands(cE, lCurrentCommand) = dblE
    Else
        dblLength = ((dblXnew2 - dblCurrentX) ^ 2 + (dblYnew2 - dblCurrentY) ^ 2 + (dblZnew2 - dblCurrentZ) ^ 2) ^ 0.5
        arrCommands(cE, lCurrentCommand) = dblLength * dblWidthNew * dblHeightNew * Emultiplier
    End If
    
    If dblFspeed <> 0 Then
        arrCommands(cF, lCurrentCommand) = dblFspeed
    Else
        arrCommands(cF, lCurrentCommand) = dblPrintSpeed
    End If
    
Else
    arrCommands(cW, lCurrentCommand) = 0
    arrCommands(cH, lCurrentCommand) = 0
    arrCommands(cE, lCurrentCommand) = 0
    
    If dblFspeed <> 0 Then
        arrCommands(cF, lCurrentCommand) = dblFspeed
    Else
        arrCommands(cF, lCurrentCommand) = dblTravelSpeed
    End If
    
End If

dblCurrentX = dblXnew2
dblCurrentY = dblYnew2
dblCurrentZ = dblZnew2

If iToolNumber <> -1 Then
    arrCommands(cT, lCurrentCommand) = iToolNumber
    iCurrentToolNumber = iToolNumber
Else
    arrCommands(cT, lCurrentCommand) = iCurrentToolNumber
End If

arrCommands(cID, lCurrentCommand) = i
arrCommands(cIDtree, lCurrentCommand) = strFeatIDtree
arrCommands(cNotes, lCurrentCommand) = strFeatIDrenumbered


lCurrentCommand = lCurrentCommand + 1
                
End Sub

Sub AddRetraction(dblRetractE, dblRetractSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand)

'Add data to arrCommands
Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)

arrCommands(cCommandType, lCurrentCommand) = "Retraction"
arrCommands(cRetractE, lCurrentCommand) = dblRetractE
arrCommands(cRetractSpeed, lCurrentCommand) = dblRetractSpeed
arrCommands(cID, lCurrentCommand) = i
arrCommands(cIDtree, lCurrentCommand) = strFeatIDtree
arrCommands(cNotes, lCurrentCommand) = strFeatIDrenumbered

lCurrentCommand = lCurrentCommand + 1
        
End Sub

Sub AddRetractionZhop(dblRetractZhop, dblRetractZhopSpeed, i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, dblCurrentZ)

'Add data to arrCommands
Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)

arrCommands(cCommandType, lCurrentCommand) = "RetractionZhop"
arrCommands(cRetractZhop, lCurrentCommand) = dblRetractZhop
arrCommands(cRetractZhopSpeed, lCurrentCommand) = dblRetractZhopSpeed
arrCommands(cID, lCurrentCommand) = i
arrCommands(cIDtree, lCurrentCommand) = strFeatIDtree
arrCommands(cNotes, lCurrentCommand) = strFeatIDrenumbered

dblCurrentZ = dblCurrentZ + dblRetractZhop

lCurrentCommand = lCurrentCommand + 1
        
End Sub


Sub AddCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, arrCustomGCODEtemp)

Dim strCODE As String
'Add data to arrCommands
Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)

arrCommands(cCommandType, lCurrentCommand) = "Custom GCODE"
arrCommands(cID, lCurrentCommand) = i
arrCommands(cIDtree, lCurrentCommand) = strFeatIDtree
arrCommands(cNotes, lCurrentCommand) = strFeatIDrenumbered

strCODE = ""
For iSub = 1 To UBound(arrCustomGCODEtemp)
    If arrCustomGCODEtemp(iSub) <> "" Then
        strCODE = strCODE & arrCustomGCODEtemp(iSub)
        arrCommands(iSub + 1, lCurrentCommand) = arrCustomGCODEtemp(iSub)
    End If
Next iSub

arrCommands(cGCODE, lCurrentCommand) = strCODE
'arrCommands(cGCODE, lCurrentCommand) = arrFeatures(i, 2) & Format(CDbl(arrFeatures(i, 3)), "#####0.0#####") & arrFeatures(i, 4) & Format(CDbl(arrFeatures(i, 5)), "#####0.0#####") & arrFeatures(i, 6) & Format(CDbl(arrFeatures(i, 7)), "#####0.0#####") & arrFeatures(i, 8) & Format(CDbl(arrFeatures(i, 9)), "#####0.0#####") & arrFeatures(i, 10) & Format(CDbl(arrFeatures(i, 11)), "#####0.0#####")

lCurrentCommand = lCurrentCommand + 1
        
End Sub

Sub RepeatCustomGCODE(i, strFeatIDtree, strFeatIDrenumbered, arrCommands, lCurrentCommand, lCurrentCommandBeingRepeated)


Dim strCODE As String
'Add data to arrCommands
Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)

arrCommands(cCommandType, lCurrentCommand) = "Custom GCODE"
arrCommands(cID, lCurrentCommand) = i
arrCommands(cIDtree, lCurrentCommand) = strFeatIDtree
arrCommands(cNotes, lCurrentCommand) = strFeatIDrenumbered

strCODE = ""
For iSub = 2 To 11
    If arrCommands(iSub, lCurrentCommandBeingRepeated) <> "" Then
        arrCommands(iSub, lCurrentCommand) = arrCommands(iSub, lCurrentCommandBeingRepeated)
        strCODE = strCODE & arrCommands(iSub, lCurrentCommand)
    End If
Next iSub

arrCommands(cGCODE, lCurrentCommand) = strCODE
'arrCommands(cGCODE, lCurrentCommand) = arrFeatures(i, 2) & Format(CDbl(arrFeatures(i, 3)), "#####0.0#####") & arrFeatures(i, 4) & Format(CDbl(arrFeatures(i, 5)), "#####0.0#####") & arrFeatures(i, 6) & Format(CDbl(arrFeatures(i, 7)), "#####0.0#####") & arrFeatures(i, 8) & Format(CDbl(arrFeatures(i, 9)), "#####0.0#####") & arrFeatures(i, 10) & Format(CDbl(arrFeatures(i, 11)), "#####0.0#####")

lCurrentCommand = lCurrentCommand + 1
        
End Sub

Sub RotateLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXcentre, dblYcentre, dblAngleSubRadians, dblXnew1, dblYnew1, dblXnew2, dblYnew2)

Call RotatePoint(dblXold1, dblYold1, dblXcentre, dblYcentre, dblAngleSubRadians, dblXnew1, dblYnew1)
Call RotatePoint(dblXold2, dblYold2, dblXcentre, dblYcentre, dblAngleSubRadians, dblXnew2, dblYnew2)

End Sub


Sub RotatePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblAngleSubRadians, dblXnew, dblYnew)

Dim VectorOldX As Double
Dim VectorOldY As Double
Dim VectorNewX As Double
Dim VectorNewY As Double

VectorOldX = dblXold - dblXcentre
VectorOldY = dblYold - dblYcentre

VectorNewX = VectorOldX * Cos(dblAngleSubRadians) - VectorOldY * Sin(dblAngleSubRadians)
VectorNewY = VectorOldX * Sin(dblAngleSubRadians) + VectorOldY * Cos(dblAngleSubRadians)

dblXnew = dblXcentre + VectorNewX
dblYnew = dblYcentre + VectorNewY

End Sub

Sub RadiallyDisplaceLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXcentre, dblYcentre, dblRadialDisplacementSub, dblXnew1, dblYnew1, dblXnew2, dblYnew2)

Call RadiallyDisplacePoint(dblXold1, dblYold1, dblXcentre, dblYcentre, dblRadialDisplacementSub, dblXnew1, dblYnew1)
Call RadiallyDisplacePoint(dblXold2, dblYold2, dblXcentre, dblYcentre, dblRadialDisplacementSub, dblXnew2, dblYnew2)

End Sub
Sub RadiallyDisplacePoint(dblXold, dblYold, dblXcentre, dblYcentre, dblRadialDisplacementSub, dblXnew, dblYnew)

Dim VectorOldX As Double
Dim VectorOldY As Double
Dim VectorOldLength As Double
Dim VectorNewX As Double
Dim VectorNewY As Double
Dim VectorNewLength As Double

VectorOldX = dblXold - dblXcentre
VectorOldY = dblYold - dblYcentre
VectorOldLength = (VectorOldX ^ 2 + VectorOldY ^ 2) ^ 0.5
VectorNewLength = VectorOldLength + dblRadialDisplacementSub

If Not dblCheckTheSame(VectorOldX, 0, 10) Then
    VectorNewX = VectorOldX * VectorNewLength / VectorOldLength
End If

If Not dblCheckTheSame(VectorOldY, 0, 10) Then
    VectorNewY = VectorOldY * VectorNewLength / VectorOldLength
End If

dblXnew = dblXcentre + VectorNewX
dblYnew = dblYcentre + VectorNewY

End Sub

Sub ReflectLine(dblXold1, dblYold1, dblXold2, dblYold2, dblXreflect1, dblYreflect1, dblXreflect2, dblYreflect2, dblXnew1, dblYnew1, dblXnew2, dblYnew2)

Call ReflectPoint(dblXold1, dblYold1, dblXreflect1, dblYreflect1, dblXreflect2, dblYreflect2, dblXnew1, dblYnew1)
Call ReflectPoint(dblXold2, dblYold2, dblXreflect1, dblYreflect1, dblXreflect2, dblYreflect2, dblXnew2, dblYnew2)

End Sub

Sub ReflectPoint(dblXold, dblYold, dblXreflect1, dblYreflect1, dblXreflect2, dblYreflect2, dblXnew, dblYnew)

Dim dblM As Double
Dim dblMnormal As Double
Dim dblC As Double
Dim dblCnormal As Double
Dim dblXm As Double
Dim dblYm As Double

'Useful if switching to 3D planes in the future:
    'https://math.stackexchange.com/questions/753113/how-to-find-an-equation-of-the-plane-given-its-normal-vector-and-a-point-on-the
    'http://mathforum.org/library/drmath/view/54763.html
    
'Avoid numerical errors by checked for lines in X or Y direction only
If dblXreflect2 - dblXreflect1 = 0 Then 'reflection line in Y direction
    dblXnew = dblXold + 2 * (dblXreflect1 - dblXold)
    dblYnew = dblYold
ElseIf dblYreflect2 - dblYreflect1 = 0 Then 'reflection line in X direction
    dblXnew = dblXold
    dblYnew = dblYold + 2 * (dblYreflect1 - dblYold)
Else
    'First find the equation of the line
    dblM = (dblYreflect2 - dblYreflect1) / (dblXreflect2 - dblXreflect1)
    dblC = dblYreflect1 - (dblM * dblXreflect1)
    
    'Gradient of the line from the original point to the reflected point is:
    dblMnormal = -1 / dblM
    
    'Intersection of this line with the reflection line is:
    dblCnormal = dblYold - (dblMnormal * dblXold)
    
    'See my onenote note:
    dblXm = (dblCnormal - dblC) / (dblM - dblMnormal)
    dblXnew = dblXold + 2 * (dblXm - dblXold)
    dblYm = (dblCnormal - ((dblMnormal / dblM) * dblC)) / (1 - (dblMnormal / dblM))
    dblYnew = dblYold + 2 * (dblYm - dblYold)
End If


End Sub
Sub DetermineMathsOffset(strXequation, strYequation, strZequation, dblXnew, dblYnew, dblZnew, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)

strXequation = Replace(strXequation, "Xval", CStr(dblXnew))
strXequation = Replace(strXequation, "Yval", CStr(dblYnew))
strXequation = Replace(strXequation, "Zval", CStr(dblZnew))
strXequation = Replace(strXequation, "REPval", CStr(LayerCounterForCurrentRule))
dblLinearMathsOffsetX = Evaluate(strXequation)
strYequation = Replace(strYequation, "Xval", CStr(dblXnew))
strYequation = Replace(strYequation, "Yval", CStr(dblYnew))
strYequation = Replace(strYequation, "Zval", CStr(dblZnew))
strYequation = Replace(strYequation, "REPval", CStr(LayerCounterForCurrentRule))
dblLinearMathsOffsetY = Evaluate(strYequation)
strZequation = Replace(strZequation, "Xval", CStr(dblXnew))
strZequation = Replace(strZequation, "Yval", CStr(dblYnew))
strZequation = Replace(strZequation, "Zval", CStr(dblZnew))
strZequation = Replace(strZequation, "REPval", CStr(LayerCounterForCurrentRule))
dblLinearMathsOffsetZ = Evaluate(strZequation)

End Sub


Sub DetermineMathsOffsetPolar(dblXcentre, dblYcentre, strAngleEquationRadians, strRadiusEquation As String, strZequation, dblXnew1, dblYnew1, dblZnew, LayerCounterForCurrentRule, dblLinearMathsOffsetX, dblLinearMathsOffsetY, dblLinearMathsOffsetZ)


Dim dblVectorX As Double
Dim dblVectorY As Double

Dim dblPolarAngleRadians1 As Double
Dim dblRadius1 As Double

Dim dblPolarRotationRadians As Double
Dim dblRadialDisplacement As Double

Dim dblXnew_RotatedButNotYetRadiallyDisplaced As Double
Dim dblYnew_RotatedButNotYetRadiallyDisplaced As Double

Dim dblXnew_RotatedAndRadiallyDisplaced As Double
Dim dblYnew_RotatedAndRadiallyDisplaced As Double

'Find the polar angle and radius of the point
dblVectorX = dblXnew1 - dblXcentre
dblVectorY = dblYnew1 - dblYcentre
'If the point coincides with the polar centre point, it is impossible to calculate the polar angle so just set it to zero
If dblVectorX = 0 And dblVectorY = 0 Then
    dblPolarAngleRadians1 = 0
Else
    dblPolarAngleRadians1 = Application.Atan2(dblVectorX, dblVectorY)
End If

'Make the polar angle go from 0 to 360 instead of +/-180 - ALL IN RADIANS!!!
If dblPolarAngleRadians1 < 0 Then dblPolarAngleRadians1 = dblPolarAngleRadians1 + 2 * Pi
dblRadius1 = (((dblVectorX) ^ 2) + ((dblVectorY) ^ 2)) ^ 0.5

'Find the angular displacement for the point (in radians)
strAngleEquationRadians = Replace(strAngleEquationRadians, "Xval", CStr(dblXnew1))
strAngleEquationRadians = Replace(strAngleEquationRadians, "Yval", CStr(dblYnew1))
strAngleEquationRadians = Replace(strAngleEquationRadians, "Zval", CStr(dblZnew))
strAngleEquationRadians = Replace(strAngleEquationRadians, "Aval", CStr(dblPolarAngleRadians1)) 'Angle value
strAngleEquationRadians = Replace(strAngleEquationRadians, "Rval", CStr(dblRadius1)) 'Radius value
strAngleEquationRadians = Replace(strAngleEquationRadians, "REPval", CStr(LayerCounterForCurrentRule))
dblPolarRotationRadians = Evaluate(strAngleEquationRadians)
'Find the new X and Y cooardinates after angular displacement
Call RotatePoint(dblXnew1, dblYnew1, dblXcentre, dblYcentre, dblPolarRotationRadians, dblXnew_RotatedButNotYetRadiallyDisplaced, dblYnew_RotatedButNotYetRadiallyDisplaced)


'Find the radial displacement for the point
strRadiusEquation = Replace(strRadiusEquation, "Xval", CStr(dblXnew1))
strRadiusEquation = Replace(strRadiusEquation, "Yval", CStr(dblYnew1))
strRadiusEquation = Replace(strRadiusEquation, "Zval", CStr(dblZnew))
'strRadiusEquation = Replace(strRadiusEquation, "Aval", CStr(dblPolarAngleRadians1)) 'Angle value
strRadiusEquation = Replace(strRadiusEquation, "Aval", Format(CDbl(dblPolarAngleRadians1), "#####0.0########"))
strRadiusEquation = Replace(strRadiusEquation, "Rval", CStr(dblRadius1)) 'Radius value
strRadiusEquation = Replace(strRadiusEquation, "REPval", CStr(LayerCounterForCurrentRule))
'Range("O5").Value = "=" & strRadiusEquation
'dblRadialDisplacement = Range("O5").Value
dblRadialDisplacement = Evaluate(strRadiusEquation)
'Find the new X and Y cooardinates after radial displacement
Call RadiallyDisplacePoint(dblXnew_RotatedButNotYetRadiallyDisplaced, dblYnew_RotatedButNotYetRadiallyDisplaced, dblXcentre, dblYcentre, dblRadialDisplacement, dblXnew_RotatedAndRadiallyDisplaced, dblYnew_RotatedAndRadiallyDisplaced)



dblLinearMathsOffsetX = dblXnew_RotatedAndRadiallyDisplaced - dblXnew1
dblLinearMathsOffsetY = dblYnew_RotatedAndRadiallyDisplaced - dblYnew1

'Now find the Z displacement
strZequation = Replace(strZequation, "Xval", CStr(dblXnew1))
strZequation = Replace(strZequation, "Yval", CStr(dblYnew1))
strZequation = Replace(strZequation, "Zval", CStr(dblZnew))
strZequation = Replace(strZequation, "Aval", CStr(dblPolarAngleRadians1))
strZequation = Replace(strZequation, "Rval", CStr(dblRadius1))
strZequation = Replace(strZequation, "REPval", CStr(LayerCounterForCurrentRule))
dblLinearMathsOffsetZ = Evaluate(strZequation)


End Sub

Sub AddTravelAndToolChangeCommandsToWholeArray(lAddedCommandsCounter, lCurrentCommand, arrCommands, arrToolChangeGCODE, dblInitialX, dblInitialY, dblInitialZ, iInitialTool, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)

Dim arrCommandsTemp As Variant
Dim l_OldArrayCounter As Long
Dim l_TempArrayCounter As Long
Dim lNumberOfCommandsSub As Long
Dim lNumberOfParametersPerCommand As Long

Dim dblCurrentXsub As Double
Dim dblCurrentYsub As Double
Dim dblCurrentZsub As Double
Dim iCurrentToolNumberSub As Integer
Dim dblNextXsub As Double
Dim dblNextYsub As Double
Dim dblNextZsub As Double
Dim iNextToolNumberSub As Integer

lNumberOfParametersPerCommand = UBound(arrCommands, 1)
'Removed the following line of code because the matrix is actually too big (incresed in blocks of 10,000) and the black lines caused travel to be added unnecessarily
'lNumberOfCommandsSub = UBound(arrCommands, 2)
lNumberOfCommandsSub = lCurrentCommand - 1

ReDim arrCommandsTemp(1 To lNumberOfParametersPerCommand, 1 To lNumberOfCommandsSub)


lAddedCommandsCounter = 0
dblCurrentXsub = CDbl(dblInitialX)
dblCurrentYsub = dblInitialY
dblCurrentZsub = dblInitialZ
iCurrentToolNumberSub = iInitialTool
            
'Now copy the whole array accross and add travel/toolchanging as required
For l_OldArrayCounter = 1 To lNumberOfCommandsSub

    'if the current feature is a printed line or travel line, do all of the following:
    If Not arrCommands(1, l_OldArrayCounter) = "Retraction" _
    And Not arrCommands(1, l_OldArrayCounter) = "RetractionZhop" _
    And Not arrCommands(1, l_OldArrayCounter) = "Custom GCODE" _
    Then
        
        'Add travel/toolchange before the command if necessary
        
        'Set the "next" coordinates to be the beginning of this line. The "current" coordinates are set when the line is copied to the new array later in the subrouting, or are set to the initial start position after the start GCODE
        dblNextXsub = arrCommands(cX1, l_OldArrayCounter)
        dblNextYsub = arrCommands(cY1, l_OldArrayCounter)
        dblNextZsub = arrCommands(cZ1, l_OldArrayCounter)
        iNextToolNumberSub = arrCommands(cT, l_OldArrayCounter)
        
        'Set the tempArray counter to be the number of the current feature in the original array plus the number of new commands added
        l_TempArrayCounter = l_OldArrayCounter + lAddedCommandsCounter
        Call AddTravelAndChangeToolIfRequired(iNextToolNumberSub, iCurrentToolNumberSub, arrToolChangeGCODE, dblCurrentXsub, dblCurrentYsub, dblCurrentZsub, dblNextXsub, dblNextYsub, dblNextZsub, arrCommandsTemp, l_TempArrayCounter, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
        'Update the counter for added commands if new commands have been added
        lAddedCommandsCounter = l_TempArrayCounter - l_OldArrayCounter
    
        'Copy the current feature into the new array
        For iSub = 1 To lNumberOfParametersPerCommand
            Call IncreaseArrCommandsSizeIfRequired(arrCommandsTemp, l_OldArrayCounter + lAddedCommandsCounter)
            arrCommandsTemp(iSub, l_OldArrayCounter + lAddedCommandsCounter) = arrCommands(iSub, l_OldArrayCounter)
        Next iSub
        'Update the current coordinates and tool number to be the end of this line
        dblCurrentXsub = arrCommands(cX2, l_OldArrayCounter)
        dblCurrentYsub = arrCommands(cY2, l_OldArrayCounter)
        dblCurrentZsub = arrCommands(cZ2, l_OldArrayCounter)
        'Only update the toolnumber if it's not set to the "unset" value (which just takes the existing tool number)
        If arrCommands(cT, l_OldArrayCounter) <> -1 Then
            iCurrentToolNumberSub = arrCommands(cT, l_OldArrayCounter)
        End If
        
    Else
        'Otherwise, simply copy the feature across to the temporary array
        For iSub = 1 To lNumberOfParametersPerCommand
            Call IncreaseArrCommandsSizeIfRequired(arrCommandsTemp, l_OldArrayCounter + lAddedCommandsCounter)
            arrCommandsTemp(iSub, l_OldArrayCounter + lAddedCommandsCounter) = arrCommands(iSub, l_OldArrayCounter)
        Next iSub
    
    End If
Next l_OldArrayCounter

arrCommands = arrCommandsTemp

End Sub
Sub AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX As Double, dblCurrentY As Double, dblCurrentZ As Double, dblNextX As Double, dblNextY As Double, dblNextZ As Double, arrCommands, lCurrentCommandSub, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)
'Sub AddTravelAndChangeToolIfRequired(iToolNumber, iCurrentToolNumber, arrToolChangeGCODE, dblCurrentX As Double, dblCurrentY As Double, dblCurrentZ As Double, dblNextX As Double, dblNextY As Double, dblNextZ As Double, arrCommands, lCurrentCommandSub, FeatureNumberSub, dblTravelSpeed, strAutoRetractYesNo, dblAutoRetractThreshold, dblAutoRetractE, dblAutoRetractEspeed, dblAutoUnretractE, dblAutoUnretractEspeed, dblAutoRetractZhop, dblAutoRetractZhopSpeed)

Dim dblTravelDistanceSub As Double
Dim strCODE As String

'Only add travel if the tool ISN'T changed since the new tool will automatically travel to the correct position

'If tool number has changed (note that "-1" is the default value set for the tool if the user hasbn't specified it), change tool and move to the correct position
If iToolNumber <> iCurrentToolNumber And iToolNumber <> -1 Then
    'Add customGCODE feature for each line of code in the toolchange spreadsheet
    
    'The first row of the DEACTIVATE toolchange code is in row 13+22*toolNumber (of PREVIOUS tool - iCurrentToolNumber)
    For iSub = 13 + 22 * iCurrentToolNumber To 13 + 22 * iCurrentToolNumber + 9
        If arrToolChangeGCODE(iSub, 1) <> "" Then
            
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "ToolChange GCODE"
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "ToolChange GCODE added for feature number " & FeatureNumberSub
            arrCommands(cGCODE, lCurrentCommandSub) = arrToolChangeGCODE(iSub, 1)
            
            lCurrentCommandSub = lCurrentCommandSub + 1
            
            
        End If
    Next iSub
    
    'The first row of the ACTIVATE toolchange code is in row 2+22*toolNumber (of NEXT tool - iToolNumber)
    For iSub = 2 + 22 * iToolNumber To 2 + 22 * iToolNumber + 9
        If arrToolChangeGCODE(iSub, 1) <> "" Then
            
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "ToolChange GCODE"
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "ToolChange GCODE added for feature number " & FeatureNumberSub
            arrCommands(cGCODE, lCurrentCommandSub) = arrToolChangeGCODE(iSub, 1)
            
            lCurrentCommandSub = lCurrentCommandSub + 1
            
            
        End If
    Next iSub
    
    iCurrentToolNumber = iToolNumber
    
    'Reset the value of current X/Y/Z so that there is definitely a GCODE travel command to move the new nozzle to the right position (e.g. if the new line is printed directly from the end of the previous line, the current X/Y/Z values will say the nozzle is in the correct position, but actually it isn't.
    dblCurrentX = 0: dblCurrentY = 0: dblCurrentZ = 0
    
End If

'Add travel (if a ToolChange has happenned, this will be a travel movement for the new tool to the correct position)
If Not dblCheckSimilar(dblCurrentX, dblNextX, 0.001) Or Not dblCheckSimilar(dblCurrentY, dblNextY, 0.001) Or Not dblCheckSimilar(dblCurrentZ, dblNextZ, 0.001) Then

    dblTravelDistanceSub = ((dblCurrentX - dblNextX) ^ 2 + (dblCurrentY - dblNextY) ^ 2 + (dblCurrentZ - dblNextZ) ^ 2) ^ 0.5

    'If retraction is required (and the minimus travel threshold is met) add the reatraction, along with the travel line, unretraction, and Z-hop up/down (only if required)
    If (strAutoRetractYesNo = "yes" Or strAutoRetractYesNo = "Yes" Or strAutoRetractYesNo = "YES") _
    And dblTravelDistanceSub >= dblAutoRetractThreshold _
    Then
        
        'First add retraction (if required)
        If dblAutoRetractE <> 0 Then
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "Retraction"
            'Minus sign because +ive value means retract
            arrCommands(cRetractE, lCurrentCommandSub) = -dblAutoRetractE
            arrCommands(cRetractSpeed, lCurrentCommandSub) = dblAutoRetractEspeed
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "TravelRetraction automatically added"
            
            lCurrentCommandSub = lCurrentCommandSub + 1
        End If
        
        
        'Then add Z-hop up (if required)
        If dblAutoRetractZhop > 0 Then
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "RetractionZhop"
            arrCommands(cRetractZhop, lCurrentCommandSub) = dblAutoRetractZhop
            arrCommands(cRetractZhopSpeed, lCurrentCommandSub) = dblAutoRetractZhopSpeed
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "TravelZhopUp automatically added"
            
            lCurrentCommandSub = lCurrentCommandSub + 1
        End If
        
        
        'Then add the travel line
        Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
        arrCommands(cX1, lCurrentCommandSub) = dblCurrentX
        arrCommands(cY1, lCurrentCommandSub) = dblCurrentY
        'If z-hop, add the z-hop value to the travel line's z coordinate
        arrCommands(cZ1, lCurrentCommandSub) = dblCurrentZ + dblAutoRetractZhop
        arrCommands(cX2, lCurrentCommandSub) = dblNextX
        arrCommands(cY2, lCurrentCommandSub) = dblNextY
        'If z-hop, add the z-hop value to the travel line's z coordinate
        arrCommands(cZ2, lCurrentCommandSub) = dblNextZ + dblAutoRetractZhop
        arrCommands(cCommandType, lCurrentCommandSub) = "Travel"
        arrCommands(cW, lCurrentCommandSub) = 0
        arrCommands(cH, lCurrentCommandSub) = 0
        arrCommands(cF, lCurrentCommandSub) = dblTravelSpeed     'This is for AUTOMATICALLY ADDED travel
        arrCommands(cID, lCurrentCommandSub) = 0
        arrCommands(cIDtree, lCurrentCommandSub) = 0
        arrCommands(cNotes, lCurrentCommandSub) = "Travel automatically added"
        
        lCurrentCommandSub = lCurrentCommandSub + 1
        
        dblCurrentX = dblNextX
        dblCurrentY = dblNextY
        'No need to add the z-hop to this value since the Zhop (up) will be cancelled by the Z-hop (down) shortly
        dblCurrentZ = dblNextZ
        
        
        
        
        'Then add unretraction (if required)
        If dblAutoUnretractE <> 0 Then
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "Retraction"
            'Minus sign because +ive value means retract
            arrCommands(cRetractE, lCurrentCommandSub) = -dblAutoUnretractE
            arrCommands(cRetractSpeed, lCurrentCommandSub) = dblAutoUnretractEspeed
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "TravelUnretraction automatically added"
            
            lCurrentCommandSub = lCurrentCommandSub + 1
        End If

        
        'Then add Z-hop down (if required)
        If dblAutoRetractZhop > 0 Then
            Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
            arrCommands(cCommandType, lCurrentCommandSub) = "RetractionZhop"
            'Minus value to go down
            arrCommands(cRetractZhop, lCurrentCommandSub) = -dblAutoRetractZhop
            arrCommands(cRetractZhopSpeed, lCurrentCommandSub) = dblAutoRetractZhopSpeed
            arrCommands(cID, lCurrentCommandSub) = 0
            arrCommands(cIDtree, lCurrentCommandSub) = 0
            arrCommands(cNotes, lCurrentCommandSub) = "TravelZhopDown automatically added"
            
            lCurrentCommandSub = lCurrentCommandSub + 1
        End If
        
        
        
        
        
    'Else (if retraction not required), simply add the travel
    Else
        
        'Then add the travel line
        Call IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommandSub)
        arrCommands(cX1, lCurrentCommandSub) = dblCurrentX
        arrCommands(cY1, lCurrentCommandSub) = dblCurrentY
        arrCommands(cZ1, lCurrentCommandSub) = dblCurrentZ
        arrCommands(cX2, lCurrentCommandSub) = dblNextX
        arrCommands(cY2, lCurrentCommandSub) = dblNextY
        arrCommands(cZ2, lCurrentCommandSub) = dblNextZ
        arrCommands(cCommandType, lCurrentCommandSub) = "Travel"
        arrCommands(cW, lCurrentCommandSub) = 0
        arrCommands(cH, lCurrentCommandSub) = 0
        arrCommands(cF, lCurrentCommandSub) = dblTravelSpeed     'This is for AUTOMATICALLY ADDED travel
        arrCommands(cID, lCurrentCommandSub) = 0
        arrCommands(cIDtree, lCurrentCommandSub) = 0
        arrCommands(cNotes, lCurrentCommandSub) = "Travel automatically added"
        
        lCurrentCommandSub = lCurrentCommandSub + 1
        
        
        dblCurrentX = dblNextX
        dblCurrentY = dblNextY
        dblCurrentZ = dblNextZ
    End If
    
End If


End Sub
            
Sub GenerateGCODE(arrCommands, dblXoffset, dblYoffset, dblZoffset, dblLayer1EMultiplier, dblLayer1SpeedMultiplier, dblCurrentSpeed)

Dim iSub As Long
Dim strCODE As String
Dim dblXsub As Double
Dim dblYsub As Double
Dim dblZsub As Double
Dim dblEsub As Double
Dim dblFsub As Double

'Thsi is set to zero in case a retraction feature is called first (with Zhop)
dblZsub = 0 + dblZoffset

For iSub = 1 To UBound(arrCommands, 2)
    
    'First apply any first layer multipliers
    If CDbl(arrCommands(cZ2, iSub)) = 0 Then
        dblEsub = CDbl(arrCommands(cE, iSub)) * dblLayer1EMultiplier
    Else
        dblEsub = CDbl(arrCommands(cE, iSub))
    End If
    If CDbl(arrCommands(cZ2, iSub)) = 0 Then
        dblFsub = CDbl(arrCommands(cF, iSub)) * dblLayer1SpeedMultiplier
    Else
        dblFsub = CDbl(arrCommands(cF, iSub))
    End If
    

    'Then do the GCODE for retraction, or custom gecode, or line commands
    If arrCommands(cCommandType, iSub) = "Retraction" Then
    
        strCODE = "G1 F" & Format(CDbl(arrCommands(cRetractSpeed, iSub)), "#####0.0#####") & " E" & Format(CDbl(arrCommands(cRetractE, iSub)), "##0.0#####")
        dblCurrentSpeed = CDbl(arrCommands(cRetractSpeed, iSub))
    
    ElseIf arrCommands(cCommandType, iSub) = "RetractionZhop" Then
    
        dblZsub = dblZsub + CDbl(arrCommands(cRetractZhop, iSub))
        strCODE = "G0 F" & Format(CDbl(arrCommands(cRetractZhopSpeed, iSub)), "#####0.0#####") & " Z" & Format(dblZsub, "##0.0#####")
        dblCurrentSpeed = CDbl(arrCommands(cRetractZhopSpeed, iSub))
    
    
    ElseIf arrCommands(cCommandType, iSub) = "Custom GCODE" _
    Or arrCommands(cCommandType, iSub) = "ToolChange GCODE" Then
        
        'Do nothing because the GCODE has already been generated
        strCODE = arrCommands(cGCODE, iSub)
        'VBA_edit_001 start
        'reset speed since extrusion commands may have led to a slow speed being set. Setting it to "-1" means it will be identified as different to the speed of the next line of GCode, so an F-speed command will be written in the GCode.
        dblCurrentSpeed = -1
        'VBA_edit_001 end
    
    Else
    
        'Begin the GCODE string for a printed line command
        If arrCommands(cCommandType, iSub) = "Print" Then
            strCODE = "G1"
        Else
            strCODE = "G0"
        End If
            
        'If speed has changed, add it to the GCODE
        If Not dblCheckTheSame(CDbl(dblFsub), CDbl(dblCurrentSpeed), 10) Then
            strCODE = strCODE & " F" & Format(dblFsub, "###")
            dblCurrentSpeed = dblFsub
        End If
        'If X has changed, add it to the GCODE
        If Not dblCheckTheSame(CDbl(arrCommands(cX1, iSub)), CDbl(arrCommands(cX2, iSub)), 10) Then
            dblXsub = CDbl(arrCommands(cX2, iSub)) + dblXoffset
            strCODE = strCODE & " X" & Format(dblXsub, "##0.0##")
        End If
        'If Y has changed, add it to the GCODE
        If Not dblCheckTheSame(CDbl(arrCommands(cY1, iSub)), CDbl(arrCommands(cY2, iSub)), 10) Then
            dblYsub = CDbl(arrCommands(cY2, iSub)) + dblYoffset
            strCODE = strCODE & " Y" & Format(dblYsub, "##0.0##")
        End If
        'If Z has changed, add it to the GCODE
        If Not dblCheckTheSame(CDbl(arrCommands(cZ1, iSub)), CDbl(arrCommands(cZ2, iSub)), 10) Then
            dblZsub = CDbl(arrCommands(cZ2, iSub)) + dblZoffset
            strCODE = strCODE & " Z" & Format(dblZsub, "##0.0##")
        End If
        
        'If printing occurs, add the E value to the GCODE string
        If arrCommands(cCommandType, iSub) = "Print" Then
            'If any of X, Y or Z have changed, add an E term to the GCODE
            If Len(strCODE) > 2 Then
                strCODE = strCODE & " E" & Format(dblEsub, "##0.0#####")
            End If
        End If
    
    End If
    
    'if the is no printing or traveling happening, do not write "G1" or "G0" on their own.
    If strCODE = "G1" Or strCODE = "G0" Then
        strCODE = ""
    End If
        
    arrCommands(cGCODE, iSub) = strCODE
    
Next iSub

End Sub


Sub ReplaceDashesInString(strFeatureList As String)

Dim arrFeatureListSub As Variant
Dim arrFeatureListSubMini_1 As Variant
Dim arrFeatureListSubMini_2 As Variant
Dim lCurrentRangeLowValue As Long
Dim lCurrentRangeHighValue As Long

arrFeatureListSub = Split(strFeatureList, "-")

If UBound(arrFeatureListSub) > 0 Then
    'loop once for each dash
    For iSub = 1 To UBound(arrFeatureListSub)
        
        arrFeatureListSubMini_1 = Split(arrFeatureListSub(iSub - 1), ",")
        'array element 1: if there is a comma, find the number to the right of the last comma and use as start of range (this is the number that was orginially immediately before the dash).
        If UBound(arrFeatureListSubMini_1) > 0 Then
            lCurrentRangeLowValue = arrFeatureListSubMini_1(UBound(arrFeatureListSubMini_1))
        'else use full number and use as start of range.
        Else
            lCurrentRangeLowValue = arrFeatureListSubMini_1(0)
        End If
        
        arrFeatureListSubMini_2 = Split(arrFeatureListSub(iSub), ",")
        'array element 2: if there is a comma, find the number to the left of the first comma and use as end of range (this is the number that was orginially immediately after the dash).
        If UBound(arrFeatureListSubMini_2) > 0 Then
            lCurrentRangeHighValue = arrFeatureListSubMini_2(0)
        'else use full number and use as end of range.
        Else
            lCurrentRangeHighValue = arrFeatureListSubMini_2(0)
        End If
        
        
        'Add all the comma-separated values to the first array element (which will not be accessed in future iterations of this loop
        For jSub = lCurrentRangeLowValue + 1 To lCurrentRangeHighValue - 1
            
            arrFeatureListSub(iSub - 1) = arrFeatureListSub(iSub - 1) & "," & CStr(CLng(jSub))
            
        Next jSub
        
    Next iSub
    
    'Write the new full comma-separated string
    strFeatureList = ""
    For iSub = 0 To UBound(arrFeatureListSub)
        strFeatureList = strFeatureList & arrFeatureListSub(iSub)
        'Add a comma, except for the last time we run this loop
        If iSub < UBound(arrFeatureListSub) Then
            strFeatureList = strFeatureList & ","
        End If
    Next iSub
    
End If



End Sub

Sub DetermineAdditionalParams(strAdditionalParams, arrFeatures, i, dblE, dblFspeed, iToolNumber)

Dim arrAdditionalParamsSub As Variant
Dim arrFeatureListSubMini_1 As Variant
Dim arrFeatureListSubMini_2 As Variant
Dim lCurrentRangeLowValue As Long
Dim lCurrentRangeHighValue As Long

arrAdditionalParamsSub = Split(strAdditionalParams, ";")

'First set the values to be the default values
dblE = 0
dblFspeed = 0
iToolNumber = -1

If strAdditionalParams <> "" Then

    'First check for E valuess
    'loop once for each additional parameter written by the user
    For iSub = 0 To UBound(arrAdditionalParamsSub)
        'If the user has written "E=", set dblE to be the value after the bracket
        If Left(arrAdditionalParamsSub(iSub), 2) = "E=" Then
            dblE = Right(arrAdditionalParamsSub(iSub), Len(arrAdditionalParamsSub(iSub)) - 2)
        End If
    Next iSub
    
    'Then do the same for Fspeed
    For iSub = 0 To UBound(arrAdditionalParamsSub)
        If Left(arrAdditionalParamsSub(iSub), 2) = "F=" Then
            dblFspeed = Right(arrAdditionalParamsSub(iSub), Len(arrAdditionalParamsSub(iSub)) - 2)
        End If
    Next iSub
        
    'Then do the same for tool number
    For iSub = 0 To UBound(arrAdditionalParamsSub)
        If Left(arrAdditionalParamsSub(iSub), 2) = "T=" Then
            iToolNumber = Right(arrAdditionalParamsSub(iSub), Len(arrAdditionalParamsSub(iSub)) - 2)
        End If
    Next iSub
End If


End Sub

Sub ConvertRelativeToAbsoluteCoordintes(arrFeatures, i, dblCurrentX, dblCurrentY, dblCurrentZ)
    
    
    ''''''CONVERT RELATIVE COORDINATES TO ABSOLUTE FOR THIS FEATURE
    If arrFeatures(i, 1) = "Line" Then
        If arrFeatures(i, 2) = "Cartesian" Then
            'If X1, Y1, Z1 are relative, set the value to be the current value of X, Y, Z plus relative value
            If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentX + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
            If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentY + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
            If Left(arrFeatures(i, 5), 1) = "R" Then: arrFeatures(i, 5) = dblCurrentZ + CDbl(Right(arrFeatures(i, 5), Len(arrFeatures(i, 5)) - 1))
            'If X2, Y2, Z2 are relative, set the value X1, Y1, Z1 plus relative value
            If Left(arrFeatures(i, 6), 1) = "R" Then: arrFeatures(i, 6) = arrFeatures(i, 3) + CDbl(Right(arrFeatures(i, 6), Len(arrFeatures(i, 6)) - 1))
            If Left(arrFeatures(i, 7), 1) = "R" Then: arrFeatures(i, 7) = arrFeatures(i, 4) + CDbl(Right(arrFeatures(i, 7), Len(arrFeatures(i, 7)) - 1))
            If Left(arrFeatures(i, 8), 1) = "R" Then: arrFeatures(i, 8) = arrFeatures(i, 5) + CDbl(Right(arrFeatures(i, 8), Len(arrFeatures(i, 8)) - 1))
        ElseIf arrFeatures(i, 2) = "Polar" Then
            'If X_centre, Y_centre, Z1 are relative, set the value to be the current value of X, Y, Z plus relative value
            If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentX + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
            If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentY + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
            If Left(arrFeatures(i, 7), 1) = "R" Then: arrFeatures(i, 7) = dblCurrentZ + CDbl(Right(arrFeatures(i, 7), Len(arrFeatures(i, 7)) - 1))
            'If Z2 are relative, set the value Z1 plus relative value
            If Left(arrFeatures(i, 10), 1) = "R" Then: arrFeatures(i, 10) = arrFeatures(i, 7) + CDbl(Right(arrFeatures(i, 10), Len(arrFeatures(i, 10)) - 1))
        End If
    End If
    If arrFeatures(i, 1) = "Line equation" Then
        If Left(arrFeatures(i, 2), 1) = "R" And IsNumeric(Right(arrFeatures(i, 2), Len(arrFeatures(i, 2)) - 1)) Then: arrFeatures(i, 2) = dblCurrentX + CDbl(Right(arrFeatures(i, 2), Len(arrFeatures(i, 2)) - 1))
        If Left(arrFeatures(i, 3), 1) = "R" And IsNumeric(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1)) Then: arrFeatures(i, 3) = dblCurrentY + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
        If Left(arrFeatures(i, 4), 1) = "R" And IsNumeric(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1)) Then: arrFeatures(i, 4) = dblCurrentZ + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
    End If
    If arrFeatures(i, 1) = "Line equation polar" Then
        If Left(arrFeatures(i, 2), 1) = "R" Then: arrFeatures(i, 2) = dblCurrentX + CDbl(Right(arrFeatures(i, 2), Len(arrFeatures(i, 2)) - 1))
        If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentY + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
        If Left(arrFeatures(i, 6), 1) = "R" And IsNumeric(Right(arrFeatures(i, 6), Len(arrFeatures(i, 6)) - 1)) Then: arrFeatures(i, 6) = dblCurrentZ + CDbl(Right(arrFeatures(i, 6), Len(arrFeatures(i, 6)) - 1))
    End If
    If arrFeatures(i, 1) = "Rectangle" Then
        'If X1 and Y1 are relative, set the value to be the current value of X and Y plus relative value
        If Left(arrFeatures(i, 2), 1) = "R" Then: arrFeatures(i, 2) = dblCurrentX + CDbl(Right(arrFeatures(i, 2), Len(arrFeatures(i, 2)) - 1))
        If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentY + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
        'If X2 and Y2 are relative, set to the value X1 and Y1 plus relative value
        If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = arrFeatures(i, 2) + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
        If Left(arrFeatures(i, 5), 1) = "R" Then: arrFeatures(i, 5) = arrFeatures(i, 3) + CDbl(Right(arrFeatures(i, 5), Len(arrFeatures(i, 5)) - 1))
        'If Z is relative, set the value to be the current value of Z plus relative value
        If Left(arrFeatures(i, 6), 1) = "R" Then: arrFeatures(i, 6) = dblCurrentZ + CDbl(Right(arrFeatures(i, 6), Len(arrFeatures(i, 6)) - 1))
    End If
    If arrFeatures(i, 1) = "Reflect" Then
        If arrFeatures(i, 3) = "Polar" Then
            'If X_centre and Y_centre are relative, set the value to be the current value of X and Y plus relative value
            If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentX + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
            If Left(arrFeatures(i, 5), 1) = "R" Then: arrFeatures(i, 5) = dblCurrentY + CDbl(Right(arrFeatures(i, 5), Len(arrFeatures(i, 5)) - 1))
        ElseIf arrFeatures(i, 3) = "XY" Then
            'If X1, Y1 are relative, set the value to be the current value of X, Y plus relative value
            If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentX + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
            If Left(arrFeatures(i, 5), 1) = "R" Then: arrFeatures(i, 5) = dblCurrentY + CDbl(Right(arrFeatures(i, 5), Len(arrFeatures(i, 5)) - 1))
            'If X2, Y2 are relative, set the value X1, Y1 plus relative value
            If Left(arrFeatures(i, 6), 1) = "R" Then: arrFeatures(i, 6) = arrFeatures(i, 4) + CDbl(Right(arrFeatures(i, 6), Len(arrFeatures(i, 6)) - 1))
            If Left(arrFeatures(i, 7), 1) = "R" Then: arrFeatures(i, 7) = arrFeatures(i, 5) + CDbl(Right(arrFeatures(i, 7), Len(arrFeatures(i, 7)) - 1))
        ElseIf arrFeatures(i, 3) = "Z" Then
            'If X1, Y1 are relative, set the value to be the current value of X, Y plus relative value
            If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentZ + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
        End If
    End If
    If arrFeatures(i, 1) = "Polar repeat" Then
        'If Xcentre, Ycentre are relative, set the value to be the current value of X, Y plus relative value
        If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentX + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
        If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentY + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
    End If
    If arrFeatures(i, 1) = "Circle/arc" _
    Or arrFeatures(i, 1) = "Polygon" Then
        If Left(arrFeatures(i, 2), 1) = "R" Then: arrFeatures(i, 2) = dblCurrentX + CDbl(Right(arrFeatures(i, 2), Len(arrFeatures(i, 2)) - 1))
        If Left(arrFeatures(i, 3), 1) = "R" Then: arrFeatures(i, 3) = dblCurrentY + CDbl(Right(arrFeatures(i, 3), Len(arrFeatures(i, 3)) - 1))
        If Left(arrFeatures(i, 4), 1) = "R" Then: arrFeatures(i, 4) = dblCurrentZ + CDbl(Right(arrFeatures(i, 4), Len(arrFeatures(i, 4)) - 1))
    End If
    
End Sub
Sub IncreaseArrCommandsSizeIfRequired(arrCommands, lCurrentCommand)

If lCurrentCommand >= UBound(arrCommands, 2) Then
    'End the program if using too many commands (beyond limit in Excel)
    If 1 > 1038000 Then
        MsgBox "This design creates more than a million lines of GCODE (which is the upper limit for this version of FullControl GCODE Designer)" & vbNewLine & "Contact a.gleadall@lboro.ac.uk or info@FullControlGCODE.com to request software enhancements" & vbNewLine & vbNewLine & "The program will now end"
        End
    End If
    ReDim Preserve arrCommands(1 To UBound(arrCommands, 1), 1 To lCurrentCommand + 10000)
End If
 
End Sub


Function dblCheckSimilar(dbl1 As Double, dbl2 As Double, dblTolerance As Double)

If dbl1 >= dbl2 And (dbl1 - dbl2) < dblTolerance Then
    dblCheckSimilar = True
ElseIf dbl1 < dbl2 And (dbl2 - dbl1) < dblTolerance Then
    dblCheckSimilar = True
Else
    dblCheckSimilar = False
End If

End Function

Function dblCheckTheSame(number1 As Double, number2 As Double, Optional Digits As Integer = 12) As Boolean

If (number1 - number2) ^ 2 < (10 ^ -Digits) ^ 2 Then
    dblCheckTheSame = True
Else
    dblCheckTheSame = False
End If

End Function


Function CheckIfCurrentCommandSatisfiesInclusionCriteria(arrCommands, lCurrentCommandBeingRepeated, strFeatureList As String, strFeatIDtree As String) As Boolean

Dim arrFeaturesWrittenByUser As Variant
'Dim arrFeaturesExplicitlyIncluded
Dim arrFeaturesIN As Variant
Dim arrFeaturesOUT As Variant
Dim arrIDtreeParents As Variant
Dim strDashLeft As String
Dim strDashRight As String
Dim strFeatIDtemp As String


Dim bFeatIN As Boolean 'This is set to true if the current command has a parent feature listed as an inclusion criteria
bFeatIN = False
Dim bFeatOUT As Boolean 'This is set to true if the current command has a parent feature listed as an exclusion criteria
bFeatOUT = False

''''''THE FOLLOWING TWO LINES ARE SUPERCEDED BY THE VARIABLE strFeatIDtree BEING PASSED TO THE FUNCTION. THIS MEANS THE REPEATED FEATURE NUMBER IS INCLUDED RATHER THAN JUST THE NUMBERS OF THE FEATURES BEING REPEATED.
'''''Dim strFeatIDtree As String
'''''strFeatIDtree = arrCommands(cIDtree, lCurrentCommandBeingRepeated)

'If the user has written feature numbers (no "Y" or "N") to describe the features covered by the rule, do this...
If strFeatureList Like "*Y*" = False And strFeatureList Like "*N*" = False Then
    'Split the dashes in the list of features ONLY if the user is describing in terms of feature numbers, not parent inclusion/exclusion criteria
    Call ReplaceDashesInString(strFeatureList)
    arrFeaturesWrittenByUser = Split(strFeatureList, ",")
   
    For iSub = 0 To UBound(arrFeaturesWrittenByUser)
        'If the number is the same as the current command's feature number, exit this subroutine and say that the rule is included
        If CInt(arrFeaturesWrittenByUser(iSub)) = arrCommands(cID, lCurrentCommandBeingRepeated) Then
            CheckIfCurrentCommandSatisfiesInclusionCriteria = True
            Exit Function
        End If
    Next iSub

'Else the user has written boolean includsion/exclusion criteria for parent features with "Y" and "N" terminology
Else
    
    arrFeaturesWrittenByUser = Split(strFeatureList, ",")
    ReDim arrFeaturesIN(0 To UBound(arrFeaturesWrittenByUser))
    ReDim arrFeaturesOUT(0 To UBound(arrFeaturesWrittenByUser))

    For iSub = 0 To UBound(arrFeaturesWrittenByUser)
    
        'If there is a "Y" before the number, put the text to the right of the "Y" (the actual string of the parent feature) into the array of INCLUDED features, and put a "-1" in the array of EXCLUDED features
        If (Left(arrFeaturesWrittenByUser(iSub), 1) = "Y") Then
            arrFeaturesIN(iSub) = Right(arrFeaturesWrittenByUser(iSub), Len(arrFeaturesWrittenByUser(iSub)) - 1)
            arrFeaturesOUT(iSub) = -1 'just setting a number that will never be found since no feature has a number of "-1".
        
        'If there is a "N" before the number, put the text to the right of the "N" (the actual string of the parent feature) into the array of EXCLUDED features
        ElseIf (Left(arrFeaturesWrittenByUser(iSub), 1) = "N") Then
            arrFeaturesIN(iSub) = -1 'just setting a number that will never be found since no feature has a number of "-1".
            arrFeaturesOUT(iSub) = Right(arrFeaturesWrittenByUser(iSub), Len(arrFeaturesWrittenByUser(iSub)) - 1)
            
        End If
    
    Next iSub

    'The strFeatIDtree string (column 14 in the printpath worksheet) is written in a format that means feature numbers are written either at the very beginning of the string (before the first dash if there is a dash) or immediately after a dash.
    'For each parent feature, there may be a decimal point separater. The number after the decimal point is the repeat number. The user can search for this by writing "2.1" for eaxample. This function simply checks the string matches.
    arrIDtreeParents = Split(strFeatIDtree, "-")
    For iSub = 0 To UBound(arrFeaturesWrittenByUser)
        For jSub = 0 To UBound(arrIDtreeParents)
            
            
            'If the user has written a range, check the value is within the range.
            If arrFeaturesIN(iSub) Like "*-*" And arrFeaturesIN(iSub) <> "-1" Then
                
                'An issue is that "5.4" is within the range "5.38-5.42" when actually the user only wants 5.38,5.39,5.40,5.41,5.42.
                'Therefore, add zeros after the decimal point appropriately to ensure the user's input (after the decimal place) is treated correctly.
                
                'find the number written to the left of the dash
                strDashLeft = Split(arrFeaturesIN(iSub), "-")(0)
                If strDashLeft Like "*.*" Then
                    'add zeros after the decimal place until there re 6 digits to the right of the decimal place
                    Do While Len(Split(strDashLeft, ".")(1)) < 6
                        strDashLeft = Split(strDashLeft, ".")(0) & ".0" & Split(strDashLeft, ".")(1)
                    Loop
                End If
                
                'do the same for the string on the right of the dash
                strDashRight = Split(arrFeaturesIN(iSub), "-")(1)
                If strDashRight Like "*.*" Then
                    Do While Len(Split(strDashRight, ".")(1)) < 6
                        strDashRight = Split(strDashRight, ".")(0) & ".0" & Split(strDashRight, ".")(1)
                    Loop
                End If
                
                strFeatIDtemp = arrIDtreeParents(jSub)
                If strFeatIDtemp Like "*.*" Then
                    Do While Len(Split(strFeatIDtemp, ".")(1)) < 6
                        strFeatIDtemp = Split(strFeatIDtemp, ".")(0) & ".0" & Split(strFeatIDtemp, ".")(1)
                    Loop
                End If
                
                'Check if the current feature is within the range indicated by the user (the 0.0000000001 just avoids issues with comparisons of doubles).
                If CDbl(strFeatIDtemp) > CDbl(strDashLeft) - 0.0000000001 And CDbl(strFeatIDtemp) < CDbl(strDashRight) + 0.0000000001 Then
'                If arrIDtreeParents(jSub) > CDbl(Split(arrFeaturesIN(iSub), "-")(0)) - 0.0000000001 And arrIDtreeParents(jSub) < CDbl(Split(arrFeaturesIN(iSub), "-")(1)) + 0.0000000001 Then
                    bFeatIN = True
                End If
            'If the user has written a dot, check that the whole string matches (e.g. "5.2" but not "5.20" is found by the user writing "5.2")
            ElseIf arrFeaturesIN(iSub) Like "*.*" Then
                'the "len" bit means that if the user writes that feature 2 is included, then features 2.1, 2.2, 2.3, etc., are all included.
                If arrFeaturesIN(iSub) = arrIDtreeParents(jSub) Then
                    bFeatIN = True
                End If
            'Else just check than the beginning bit of the feature number is checked (e.g. "5.2" and "5.20" is found by the user writing "5")
            Else
                'the "len" bit means that if the user writes that feature 2 is included, then features 2.1, 2.2, 2.3, etc., are all included.
                If arrFeaturesIN(iSub) = Left(arrIDtreeParents(jSub), Len(arrFeaturesIN(iSub))) Then
                    bFeatIN = True
                End If
            End If
            
            'See comments about for arrFeaturesIN about this if statement
            If arrFeaturesOUT(iSub) Like "*-*" And arrFeaturesOUT(iSub) <> "-1" Then
            
                'find the number written to the left of the dash
                strDashLeft = Split(arrFeaturesOUT(iSub), "-")(0)
                If strDashLeft Like "*.*" Then
                    'add zeros after the decimal place until there re 6 digits to the right of the decimal place
                    Do While Len(Split(strDashLeft, ".")(1)) < 6
                        strDashLeft = Split(strDashLeft, ".")(0) & ".0" & Split(strDashLeft, ".")(1)
                    Loop
                End If
                
                'do the same for the string on the right of the dash
                strDashRight = Split(arrFeaturesOUT(iSub), "-")(1)
                If strDashRight Like "*.*" Then
                    Do While Len(Split(strDashRight, ".")(1)) < 6
                        strDashRight = Split(strDashRight, ".")(0) & ".0" & Split(strDashRight, ".")(1)
                    Loop
                End If
                
                strFeatIDtemp = arrIDtreeParents(jSub)
                If strFeatIDtemp Like "*.*" Then
                    Do While Len(Split(strFeatIDtemp, ".")(1)) < 6
                        strFeatIDtemp = Split(strFeatIDtemp, ".")(0) & ".0" & Split(strFeatIDtemp, ".")(1)
                    Loop
                End If
                
                'Check if the current feature is within the range indicated by the user (the 0.0000000001 just avoids issues with comparisons of doubles).
                If CDbl(strFeatIDtemp) > CDbl(strDashLeft) - 0.0000000001 And CDbl(strFeatIDtemp) < CDbl(strDashRight) + 0.0000000001 Then
                    bFeatOUT = True
                End If
            'See comments about for arrFeaturesIN about this if statement
            ElseIf arrFeaturesOUT(iSub) Like "*.*" Then
                If arrFeaturesOUT(iSub) = arrIDtreeParents(jSub) Then
                    bFeatOUT = True
                End If
            'See comments about for arrFeaturesIN about this if statement
            Else
                If arrFeaturesOUT(iSub) = Left(arrIDtreeParents(jSub), Len(arrFeaturesOUT(iSub))) Then
                    bFeatOUT = True
                End If
            End If
    
        Next jSub
    Next iSub
    
    
    'Only return a positive value if the
    If bFeatIN And Not bFeatOUT Then
        CheckIfCurrentCommandSatisfiesInclusionCriteria = True
    Else
        CheckIfCurrentCommandSatisfiesInclusionCriteria = False
    End If

End If

                                            
End Function
                                    

Function CheckIfCurrentCommandSatisfiesInclusionCriteria_QUICK(arrCommands, lCurrentCommandBeingRepeated, strFeatureList As String, strFeatIDtree As String) As Boolean
                           
            
Dim arrFeaturesWrittenByUser As Variant

Call ReplaceDashesInString(strFeatureList)
arrFeaturesWrittenByUser = Split(strFeatureList, ",")

For iSub = 0 To UBound(arrFeaturesWrittenByUser)
    'If the number is the same as the current command's feature number, exit this subroutine and say that the rule is included
    If CInt(arrFeaturesWrittenByUser(iSub)) = arrCommands(cID, lCurrentCommandBeingRepeated) Then
        CheckIfCurrentCommandSatisfiesInclusionCriteria_QUICK = True
        Exit Function
    End If
Next iSub

End Function




