Attribute VB_Name = "ReproduceFeature"
Sub UpdateCoordsOfReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, dblXdisp, dblYdisp, dblZdisp, iRep)
    
    'This code is in its own module because it is long and complicated and I didn't want it in the main module (GenerateModel)
    
    ''''''CONVERT RELATIVE COORDINATES TO ABSOLUTE FOR THIS FEATURE
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Line" Then
        If arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = "Cartesian" Then
            'If X1, Y1, Z1 are NOT relative, set the value to be the current value of X, Y, Z plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblYdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 5, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblYdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 5), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 5) = arrFeaturesTemp(iCurrentFeatureInNewLists, 5) + dblZdisp * iRep
            'If X2, Y2, Z2 are relative, set the value X1, Y1, Z1 plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 6, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 7, dblYdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 8, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 6), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = arrFeaturesTemp(iCurrentFeatureInNewLists, 6) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 7), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 7) = arrFeaturesTemp(iCurrentFeatureInNewLists, 7) + dblYdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 8), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 8) = arrFeaturesTemp(iCurrentFeatureInNewLists, 8) + dblZdisp * iRep
        ElseIf arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = "Polar" Then
            'If X_centre, Y_centre, Z1 are relative, set the value to be the current value plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblYdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 7, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblYdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 7), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 7) = arrFeaturesTemp(iCurrentFeatureInNewLists, 7) + dblZdisp * iRep
            'If Z2 is relative, set the value Z1 plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 10, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 10), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 10) = arrFeaturesTemp(iCurrentFeatureInNewLists, 10) + dblZdisp * iRep
        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Line equation" Then
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 2, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblYdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblZdisp * iRep)
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 2), 1) <> "R" Then
'            If IsNumeric(arrFeaturesTemp(iCurrentFeatureInNewLists, 2)) Then
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) + dblXdisp * iRep
'            Else
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) & "+" & dblXdisp * iRep
'            End If
'        End If
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then
'            If IsNumeric(arrFeaturesTemp(iCurrentFeatureInNewLists, 3)) Then
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblYdisp * iRep
'            Else
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) & "+" & dblYdisp * iRep
'            End If
'        End If
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then
'            If IsNumeric(arrFeaturesTemp(iCurrentFeatureInNewLists, 4)) Then
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblZdisp * iRep
'            Else
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) & "+" & dblZdisp * iRep
'            End If
'        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Line equation polar" Then
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 2, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblYdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 6, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 2), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblYdisp * iRep
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 6), 1) <> "R" Then
'            If IsNumeric(arrFeaturesTemp(iCurrentFeatureInNewLists, 6)) Then
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = arrFeaturesTemp(iCurrentFeatureInNewLists, 6) + dblZdisp * iRep
'            Else
'                arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = arrFeaturesTemp(iCurrentFeatureInNewLists, 6) & "+" & dblZdisp * iRep
'            End If
'        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Rectangle" Then
        'If X1 and Y1 are relative, set the value to be the current value of X and Y plus the offset
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 2, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 2), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblYdisp * iRep
        'If X2 and Y2 are relative, set to the value X1 and Y1 plus the offset
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 5, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 5), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 5) = arrFeaturesTemp(iCurrentFeatureInNewLists, 5) + dblYdisp * iRep
        'If Z is relative, set the value to be the current value of Z plus the offset
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 6, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 6), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = arrFeaturesTemp(iCurrentFeatureInNewLists, 6) + dblZdisp * iRep
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reflect" Then
        If arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = "Polar" Then
            'If X_centre and Y_centre are relative, set the value to be the current value of X and Y plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 5, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 5), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 5) = arrFeaturesTemp(iCurrentFeatureInNewLists, 5) + dblYdisp * iRep
        ElseIf arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = "XY" Then
            'If X1, Y1 are relative, set the value to be the current value of X, Y plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 5, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 5), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 5) = arrFeaturesTemp(iCurrentFeatureInNewLists, 5) + dblYdisp * iRep
            'If X2, Y2 are relative, set the value X1, Y1 plus relative value
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 6, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 7, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 6), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = arrFeaturesTemp(iCurrentFeatureInNewLists, 6) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 7), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 7) = arrFeaturesTemp(iCurrentFeatureInNewLists, 7) + dblYdisp * iRep
        ElseIf arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = "Z" Then
            'If X1, Y1 are relative, set the value to be the current value of X, Y plus the offset
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblZdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblZdisp * iRep
        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Polar repeat" Then
        'If Xcentre, Ycentre are relative, set the value to be the current value of X, Y plus the offset
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblYdisp * iRep)
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblXdisp * iRep
'            If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblYdisp * iRep
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Circle/arc" _
    Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Polygon" Then
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 2, dblXdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 3, dblYdisp * iRep)
        Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblZdisp * iRep)
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 2), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) + dblXdisp * iRep
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 3), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) + dblYdisp * iRep
'        If Left(arrFeaturesTemp(iCurrentFeatureInNewLists, 4), 1) <> "R" Then: arrFeaturesTemp(iCurrentFeatureInNewLists, 4) = arrFeaturesTemp(iCurrentFeatureInNewLists, 4) + dblZdisp * iRep
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Repeat rule" Then
        If arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = "OffsetPolar" _
        Or arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = "OffsetPolarIncrement" _
        Or arrFeaturesTemp(iCurrentFeatureInNewLists, 6) = "OffsetPolarMaths" _
        Then
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 7, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 8, dblYdisp * iRep)
        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Postprocess" Then
        If arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = "OffsetPolar" _
        Or arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = "OffsetPolarMaths" _
        Then
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 4, dblXdisp * iRep)
            Call OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, iCurrentFeatureInNewLists, 5, dblYdisp * iRep)
        End If
    End If
    If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reproduce and recalculate" Then
        'Do nothing, this is handled elsewhere
    End If

End Sub

Sub OffsetNumberInArrayIfNotRelative(arrFeaturesTemp, lRow, lColumn, dblOffsetValue)

If Left(arrFeaturesTemp(lRow, lColumn), 1) <> "R" Then
    If IsNumeric(arrFeaturesTemp(lRow, lColumn)) Then
        arrFeaturesTemp(lRow, lColumn) = arrFeaturesTemp(lRow, lColumn) + dblOffsetValue
    Else
        arrFeaturesTemp(lRow, lColumn) = arrFeaturesTemp(lRow, lColumn) & "+" & dblOffsetValue
    End If
End If

End Sub
                                    
Sub UpdateFeatNumbersOf_NOT_ReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, arrOriginalFeaturesNumbers)

Dim strFeatureListSub As String

'Replace the numbers written by the user with the new numbers of those features

If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Cartesian repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Polar repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reflect" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Concentric repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reproduce and recalculate" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Repeat rule" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Postprocess" _
Then
'    'DEBUGGING
'    If iCurrentFeatureInNewLists = 45 Then
'        MsgBox "Hi"
'    End If
'    'END DEBUGGING
    
    strFeatureListSub = arrFeaturesTemp(iCurrentFeatureInNewLists, 2)
    'If the use has written the numbers of features (they're not using "Y" and "N" inclusion/exclusion criteria) then do this stuff:
    If strFeatureListSub Like "*Y*" = False And strFeatureListSub Like "*N*" = False Then
    
        Call ReplaceDashesInString(strFeatureListSub)
        arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
        For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
            For jj = 1 To iCurrentFeatureInNewLists - 1
                If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = arrOriginalFeaturesNumbers(jj) Then
                    arrFeaturesWrittenByUser_REPRODUCED(j) = jj
                    'Since a "Reproduce and recalculate" feature adds lots of features that satisfy this condition, we need to write all of those features
                    'If the next feature in the list arrOriginalFeaturesNumbers(jj) has the same original feature number, that means it was generated by a "Reproduce and recalculate" feature so should also be included.
                    Do While arrOriginalFeaturesNumbers(jj + 1) = arrOriginalFeaturesNumbers(jj)
                        arrFeaturesWrittenByUser_REPRODUCED(j) = arrFeaturesWrittenByUser_REPRODUCED(j) & "," & jj + 1
                        jj = jj + 1
                    Loop
                    jj = iCurrentFeatureInNewLists - 1 + 1
                End If
            Next jj
        Next j
        arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
        If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
            For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
            Next j
        End If
      
    'Else the user has written in "Y" or "N" format
    Else
        'If the user has written a "-" or a ".", tell them that functionality is not programmed yet
        If strFeatureListSub Like "*-*" = True Or strFeatureListSub Like "*.*" = True Then
            MsgBox "I have not yet programmed the ability to use ""Reproduce and recalculate"" features if you're using ""Y""/""N"" AND ""-""/""."" notation for inclusion/exclusion criteria in Repeat rules, etc."
            End
        Else
            
    '        Call ReplaceDashesInString(strFeatureListSub)
            arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
            For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                For jj = 1 To iCurrentFeatureInNewLists - 1
                    'If the number to the right of the "Y" or "N" is...
                    If CInt(Right(arrFeaturesWrittenByUser_REPRODUCED(j), Len(arrFeaturesWrittenByUser_REPRODUCED(j)) - 1)) = arrOriginalFeaturesNumbers(jj) Then
                        'Set the new feature number but keep the "Y" or "N"...
                        arrFeaturesWrittenByUser_REPRODUCED(j) = Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj
                        'Since a "Reproduce and recalculate" feature adds lots of features that satisfy this condition, we need to write all of those features
                        'If the next feature in the list arrOriginalFeaturesNumbers(jj) has the same original feature number, that means it was generated by a "Reproduce and recalculate" feature so should also be included.
                        Do While arrOriginalFeaturesNumbers(jj + 1) = arrOriginalFeaturesNumbers(jj)
                            arrFeaturesWrittenByUser_REPRODUCED(j) = arrFeaturesWrittenByUser_REPRODUCED(j) & "," & Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj + 1
                            jj = jj + 1
                        Loop
                        jj = iCurrentFeatureInNewLists - 1 + 1
                    End If
                Next jj
            Next j
            arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
            If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
                For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                    arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
                Next j
            End If
                
            
        End If
        
    End If
    
End If

    
'Also do it again for Repeat rule (which has two lists of features)
If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Repeat rule" Then

    strFeatureListSub = arrFeaturesTemp(iCurrentFeatureInNewLists, 3)

    'If the use has written the numbers of features (they're not using "Y" and "N" inclusion/exclusion criteria) then do this stuff:
    If strFeatureListSub Like "*Y*" = False And strFeatureListSub Like "*N*" = False Then
    
        Call ReplaceDashesInString(strFeatureListSub)
        arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
        For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
            For jj = 1 To iCurrentFeatureInNewLists - 1
                If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = arrOriginalFeaturesNumbers(jj) Then
                    arrFeaturesWrittenByUser_REPRODUCED(j) = jj
                    'Since a "Reproduce and recalculate" feature adds lots of features that satisfy this condition, we need to write all of those features
                    'If the next feature in the list arrOriginalFeaturesNumbers(jj) has the same original feature number, that means it was generated by a "Reproduce and recalculate" feature so should also be included.
                    Do While arrOriginalFeaturesNumbers(jj + 1) = arrOriginalFeaturesNumbers(jj)
                        arrFeaturesWrittenByUser_REPRODUCED(j) = arrFeaturesWrittenByUser_REPRODUCED(j) & "," & jj + 1
                        jj = jj + 1
                    Loop
                    jj = iCurrentFeatureInNewLists - 1 + 1
                End If
            Next jj
        Next j
        arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
        If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
            For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
            Next j
        End If
    
    'Else the user has written in "Y" or "N" format
    Else
        'If the user has written a "-" or a ".", tell them that functionality is not programmed yet
        If strFeatureListSub Like "*-*" = True Or strFeatureListSub Like "*.*" = True Then
            MsgBox "I have not yet programmed the ability to use ""Reproduce and recalculate"" features if you're using ""Y""/""N"" notation for inclusion/exclusion criteria in Repeat rules, etc."
            End
        Else
        
            Call ReplaceDashesInString(strFeatureListSub)
            arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
            For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                For jj = 1 To iCurrentFeatureInNewLists - 1
                    'If the number to the right of the "Y" or "N" is...
                    If CInt(Right(arrFeaturesWrittenByUser_REPRODUCED(j), Len(arrFeaturesWrittenByUser_REPRODUCED(j)) - 1)) = arrOriginalFeaturesNumbers(jj) Then
                        'Set the new feature number but keep the "Y" or "N"...
                        arrFeaturesWrittenByUser_REPRODUCED(j) = Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj
                        'Since a "Reproduce and recalculate" feature adds lots of features that satisfy this condition, we need to write all of those features
                        'If the next feature in the list arrOriginalFeaturesNumbers(jj) has the same original feature number, that means it was generated by a "Reproduce and recalculate" feature so should also be included.
                        Do While arrOriginalFeaturesNumbers(jj + 1) = arrOriginalFeaturesNumbers(jj)
                            arrFeaturesWrittenByUser_REPRODUCED(j) = arrFeaturesWrittenByUser_REPRODUCED(j) & "," & Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj + 1
                            jj = jj + 1
                        Loop
                        jj = iCurrentFeatureInNewLists - 1 + 1
                    End If
                Next jj
            Next j
            arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
            If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
                For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                    arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
                Next j
            End If
        End If
    End If
        
End If

        
        
End Sub

Sub UpdateFeatNumbersOfReproducedFeature(arrFeaturesTemp, iCurrentFeatureInNewLists, iFirstFeatureForTheseRepeats, arrFeatureReproduced, arrNewFeaturesNumbers, iRep)

Dim strFeatureListSub As String

If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Cartesian repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Polar repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reflect" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Concentric repeat" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Reproduce and recalculate" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Repeat rule" _
Or arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Postprocess" _
Then
    strFeatureListSub = arrFeaturesTemp(iCurrentFeatureInNewLists, 2)
    'If the use has written the numbers of features (they're not using "Y" and "N" inclusion/exclusion criteria) then do this stuff:
    If strFeatureListSub Like "*Y*" = False And strFeatureListSub Like "*N*" = False Then
        Call ReplaceDashesInString(strFeatureListSub)
        arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
        For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
            For jj = iFirstFeatureForTheseRepeats To iCurrentFeatureInNewLists - 1
                If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = CInt(arrFeatureReproduced(jj)) Then
                    arrFeaturesWrittenByUser_REPRODUCED(j) = arrNewFeaturesNumbers(jj)
                End If
            Next jj
        Next j
        arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
        If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
            For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
            Next j
        End If

    'Else the user has written in "Y" or "N" format
    Else
    
        'If the user has written a "-" or a ".", tell them that functionality is not programmed yet
        If strFeatureListSub Like "*-*" = True Or strFeatureListSub Like "*.*" = True Then
            MsgBox "I have not yet programmed the ability to use ""Reproduce and recalculate"" features if you're using ""Y""/""N"" notation for inclusion/exclusion criteria in Repeat rules, etc."
            End
        Else
            Call ReplaceDashesInString(strFeatureListSub)
            arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
            For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                For jj = iFirstFeatureForTheseRepeats To iCurrentFeatureInNewLists - 1
                
                
                    'If the number to the right of the "Y" or "N" is...
                    If CInt(Right(arrFeaturesWrittenByUser_REPRODUCED(j), Len(arrFeaturesWrittenByUser_REPRODUCED(j)) - 1)) = CInt(arrFeatureReproduced(jj)) Then
                        'Set the new feature number but keep the "Y" or "N"...
                        arrFeaturesWrittenByUser_REPRODUCED(j) = Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj
'                    If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = CInt(arrFeatureReproduced(jj)) Then
'                        arrFeaturesWrittenByUser_REPRODUCED(j) = arrNewFeaturesNumbers(jj)
                    End If
                Next jj
            Next j
            arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
            If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
                For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                    arrFeaturesTemp(iCurrentFeatureInNewLists, 2) = arrFeaturesTemp(iCurrentFeatureInNewLists, 2) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
                Next j
            End If
        
        
        End If
    End If
    
    
End If


'Also do it again for Repeat rule (which has two lists of features)
If arrFeaturesTemp(iCurrentFeatureInNewLists, 1) = "Repeat rule" Then

    strFeatureListSub = arrFeaturesTemp(iCurrentFeatureInNewLists, 3)

    'If the use has written the numbers of features (they're not using "Y" and "N" inclusion/exclusion criteria) then do this stuff:
    If strFeatureListSub Like "*Y*" = False And strFeatureListSub Like "*N*" = False Then
        
        Call ReplaceDashesInString(strFeatureListSub)
        arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
        For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
            For jj = iFirstFeatureForTheseRepeats To iCurrentFeatureInNewLists - 1
                If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = CInt(arrFeatureReproduced(jj)) Then
                    arrFeaturesWrittenByUser_REPRODUCED(j) = arrNewFeaturesNumbers(jj)
                End If
            Next jj
        Next j
        arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
        If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
            For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
            Next j
        End If
        
      
    'Else the user has written in "Y" or "N" format
    Else
        
        'If the user has written a "-" or a ".", tell them that functionality is not programmed yet
        If strFeatureListSub Like "*-*" = True Or strFeatureListSub Like "*.*" = True Then
            MsgBox "I have not yet programmed the ability to use ""Reproduce and recalculate"" features if you're using ""Y""/""N"" notation for inclusion/exclusion criteria in Repeat rules, etc."
            End
        Else
        
            Call ReplaceDashesInString(strFeatureListSub)
            arrFeaturesWrittenByUser_REPRODUCED = Split(strFeatureListSub, ",")
            For j = 0 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                For jj = iFirstFeatureForTheseRepeats To iCurrentFeatureInNewLists - 1
                    'If the number to the right of the "Y" or "N" is...
                    If CInt(Right(arrFeaturesWrittenByUser_REPRODUCED(j), Len(arrFeaturesWrittenByUser_REPRODUCED(j)) - 1)) = CInt(arrFeatureReproduced(jj)) Then
                    'Set the new feature number but keep the "Y" or "N"...
                        arrFeaturesWrittenByUser_REPRODUCED(j) = Left(arrFeaturesWrittenByUser_REPRODUCED(j), 1) & jj
'                    If CInt(arrFeaturesWrittenByUser_REPRODUCED(j)) = CInt(arrFeatureReproduced(jj)) Then
'                        arrFeaturesWrittenByUser_REPRODUCED(j) = arrNewFeaturesNumbers(jj)
                    End If
                Next jj
            Next j
            arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = CStr(arrFeaturesWrittenByUser_REPRODUCED(0))
            If UBound(arrFeaturesWrittenByUser_REPRODUCED) > 0 Then
                For j = 1 To UBound(arrFeaturesWrittenByUser_REPRODUCED)
                    arrFeaturesTemp(iCurrentFeatureInNewLists, 3) = arrFeaturesTemp(iCurrentFeatureInNewLists, 3) & "," & CStr(arrFeaturesWrittenByUser_REPRODUCED(j))
                Next j
            End If
        End If
    End If
    
End If

End Sub
                



