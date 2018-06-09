Attribute VB_Name = "Cluster_Functions"
Option Explicit
' Jeff Jenness
' jeffj@jennessent.com
' Some functions hard-coded to work with specific datasets. All functions use ArcObjects and
' work best within an ArcMap VBA window.  Contact author if you can't figure
' out how to adapt these functions.


Public Sub PerformCluster_WithinSetDistance_March_12_2018()
  
  Dim booAddLayersToArcMap As Boolean
  booAddLayersToArcMap = True
  
  Dim dblThresholdDist As Double
  dblThresholdDist = 2
  
  Debug.Print "--------------------------"
  
  Dim lngClusterCount
  
  Dim dblMinX As Double
  Dim dblMaxX As Double
  Dim dblMinY As Double
  Dim dblMaxY As Double
  Dim dblMinZ As Double
  Dim dblMaxZ As Double
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Dim psbar As IStatusBar
  Dim pProg As IStepProgressor
  Set pApp = Application
  Set psbar = pApp.StatusBar
  Set pProg = psbar.ProgressBar
  
  Dim strLayerName As String
  Dim pFClass As IFeatureClass
  Dim pFLayer As IFeatureLayer
  
  
'  ' FOR DEBUGGING
'  Dim pNewPoint As IPoint
'  Dim pMarker1 As ISimpleMarkerSymbol
'  Dim pMarker2 As ISimpleMarkerSymbol
'  Dim pMarker3 As ISimpleMarkerSymbol
'
'  Dim pColor1 As IRgbColor
'  Dim pColor2 As IRgbColor
'  Dim pColor3 As IRgbColor
'
'  Set pMarker1 = New SimpleMarkerSymbol
'  Set pMarker2 = New SimpleMarkerSymbol
'  Set pMarker3 = New SimpleMarkerSymbol
'  Set pColor1 = New RgbColor
'  Set pColor2 = New RgbColor
'  Set pColor3 = New RgbColor
'
'  pColor1.RGB = RGB(255, 0, 0)
'  pColor2.RGB = RGB(0, 255, 0)
'  pColor3.RGB = RGB(0, 0, 255)
'
'  pMarker1.Color = pColor1
'  pMarker2.Color = pColor2
'  pMarker3.Color = pColor3
'
'  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  
  Dim strName As String
  
  Dim varSubsetNames() As Variant
  varSubsetNames = Array("XYPredictions_Day_GT_p01", "XYPredictions_Night_GT_p01", "XYPredictions_Diff_GT_p01")
  Dim varFullNames() As Variant
  varFullNames = Array("XYPredictions_Day", "XYPredictions_Night", "XYPredictions_Diff")
  
  Dim lngTotals() As Long
  Dim lngFractions() As Long
  Dim varXYZs() As Variant
  Dim varClusterCounts() As Variant
  Dim dblXYZ() As Double
  Dim dblMinMaxPHats() As Double
  Dim varPointSets() As Variant
  Dim dblCartExtremes() As Double
  
  Dim booUseSelected As Boolean
  Dim booIsProjected As Boolean
  Dim varOrigPoints() As Variant
  Dim dblDist As Double
  Dim dblMinPHat As Double
  Dim dblMaxPHat As Double
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPossibleIndex As Long
  Dim lngMaxIndex As Long
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblStartZ As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  Dim dblEndZ As Double
  
  Dim dblTempX As Double
  Dim dblTempY As Double
  Dim dblTempZ As Double
  
  Dim dblRunningCluster As Double
  Dim dblCurrentCluster As Double
  Dim lngTempIndices() As Long
  Dim lngPossibles() As Long
  Dim lngPossibleCounter As Long
  Dim pAlreadyFoundColl As Collection
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  
  ReDim lngTotals(2)
  ReDim lngFractions(2)
  ReDim varXYZs(2)
  ReDim varClusterCounts(2)
  ReDim varPointSets(2)
  ReDim dblMinMaxPHats(1, 2)
  ReDim dblCartExtremes(5, 2)
  
  Dim lngPopCount As Long
  Dim lngSampleCount As Long
  Dim varStatsToReturn() As Variant
  
  Dim lngArrayIndex As Long
  For lngArrayIndex = 0 To 2
    strName = varFullNames(lngArrayIndex)
    Set pFLayer = MyGeneralOperations.ReturnLayerByName(strName, pMxDoc.FocusMap)
    Set pFClass = pFLayer.FeatureClass
    Set pGeoDataset = pFClass
    Set pSpRef = pGeoDataset.SpatialReference
    
    lngTotals(lngArrayIndex) = pFClass.FeatureCount(Nothing)
    lngFractions(lngArrayIndex) = Round(CDbl(lngTotals(lngArrayIndex)) * 0.0001)
    
    Debug.Print strName & ": n = " & Format(lngFractions(lngArrayIndex), "#,##0")
      
    strName = varSubsetNames(lngArrayIndex)
    Set pFLayer = MyGeneralOperations.ReturnLayerByName(strName, pMxDoc.FocusMap)
    Set pFClass = pFLayer.FeatureClass
    
    dblXYZ = ReturnArrayOfXYZ_3(pFClass, lngFractions(lngArrayIndex), dblMinX, dblMaxX, dblMinY, _
        dblMaxY, dblMinZ, dblMaxZ, dblMinPHat, dblMaxPHat, booIsProjected, _
        varOrigPoints)
    lngMaxIndex = UBound(dblXYZ, 2)
        
    dblMinMaxPHats(0, lngArrayIndex) = dblMinPHat
    dblMinMaxPHats(1, lngArrayIndex) = dblMaxPHat
    varXYZs(lngArrayIndex) = dblXYZ
    varOrigPoints(lngArrayIndex) = varOrigPoints
    dblCartExtremes(0, lngArrayIndex) = dblMinX
    dblCartExtremes(1, lngArrayIndex) = dblMaxX
    dblCartExtremes(2, lngArrayIndex) = dblMinY
    dblCartExtremes(3, lngArrayIndex) = dblMaxY
    dblCartExtremes(4, lngArrayIndex) = dblMinZ
    dblCartExtremes(5, lngArrayIndex) = dblMaxZ
            
    Debug.Print "  --> pHats: " & Format(dblMinPHat, "0.00000") & " to " & _
        Format(dblMaxPHat, "0.00000")
        
    psbar.ShowProgressBar "Classifying into Clusters...", 0, lngMaxIndex, 1, True
    pProg.position = 0
    
    dblRunningCluster = 0
        
    ' FIRST CONVERT ALL TO CARTESIAN COORDINATES
    If Not booIsProjected Then
      For lngIndex1 = 0 To lngMaxIndex - 1
        dblTempX = dblXYZ(0, lngIndex1)
        dblTempY = dblXYZ(1, lngIndex1)
        dblTempZ = dblXYZ(2, lngIndex1)
        MyGeometricOperations.SpheroidalLatLongToCart dblTempX, dblTempY, dblStartX, dblStartY, _
            dblStartZ, , , dblTempZ
        dblXYZ(0, lngIndex1) = dblStartX
        dblXYZ(1, lngIndex1) = dblStartY
        dblXYZ(2, lngIndex1) = dblStartZ
      Next lngIndex1
    End If
    
    
    For lngIndex1 = 0 To lngMaxIndex - 1
      pProg.Step
      DoEvents
      
      ' CHECK IF ALREADY ASSIGNED TO CLUSTER
      If dblXYZ(3, lngIndex1) = 0 Then
        dblRunningCluster = dblRunningCluster + 1
        dblXYZ(3, lngIndex1) = dblRunningCluster
            
        dblStartX = dblXYZ(0, lngIndex1)
        dblStartY = dblXYZ(1, lngIndex1)
        dblStartZ = dblXYZ(2, lngIndex1)
        
        lngPossibleCounter = -1
        
  '      ' FOR DEBUGGING
  '      Set pNewPoint = New Point
  '      Set pNewPoint.SpatialReference = pSpRef
  '      pNewPoint.PutCoords dblStartX, dblStartY
  '      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker1
  '      ' --------------------------------------------
        
        For lngIndex2 = lngIndex1 + 1 To lngMaxIndex
        
          ' ONLY CONTINUE CHECKING IF NO CLUSTER ALREADY ASSIGNED
          If dblXYZ(3, lngIndex2) = 0 Then
            dblEndX = dblXYZ(0, lngIndex2)
            dblEndY = dblXYZ(1, lngIndex2)
            dblEndZ = dblXYZ(2, lngIndex2)
            
            dblDist = MyGeometricOperations.DistancePythagoreanNumbers_3D(dblStartX, dblStartY, _
                dblStartZ, dblEndX, dblEndY, dblEndZ)
            
            If dblDist <= dblThresholdDist Then
              dblXYZ(3, lngIndex2) = dblRunningCluster
              lngPossibleCounter = lngPossibleCounter + 1
              ReDim Preserve lngPossibles(lngPossibleCounter)
              lngPossibles(lngPossibleCounter) = lngIndex2
              
  '            ' FOR DEBUGGING
  '            Set pNewPoint = New Point
  '            Set pNewPoint.SpatialReference = pSpRef
  '            pNewPoint.PutCoords dblEndX, dblEndY
  '            MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker2
  '            ' --------------------------------------------
              
            End If
          End If
        Next lngIndex2
        
        ' NEXT, GO THROUGH ALL POINTS IN POSSIBLES LIST TO FIND NEW POINTS WITHIN DISTANCE
        Do Until lngPossibleCounter = -1
          lngPossibleCounter = -1
  '        Set pAlreadyFoundColl = New Collection
          
          lngTempIndices = lngPossibles
          Erase lngPossibles
          For lngPossibleIndex = 0 To UBound(lngTempIndices)
            dblStartX = dblXYZ(0, lngTempIndices(lngPossibleIndex))
            dblStartY = dblXYZ(1, lngTempIndices(lngPossibleIndex))
            dblStartZ = dblXYZ(2, lngTempIndices(lngPossibleIndex))
            
            ' GO THROUGH ALL POINTS GREATER THAN INDEX 1 AGAIN
            For lngIndex2 = lngIndex1 + 1 To lngMaxIndex
            
              ' ONLY CONTINUE CHECKING IF NO CLUSTER ALREADY ASSIGNED
              If dblXYZ(3, lngIndex2) = 0 Then
                
                dblEndX = dblXYZ(0, lngIndex2)
                dblEndY = dblXYZ(1, lngIndex2)
                dblEndZ = dblXYZ(2, lngIndex2)
                
                dblDist = MyGeometricOperations.DistancePythagoreanNumbers_3D(dblStartX, dblStartY, _
                    dblStartZ, dblEndX, dblEndY, dblEndZ)
                
                If dblDist <= dblThresholdDist Then
                  dblXYZ(3, lngIndex2) = dblRunningCluster
                  lngPossibleCounter = lngPossibleCounter + 1
                  ReDim Preserve lngPossibles(lngPossibleCounter)
                  lngPossibles(lngPossibleCounter) = lngIndex2
                              
  '                ' FOR DEBUGGING
  '                Set pNewPoint = New Point
  '                Set pNewPoint.SpatialReference = pSpRef
  '                pNewPoint.PutCoords dblEndX, dblEndY
  '                MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker3
  '                ' --------------------------------------------
                  
                End If
              End If
            Next lngIndex2
            
          Next lngPossibleIndex
            
  '        ' FOR DEBUGGING
  '        QuickSort.LongAscending lngPossibles, 0, UBound(lngPossibles)
  '        For lngPossibleCounter = 0 To UBound(lngPossibles)
  '          Debug.Print CStr(lngPossibleCounter) & "] " & CStr(lngPossibles(lngPossibleCounter))
  '        Next lngPossibleCounter
            
        Loop
        
      End If
    Next lngIndex1
      
    lngPopCount = lngTotals(lngArrayIndex)
    lngSampleCount = lngFractions(lngArrayIndex)
    dblMinPHat = dblMinMaxPHats(0, lngArrayIndex)
    dblMaxPHat = dblMinMaxPHats(1, lngArrayIndex)
    
    Debug.Print "lngPopCount = " & Format(lngPopCount, "#,##0")
    Debug.Print "lngSampleCount = " & Format(lngSampleCount, "#,##0")
    Debug.Print "dblMinPHat = " & Format(dblMinPHat, "0.000000")
    Debug.Print "dblMaxPHat = " & Format(dblMaxPHat, "0.000000")
    
    Call CreateInMem_Of_Points(dblXYZ, pSpRef, strName & "_Clusters", pMxDoc, dblThresholdDist, _
        varStatsToReturn, booAddLayersToArcMap)
    
    
'    varStatsToReturn = Array(UBound(dblXYZ, 2) + 1, lngNearCaveCount, lngNotNearCaveCount, _
'        UBound(varData, 2) + 1, lngInCount, lngTotalCount, CDbl(lngInCount) / CDbl(lngTotalCount))
    varClusterCounts(lngArrayIndex) = varStatsToReturn
    
  Next lngArrayIndex
  
  Dim strReturn As String
  Dim varStatNames() As Variant
  varStatNames = Array("Total Cell Count", "Subset Cell Count (Top 0.01%)", "Minimum pHat in Subset", _
      "Maximum pHat in Subset", "Cell Cluster Count in Subset", "Clusters Near Known Caves", _
      "Clusters Not Near Known Caves", "Proportion of Cells Near Known Caves")
  
  strReturn = "Statistic" & vbTab & "Predawn" & vbTab & "Midday" & vbTab & "Difference" & vbCrLf
  
  Dim lngIndex As Long
  
  For lngIndex = 0 To 7
    For lngIndex2 = 0 To 3
      Select Case lngIndex2
        Case 0
          lngArrayIndex = -999
        Case 1 ' predawn; Index 1
          lngArrayIndex = 1
        Case 2 ' midday; Index 0
          lngArrayIndex = 0
        Case 3 ' Difference; Index 2
          lngArrayIndex = 2
      End Select
      
      If lngIndex2 = 0 Then
        strReturn = strReturn & CStr(varStatNames(lngIndex)) & vbTab
      Else
        varStatsToReturn = varClusterCounts(lngArrayIndex)
        lngPopCount = lngTotals(lngArrayIndex)
        lngSampleCount = lngFractions(lngArrayIndex)
        dblMinPHat = dblMinMaxPHats(0, lngArrayIndex)
        dblMaxPHat = dblMinMaxPHats(1, lngArrayIndex)
        
        Select Case lngIndex
          Case 0  ' Population Count
            strReturn = strReturn & Format(lngPopCount, "#,##0") & vbTab
          Case 1  ' Sample Count
            strReturn = strReturn & Format(lngSampleCount, "#,##0") & vbTab
          Case 2  ' Minimum pHat
            strReturn = strReturn & Format(dblMinPHat, "0.000000") & vbTab
          Case 3  ' Maximum pHat
            strReturn = strReturn & Format(dblMaxPHat, "0.000000") & vbTab
          Case 4  ' Cluster Count
            strReturn = strReturn & Format(varStatsToReturn(3), "#,##0") & vbTab
          Case 5  ' Clusters Near Caves
            strReturn = strReturn & Format(varStatsToReturn(1), "#,##0") & vbTab
          Case 6  ' Clusters Not Near Caves
            strReturn = strReturn & Format(varStatsToReturn(2), "#,##0") & vbTab
          Case 7  ' Proportion Cells Near Caves
            strReturn = strReturn & Format(varStatsToReturn(6), "0%") & _
                " [" & Format(varStatsToReturn(4), "0") & " of " & _
                Format(lngSampleCount) & "]" & vbTab
        End Select
        
        If lngIndex2 = 3 Then strReturn = Left(strReturn, Len(strReturn) - 1) & vbCrLf
      End If
    Next lngIndex2
  Next lngIndex
  
  psbar.HideProgressBar
  pProg.position = 0
  
  Dim pDataObj As New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReturn
  pDataObj.PutInClipboard
  Set pDataObj = Nothing
  
  Debug.Print strReturn
  
  Debug.Print "Done..."
'  Debug.Print CStr(dblRunningCluster) & " clusters..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set psbar = Nothing
  Set pProg = Nothing
  Set pFClass = Nothing
  Set pFLayer = Nothing
  Erase varSubsetNames
  Erase varFullNames
  Erase lngTotals
  Erase lngFractions
  Erase varXYZs
  Erase varClusterCounts
  Erase dblXYZ
  Erase dblMinMaxPHats
  Erase varPointSets
  Erase dblCartExtremes
  Erase varOrigPoints
  Erase lngTempIndices
  Erase lngPossibles
  Set pAlreadyFoundColl = Nothing
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing




End Sub



Public Sub ExtractHighPHatVals()
  
  Dim dblThreshold As Double
  dblThreshold = 0.01
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Dim psbar As IStatusBar
  Dim pProg As IStepProgressor
  Set pApp = Application
  Set psbar = pApp.StatusBar
  Set pProg = psbar.ProgressBar
  
  Dim strLayerName As String
  Dim pDayFClass As IFeatureClass
  Dim pNightFClass As IFeatureClass
  Dim pDiffFClass As IFeatureClass
  
  strLayerName = "Cave_like_points"
  Dim pFLayer As IFeatureLayer
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("XYPredictions_Day", pMxDoc.FocusMap)
  Set pDayFClass = pFLayer.FeatureClass
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("XYPredictions_Night", pMxDoc.FocusMap)
  Set pNightFClass = pFLayer.FeatureClass
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("XYPredictions_Diff", pMxDoc.FocusMap)
  Set pDiffFClass = pFLayer.FeatureClass
  
  Dim pDataset As IDataset
  Dim pWS As IWorkspace
  Set pDataset = pDayFClass
  Set pWS = pDataset.Workspace
  
  Dim strDayName As String
  Dim varDayIndexes() As Variant
  Dim strNightName As String
  Dim varNightIndexes() As Variant
  Dim strDifName As String
  Dim varDiffIndexes() As Variant
  
  Dim strThreshold As String
  strThreshold = Format(dblThreshold * 100, "00")
  
  strDayName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS, _
      "XYPredictions_Day_GT_p" & strThreshold)
  strNightName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS, _
      "XYPredictions_Night_GT_p" & strThreshold)
  strDifName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS, _
      "XYPredictions_Diff_GT_p" & strThreshold)
  
  Dim lngCount As Long
  Dim lngCounter As Long
  
  lngCount = pDayFClass.FeatureCount(Nothing) + pNightFClass.FeatureCount(Nothing) + _
      pDiffFClass.FeatureCount(Nothing)
  lngCounter = 0
  
  Dim pNewDayFClass As IFeatureClass
  Dim pNewNightFClass As IFeatureClass
  Dim pNewDiffFClass As IFeatureClass
  
  Set pNewDayFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pDayFClass, pWS, varDayIndexes, _
      strDayName, True)
  Set pNewNightFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pNightFClass, pWS, varNightIndexes, _
      strNightName, True)
  Set pNewDiffFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pDiffFClass, pWS, varDiffIndexes, _
      strDifName, True)
      
  psbar.ShowProgressBar "Working on '" & strDayName & "'...", 0, lngCount, 1, True
  pProg.position = 0
  
  Dim lngFieldIndex As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pNewFCursor As IFeatureCursor
  Dim pNewBuffer As IFeatureBuffer
  Dim lngpHatIndex As Long
  Dim dblpHat As Double
  Dim pNewFLayer As IFeatureLayer
  Dim lngFClassCount As Long
  Dim strFClassCount As Long
  
  ' DAY
  Set pFCursor = pDayFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Set pNewFCursor = pNewDayFClass.Insert(True)
  Set pNewBuffer = pNewDayFClass.CreateFeatureBuffer
  lngpHatIndex = pDayFClass.FindField("phat")
  lngFClassCount = pDayFClass.FeatureCount(Nothing)
  strFClassCount = Format(lngFClassCount, "#,##0")
  
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      pNewFCursor.Flush
      DoEvents
      Debug.Print "--> " & Format(lngCounter, "#,##0") & " of " & Format(lngFClassCount, "#,##0")
    End If
    dblpHat = pFeature.Value(lngpHatIndex)
    
    If dblpHat >= dblThreshold Then
      Set pNewBuffer.Shape = pFeature.ShapeCopy
      For lngFieldIndex = 0 To UBound(varDayIndexes, 2)
        pNewBuffer.Value(varDayIndexes(3, lngFieldIndex)) = pFeature.Value(varDayIndexes(1, lngFieldIndex))
      Next lngFieldIndex
      pNewFCursor.InsertFeature pNewBuffer
    End If
    Set pFeature = pFCursor.NextFeature
  Loop
  pNewFCursor.Flush
  Set pDataset = pNewDayFClass
  Set pNewFLayer = New FeatureLayer
  Set pNewFLayer.FeatureClass = pNewDayFClass
  pNewFLayer.Name = pDataset.BrowseName
  pMxDoc.FocusMap.AddLayer pNewFLayer
  
  ' Night
  Set pFCursor = pNightFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Set pNewFCursor = pNewNightFClass.Insert(True)
  Set pNewBuffer = pNewNightFClass.CreateFeatureBuffer
  lngpHatIndex = pNightFClass.FindField("phat")
  lngFClassCount = pNightFClass.FeatureCount(Nothing)
  strFClassCount = Format(lngFClassCount, "#,##0")
  
  lngCounter = 0
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      pNewFCursor.Flush
      DoEvents
      Debug.Print "--> " & Format(lngCounter, "#,##0") & " of " & Format(lngFClassCount, "#,##0")
    End If
    dblpHat = pFeature.Value(lngpHatIndex)
    
    If dblpHat >= dblThreshold Then
      Set pNewBuffer.Shape = pFeature.ShapeCopy
      For lngFieldIndex = 0 To UBound(varNightIndexes, 2)
        pNewBuffer.Value(varNightIndexes(3, lngFieldIndex)) = pFeature.Value(varNightIndexes(1, lngFieldIndex))
      Next lngFieldIndex
      pNewFCursor.InsertFeature pNewBuffer
    End If
    Set pFeature = pFCursor.NextFeature
  Loop
  pNewFCursor.Flush
  Set pDataset = pNewNightFClass
  Set pNewFLayer = New FeatureLayer
  Set pNewFLayer.FeatureClass = pNewNightFClass
  pNewFLayer.Name = pDataset.BrowseName
  pMxDoc.FocusMap.AddLayer pNewFLayer
  
  
  ' Difference
  Set pFCursor = pDiffFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Set pNewFCursor = pNewDiffFClass.Insert(True)
  Set pNewBuffer = pNewDiffFClass.CreateFeatureBuffer
  lngpHatIndex = pDiffFClass.FindField("phat")
  lngFClassCount = pDiffFClass.FeatureCount(Nothing)
  strFClassCount = Format(lngFClassCount, "#,##0")
  
  lngCounter = 0
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 1000 = 0 Then
      pNewFCursor.Flush
      DoEvents
      Debug.Print "--> " & Format(lngCounter, "#,##0") & " of " & Format(lngFClassCount, "#,##0")
    End If
    dblpHat = pFeature.Value(lngpHatIndex)
    
    If dblpHat >= dblThreshold Then
      Set pNewBuffer.Shape = pFeature.ShapeCopy
      For lngFieldIndex = 0 To UBound(varDiffIndexes, 2)
        pNewBuffer.Value(varDiffIndexes(3, lngFieldIndex)) = pFeature.Value(varDiffIndexes(1, lngFieldIndex))
      Next lngFieldIndex
      pNewFCursor.InsertFeature pNewBuffer
    End If
    Set pFeature = pFCursor.NextFeature
  Loop
  pNewFCursor.Flush
  Set pDataset = pNewDiffFClass
  Set pNewFLayer = New FeatureLayer
  Set pNewFLayer.FeatureClass = pNewDiffFClass
  pNewFLayer.Name = pDataset.BrowseName
  pMxDoc.FocusMap.AddLayer pNewFLayer
  
  
  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh
  
  psbar.HideProgressBar
  pProg.position = 0
  
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set psbar = Nothing
  Set pProg = Nothing
  Set pDayFClass = Nothing
  Set pNightFClass = Nothing
  Set pDiffFClass = Nothing
  Set pFLayer = Nothing
  Set pDataset = Nothing
  Set pWS = Nothing
  Erase varDayIndexes
  Erase varNightIndexes
  Erase varDiffIndexes
  Set pNewDayFClass = Nothing
  Set pNewNightFClass = Nothing
  Set pNewDiffFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pNewFCursor = Nothing
  Set pNewBuffer = Nothing
  Set pNewFLayer = Nothing



End Sub

Public Sub PerformCluster_WithinSetDistance()
  
  Debug.Print "--------------------------"
  
  Dim strName As String
  strName = "All Points"
  strName = "Day Points"
'  strName = "Difference Points"
  
  Dim dblThresholdDist As Double
  dblThresholdDist = 2
    
  Dim lngClusterCount
  
  Dim dblMinX As Double
  Dim dblMaxX As Double
  Dim dblMinY As Double
  Dim dblMaxY As Double
  Dim dblMinZ As Double
  Dim dblMaxZ As Double
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Dim psbar As IStatusBar
  Dim pProg As IStepProgressor
  Set pApp = Application
  Set psbar = pApp.StatusBar
  Set pProg = psbar.ProgressBar
  
  Dim strLayerName As String
  Dim pFClass As IFeatureClass
  strLayerName = "Cave_like_points"
  Dim pFLayer As IFeatureLayer
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName(strLayerName, pMxDoc.FocusMap)
  Set pFClass = pFLayer.FeatureClass
  
  Debug.Print "Reading X, Y and Z Values..."
  Dim dblXYZ() As Double
  Dim varAttNames() As Variant
  varAttNames = Array("Time_")
  
  Dim varAttributes() As Variant
  Dim pAttFields As esriSystem.IVariantArray
  Dim booUseSelected As Boolean
  Dim booIsProjected As Boolean
  Dim varOrigPoints() As Variant
  Dim dblDist As Double
  
  Dim strQueryString As String
  Dim strPrefix As String
  Dim strSuffix As String
  
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
    
  Select Case strName
    Case "All Points"
      booUseSelected = False
      strQueryString = ""
    Case "Day Points"
      booUseSelected = True
      strQueryString = strPrefix & "Time_" & strSuffix & " = 'Day'"
    Case "Difference Points"
      booUseSelected = True
      strQueryString = strPrefix & "Time_" & strSuffix & " = 'Diff'"
    Case "Night Points"
      booUseSelected = True
      strQueryString = ""
    Case Else
      MsgBox "Problem!"
  End Select
  
  dblXYZ = ReturnArrayOfXYZ_2(strLayerName, pMxDoc, pFClass, dblMinX, dblMaxX, dblMinY, _
      dblMaxY, dblMinZ, dblMaxZ, varAttNames, varAttributes, booUseSelected, booIsProjected, _
      varOrigPoints, strQueryString)
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPossibleIndex As Long
  Dim lngMaxIndex As Long
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblStartZ As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  Dim dblEndZ As Double
  
  Dim dblTempX As Double
  Dim dblTempY As Double
  Dim dblTempZ As Double
  
  Dim dblRunningCluster As Double
  Dim dblCurrentCluster As Double
  Dim lngTempIndices() As Long
  Dim lngPossibles() As Long
  Dim lngPossibleCounter As Long
  Dim pAlreadyFoundColl As Collection
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  
  lngMaxIndex = UBound(dblXYZ, 2)
      
  psbar.ShowProgressBar "Classifying into Clusters...", 0, lngMaxIndex, 1, True
  pProg.position = 0
  
  dblRunningCluster = 0
  
  
'  ' FOR DEBUGGING
'  Dim pNewPoint As IPoint
'  Dim pMarker1 As ISimpleMarkerSymbol
'  Dim pMarker2 As ISimpleMarkerSymbol
'  Dim pMarker3 As ISimpleMarkerSymbol
'
'  Dim pColor1 As IRgbColor
'  Dim pColor2 As IRgbColor
'  Dim pColor3 As IRgbColor
'
'  Set pMarker1 = New SimpleMarkerSymbol
'  Set pMarker2 = New SimpleMarkerSymbol
'  Set pMarker3 = New SimpleMarkerSymbol
'  Set pColor1 = New RgbColor
'  Set pColor2 = New RgbColor
'  Set pColor3 = New RgbColor
'
'  pColor1.RGB = RGB(255, 0, 0)
'  pColor2.RGB = RGB(0, 255, 0)
'  pColor3.RGB = RGB(0, 0, 255)
'
'  pMarker1.Color = pColor1
'  pMarker2.Color = pColor2
'  pMarker3.Color = pColor3
'
'  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
'  ' --------------------------------------------
  
  ' FIRST CONVERT ALL TO CARTESIAN COORDINATES
  If Not booIsProjected Then
    For lngIndex1 = 0 To lngMaxIndex - 1
      dblTempX = dblXYZ(0, lngIndex1)
      dblTempY = dblXYZ(1, lngIndex1)
      dblTempZ = dblXYZ(2, lngIndex1)
      MyGeometricOperations.SpheroidalLatLongToCart dblTempX, dblTempY, dblStartX, dblStartY, _
          dblStartZ, , , dblTempZ
      dblXYZ(0, lngIndex1) = dblStartX
      dblXYZ(1, lngIndex1) = dblStartY
      dblXYZ(2, lngIndex1) = dblStartZ
    Next lngIndex1
  End If
  
  For lngIndex1 = 0 To lngMaxIndex - 1
    pProg.Step
    DoEvents
    
    ' CHECK IF ALREADY ASSIGNED TO CLUSTER
    If dblXYZ(3, lngIndex1) = 0 Then
      dblRunningCluster = dblRunningCluster + 1
      dblXYZ(3, lngIndex1) = dblRunningCluster
          
      dblStartX = dblXYZ(0, lngIndex1)
      dblStartY = dblXYZ(1, lngIndex1)
      dblStartZ = dblXYZ(2, lngIndex1)
      
      lngPossibleCounter = -1
      
'      ' FOR DEBUGGING
'      Set pNewPoint = New Point
'      Set pNewPoint.SpatialReference = pSpRef
'      pNewPoint.PutCoords dblStartX, dblStartY
'      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker1
'      ' --------------------------------------------
      
      For lngIndex2 = lngIndex1 + 1 To lngMaxIndex
      
        ' ONLY CONTINUE CHECKING IF NO CLUSTER ALREADY ASSIGNED
        If dblXYZ(3, lngIndex2) = 0 Then
          dblEndX = dblXYZ(0, lngIndex2)
          dblEndY = dblXYZ(1, lngIndex2)
          dblEndZ = dblXYZ(2, lngIndex2)
          
          dblDist = MyGeometricOperations.DistancePythagoreanNumbers_3D(dblStartX, dblStartY, _
              dblStartZ, dblEndX, dblEndY, dblEndZ)
          
          If dblDist <= dblThresholdDist Then
            dblXYZ(3, lngIndex2) = dblRunningCluster
            lngPossibleCounter = lngPossibleCounter + 1
            ReDim Preserve lngPossibles(lngPossibleCounter)
            lngPossibles(lngPossibleCounter) = lngIndex2
            
'            ' FOR DEBUGGING
'            Set pNewPoint = New Point
'            Set pNewPoint.SpatialReference = pSpRef
'            pNewPoint.PutCoords dblEndX, dblEndY
'            MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker2
'            ' --------------------------------------------
            
          End If
        End If
      Next lngIndex2
      
      ' NEXT, GO THROUGH ALL POINTS IN POSSIBLES LIST TO FIND NEW POINTS WITHIN DISTANCE
      Do Until lngPossibleCounter = -1
        lngPossibleCounter = -1
'        Set pAlreadyFoundColl = New Collection
        
        lngTempIndices = lngPossibles
        Erase lngPossibles
        For lngPossibleIndex = 0 To UBound(lngTempIndices)
          dblStartX = dblXYZ(0, lngTempIndices(lngPossibleIndex))
          dblStartY = dblXYZ(1, lngTempIndices(lngPossibleIndex))
          dblStartZ = dblXYZ(2, lngTempIndices(lngPossibleIndex))
          
          ' GO THROUGH ALL POINTS GREATER THAN INDEX 1 AGAIN
          For lngIndex2 = lngIndex1 + 1 To lngMaxIndex
          
            ' ONLY CONTINUE CHECKING IF NO CLUSTER ALREADY ASSIGNED
            If dblXYZ(3, lngIndex2) = 0 Then
              
              dblEndX = dblXYZ(0, lngIndex2)
              dblEndY = dblXYZ(1, lngIndex2)
              dblEndZ = dblXYZ(2, lngIndex2)
              
              dblDist = MyGeometricOperations.DistancePythagoreanNumbers_3D(dblStartX, dblStartY, _
                  dblStartZ, dblEndX, dblEndY, dblEndZ)
              
              If dblDist <= dblThresholdDist Then
                dblXYZ(3, lngIndex2) = dblRunningCluster
                lngPossibleCounter = lngPossibleCounter + 1
                ReDim Preserve lngPossibles(lngPossibleCounter)
                lngPossibles(lngPossibleCounter) = lngIndex2
                            
'                ' FOR DEBUGGING
'                Set pNewPoint = New Point
'                Set pNewPoint.SpatialReference = pSpRef
'                pNewPoint.PutCoords dblEndX, dblEndY
'                MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPoint, "Delete_Me", pMarker3
'                ' --------------------------------------------
                
              End If
            End If
          Next lngIndex2
          
        Next lngPossibleIndex
          
'        ' FOR DEBUGGING
'        QuickSort.LongAscending lngPossibles, 0, UBound(lngPossibles)
'        For lngPossibleCounter = 0 To UBound(lngPossibles)
'          Debug.Print CStr(lngPossibleCounter) & "] " & CStr(lngPossibles(lngPossibleCounter))
'        Next lngPossibleCounter
          
      Loop
      
    End If
  Next lngIndex1
      
  psbar.HideProgressBar
  pProg.position = 0
  
  Call CreateInMem_Of_Points(dblXYZ, pSpRef, strName, pMxDoc, dblThresholdDist)
  
  Debug.Print "Done..."
'  Debug.Print CStr(dblRunningCluster) & " clusters..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set psbar = Nothing
  Set pProg = Nothing
  Set pFClass = Nothing
  Erase dblXYZ
  Erase varAttNames
  Erase varAttributes
  Set pAttFields = Nothing
  Erase varOrigPoints
  Erase lngTempIndices
  Erase lngPossibles
  Set pAlreadyFoundColl = Nothing



End Sub

Public Function ReturnDistanceToClosestCave(pMPoint As IMultipoint, pCurrentFClass As IFeatureClass, _
    strName As String) As Double

  Dim pCurrentPoint As IPoint
  Dim pProxOp As IProximityOperator
  Dim pCurrentFCursor As IFeatureCursor
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String
  Dim pFeature As IFeature
  Dim dblDist As Double
  Dim dblMinDist As Double
  Dim booIsFirst As Boolean
  
  Set pQueryFilt = New QueryFilter
  
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pCurrentFClass, strPrefix, strSuffix)
  Select Case strName
    Case "All Points"
      pQueryFilt.WhereClause = ""
    Case "Day Points"
      pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Day'"
    Case "Night Points"
      pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Night'"
    Case "Difference Points"
      pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Diff'"
    Case Else
    
      If InStr(1, strName, "_Day_", vbTextCompare) > 0 Then
        pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Day'"
      ElseIf InStr(1, strName, "_Night_", vbTextCompare) > 0 Then
        pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Night'"
      ElseIf InStr(1, strName, "_Diff_", vbTextCompare) > 0 Then
        pQueryFilt.WhereClause = strPrefix & "Day_Night" & strSuffix & " = 'Diff'"
      Else
        MsgBox "Problem with Name! No option for '" & strName & "'"
      End If
  
  End Select
  
  Set pCurrentFCursor = pCurrentFClass.Search(pQueryFilt, False)
  Set pFeature = pCurrentFCursor.NextFeature
  booIsFirst = True
  
  Set pProxOp = pMPoint
  
  Do Until pFeature Is Nothing
    Set pCurrentPoint = pFeature.ShapeCopy
    
    If Not MyGeneralOperations.CompareSpatialReferences(pMPoint.SpatialReference, _
        pCurrentPoint.SpatialReference) Then
      pCurrentPoint.Project pMPoint.SpatialReference
    End If
    
    dblDist = pProxOp.ReturnDistance(pCurrentPoint)
    
    If booIsFirst Then
      dblMinDist = dblDist
      booIsFirst = False
    Else
      dblMinDist = MyGeometricOperations.MinDouble(dblMinDist, dblDist)
    End If
    
    Set pFeature = pCurrentFCursor.NextFeature
  Loop
  
  ReturnDistanceToClosestCave = dblMinDist
  
ClearMemory:
  Set pCurrentPoint = Nothing
  Set pProxOp = Nothing
  Set pCurrentFCursor = Nothing
  Set pQueryFilt = Nothing
  Set pFeature = Nothing




End Function

Public Sub CreateInMem_Of_Points(dblXYZ() As Double, pSpRef As ISpatialReference, _
    strName As String, pMxDoc As IMxDocument, dblCusterDist As Double, varStatsToReturn() As Variant, _
    booCreateFLayer As Boolean)
  
  Dim pCurrentFClass As IFeatureClass
  
  Dim pWSFact As IWorkspaceFactory
  Dim pFeatWS As IFeatureWorkspace
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pFeatWS = pWSFact.OpenFromFile( _
      "E:\arcGIS_stuff\consultation\Jut_Wynne\aaa_Phase_2b\Pisgah_Caves.gdb", 0)
  Set pCurrentFClass = pFeatWS.OpenFeatureClass("Pisgah_CavesAll_jsj_2016_02_13")
  
  Dim lngIndex As Long
  Dim pPoint As IPoint
  Dim pFClass As IFeatureClass
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pFieldArray As esriSystem.IVariantArray
  
  Dim lngDistanceIndex As Long
  Dim lngClusterIndex As Long
  Dim lngClusterIncludesIndex As Long
  Dim lngPointCountIndex As Long
  
  Dim pFBuffer As IFeatureBuffer
  Dim pFCursor As IFeatureCursor
  Dim pMPoint As IPointCollection
  Dim pGeom As IGeometry
  Dim varData() As Variant
  Dim pColl As Collection
  Dim lngCounter As Long
  Dim strCluster As String
  Dim lngCluster As Long
  
  If booCreateFLayer Then
    Set pFieldArray = New esriSystem.VarArray
    
  '  Set pField = New Field
  '  Set pFieldEdit = pField
  '  With pFieldEdit
  '    .Name = "X"
  '    .Type = esriFieldTypeDouble
  '  End With
  '  pFieldArray.Add pField
  '
  '  Set pField = New Field
  '  Set pFieldEdit = pField
  '  With pFieldEdit
  '    .Name = "Y"
  '    .Type = esriFieldTypeDouble
  '  End With
  '  pFieldArray.Add pField
  
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Dist_To_Cave"
      .Type = esriFieldTypeDouble
    End With
    pFieldArray.Add pField
  
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Cluster"
      .Type = esriFieldTypeInteger
    End With
    pFieldArray.Add pField
  
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "ClusterIncludesCave"
      .Type = esriFieldTypeString
      .length = 5
    End With
    pFieldArray.Add pField
  
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "PointCount"
      .Type = esriFieldTypeInteger
    End With
    pFieldArray.Add pField
    
    Set pFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pFieldArray, strName, _
        pSpRef, esriGeometryMultipoint, False, False)
        
    lngDistanceIndex = pFClass.FindField("Dist_To_Cave")
    lngClusterIndex = pFClass.FindField("Cluster")
    lngClusterIncludesIndex = pFClass.FindField("ClusterIncludesCave")
    lngPointCountIndex = pFClass.FindField("PointCount")
    
    Set pFCursor = pFClass.Insert(True)
    Set pFBuffer = pFClass.CreateFeatureBuffer
  End If
  
  lngCounter = -1
  Set pColl = New Collection
  
  Dim lngInCount As Long
  Dim lngTotalCount As Long
  
  ' MAKE CLUSTER MULTIPOINT OBJECTS, AND ADD TO ARRAY WITH CLUSTER NUMBER
  For lngIndex = 0 To UBound(dblXYZ, 2)
    Set pPoint = New Point
    Set pPoint.SpatialReference = pSpRef
    pPoint.PutCoords dblXYZ(0, lngIndex), dblXYZ(1, lngIndex)
    lngCluster = CLng(dblXYZ(3, lngIndex))
    strCluster = CStr(lngCluster)
    If MyGeneralOperations.CheckCollectionForKey(pColl, strCluster) Then
      Set pMPoint = pColl.Item(strCluster)
      pMPoint.AddPoint pPoint
      pColl.Remove strCluster
    Else
      Set pMPoint = New Multipoint
      Set pGeom = pMPoint
      Set pGeom.SpatialReference = pSpRef
      pMPoint.AddPoint pPoint
      lngCounter = lngCounter + 1
      ReDim Preserve varData(1, lngCounter)
      Set varData(0, lngCounter) = pMPoint
      varData(1, lngCounter) = lngCluster
    End If
    pColl.Add pMPoint, strCluster
  Next lngIndex
  
  Dim dblMinDist As Double
  Dim lngThreshCounter As Long
  
  lngThreshCounter = 0
  
  Dim lngNearCaveCount As Long
  Dim lngNotNearCaveCount As Long
  
  lngNearCaveCount = 0
  lngNotNearCaveCount = 0
  
  For lngIndex = 0 To UBound(varData, 2)
    Set pMPoint = varData(0, lngIndex)
    lngCluster = varData(1, lngIndex)
    dblMinDist = ReturnDistanceToClosestCave(pMPoint, pCurrentFClass, strName)
    If dblMinDist <= dblCusterDist Then
      lngThreshCounter = lngThreshCounter + 1
      If booCreateFLayer Then pFBuffer.Value(lngClusterIncludesIndex) = "True"
      lngInCount = lngInCount + pMPoint.PointCount
      lngNearCaveCount = lngNearCaveCount + 1
    Else
      If booCreateFLayer Then pFBuffer.Value(lngClusterIncludesIndex) = "False"
      lngNotNearCaveCount = lngNotNearCaveCount + 1
    End If
    lngTotalCount = lngTotalCount + pMPoint.PointCount
    If booCreateFLayer Then
      Set pFBuffer.Shape = pMPoint
      pFBuffer.Value(lngClusterIndex) = lngCluster
      pFBuffer.Value(lngDistanceIndex) = dblMinDist
      pFBuffer.Value(lngPointCountIndex) = pMPoint.PointCount
      pFCursor.InsertFeature pFBuffer
    End If
  Next lngIndex
  
  Dim pNewFLayer As IFeatureLayer
  If booCreateFLayer Then
    pFCursor.Flush
    
  '  For lngIndex = 0 To UBound(dblXYZ, 2)
  '    Set pPoint = New Point
  '    Set pPoint.SpatialReference = pSpRef
  '    pPoint.PutCoords dblXYZ(0, lngIndex), dblXYZ(1, lngIndex)
  '    Set pFBuffer.Shape = pPoint
  '    pFBuffer.Value(lngXIndex) = dblXYZ(0, lngIndex)
  '    pFBuffer.Value(lngYIndex) = dblXYZ(1, lngIndex)
  '    pFBuffer.Value(lngClusterIndex) = dblXYZ(3, lngIndex)
  '    pFCursor.InsertFeature pFBuffer
  '  Next lngIndex
  '  pFCursor.Flush
    
    Set pNewFLayer = New FeatureLayer
    pNewFLayer.Name = strName
    Set pNewFLayer.FeatureClass = pFClass
    pMxDoc.FocusMap.AddLayer pNewFLayer
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
  End If
  
  Debug.Print "Count for '" & strName & "'..."
  Debug.Print "  --> Analyzed " & CStr(UBound(dblXYZ, 2) + 1) & " points..."
  Debug.Print "  --> " & CStr(lngThreshCounter) & " points <= cluster distance (" & _
      CStr(dblCusterDist) & " m)"
  Debug.Print "  --> " & CStr(lngNearCaveCount) & " clusters <= cluster distance to existing cave (" & _
      CStr(dblCusterDist) & " m)"
  Debug.Print "  --> " & CStr(lngNotNearCaveCount) & " clusters > cluster distance to existing cave (" & _
      CStr(dblCusterDist) & " m)"
  Debug.Print "  --> Found " & CStr(UBound(varData, 2) + 1) & " clusters..."
  Debug.Print "  --> " & CStr(lngInCount) & " of " & CStr(lngTotalCount) & _
      " [" & Format(CDbl(lngInCount) / CDbl(lngTotalCount), "0%") & _
      "] cells in clusters close to known caves..."
  
  varStatsToReturn = Array(UBound(dblXYZ, 2) + 1, lngNearCaveCount, lngNotNearCaveCount, _
      UBound(varData, 2) + 1, lngInCount, lngTotalCount, CDbl(lngInCount) / CDbl(lngTotalCount))
  
ClearMemory:
  Set pPoint = Nothing
  Set pFClass = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pFieldArray = Nothing
  Set pFBuffer = Nothing
  Set pFCursor = Nothing
  Set pNewFLayer = Nothing


  

End Sub

Public Sub PerformClusterAnalysis()
  
  Debug.Print "--------------------------"
  
  Dim lngNumClusters As Long
  Dim lngNumIterations As Long
  Dim lngNumRuns As Long
  Dim dblPhi As Double
  
  lngNumClusters = 30
  lngNumIterations = 75
  lngNumRuns = 10
  dblPhi = 1.5
  
  Dim dblMinX As Double
  Dim dblMaxX As Double
  Dim dblMinY As Double
  Dim dblMaxY As Double
  Dim dblMinZ As Double
  Dim dblMaxZ As Double
  
  Dim lngStart As Long
  lngStart = GetTickCount
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim strLayerName As String
  Dim pFClass As IFeatureClass
  strLayerName = "Mammoth_Aravaipa_Springs_to_Analyze_Final"
  
  Debug.Print "Reading X, Y and Z Values..."
  Dim dblXYZ() As Double
  Dim varAttributes() As Variant
  Dim pAttFields As esriSystem.IVariantArray
  Dim pOrigPoints As esriSystem.IArray
  dblXYZ = ReturnArrayOfXYZ(strLayerName, pMxDoc, pFClass, dblMinX, dblMaxX, dblMinY, _
      dblMaxY, dblMinZ, dblMaxZ, varAttributes, pAttFields, pOrigPoints)
  
'  Call AttributesToText(varAttributes)
  
  Dim pFieldArray As esriSystem.IVariantArray
  Dim pValsArray As esriSystem.IVariantArray
  Dim pPoints As esriSystem.IArray
  Dim lngFinalClusterCount As Long
  
  Dim dblClusterCentroids() As Double
  ReDim dblClusterCentroids(2, lngNumClusters - 1)
'  Call InitializeCentroids(dblXYZ, dblClusterCentroids, dblMinX, dblMaxX, dblMinY, _
'      dblMaxY, dblMinZ, dblMaxZ)
'
''  Call AddFLayer(pMxDoc, 0, dblClusterCentroids, pFClass)
'  Call CumulativeFLayer(pMxDoc, 0, dblClusterCentroids, pFClass, pFieldArray, _
'      pValsArray, pPoints)
  
  Dim dblDistances() As Double
  ReDim dblDistances(lngNumClusters + 1, UBound(dblXYZ, 2))
  Dim dblMemberships() As Double
  ReDim dblMemberships(lngNumClusters + 1, UBound(dblXYZ, 2))
  
  Dim lngIndex As Long
  Dim lngRun As Long
  Dim pNewFClass As IFeatureClass
  Dim pNewFLayer As IFeatureLayer
  Dim pByClusterFLayer As IFeatureLayer
  Dim pGFLayer2 As IGeoFeatureLayer
  Dim pFLayerDef2 As IFeatureLayerDefinition2
  Dim pLegInfo As ILegendInfo
  Dim pLegGroup As ILegendGroup
  Dim pLayer As ILayer2
  Dim dblSSE As Double
  
  Dim dblBestDistances() As Double
  Dim dblBestSSE As Double
  Dim dblBestCentroids() As Double
  Dim pBestPoints As esriSystem.IArray
  Dim pBestValsArray As esriSystem.IVariantArray
  Dim lngBestClusterCount As Long
    
  For lngRun = 1 To lngNumRuns
    Debug.Print "  --> Run " & CStr(lngRun) & " of " & CStr(lngNumRuns)
    
    ReDim dblClusterCentroids(2, lngNumClusters - 1)
    Call InitializeCentroids(dblXYZ, dblClusterCentroids, dblMinX, dblMaxX, dblMinY, _
        dblMaxY, dblMinZ, dblMaxZ)
    
    Set pPoints = New esriSystem.Array
    Set pValsArray = New esriSystem.VarArray
        
    For lngIndex = 1 To lngNumIterations
      DoEvents
'      Debug.Print "  --> Pass " & CStr(lngIndex) & " of " & CStr(lngNumIterations)
      Call FillMemberships(dblXYZ, dblClusterCentroids, dblDistances, dblMemberships, _
          dblPhi, dblMinX, dblMaxX, dblMinY, dblMaxY, dblMinZ, dblMaxZ)
      
  '    Call AddFLayer(pMxDoc, lngIndex, dblClusterCentroids, pFClass)
      Call CumulativeFLayer(pMxDoc, lngIndex, dblClusterCentroids, pFClass, pFieldArray, _
          pValsArray, pPoints)
    Next lngIndex
    
    lngFinalClusterCount = ReturnNumberOfClusters(dblDistances, dblSSE)
    Debug.Print "      ... " & CStr(lngFinalClusterCount) & " clusters, " & _
          "SSE = " & Format(dblSSE, "#,##0") & "..."
    
    If lngRun = 1 Then
      dblBestSSE = dblSSE
      dblBestDistances = dblDistances
      dblBestCentroids = dblClusterCentroids
      Set pBestPoints = pPoints
      Set pBestValsArray = pValsArray
      lngBestClusterCount = lngFinalClusterCount
    Else
      If dblSSE < dblBestSSE Then
        dblBestSSE = dblSSE
        dblBestDistances = dblDistances
        dblBestCentroids = dblClusterCentroids
        Set pBestPoints = pPoints
        Set pBestValsArray = pValsArray
        lngBestClusterCount = lngFinalClusterCount
      End If
    End If
    
'    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pPoints, _
'        pValsArray, pFieldArray)
'    Set pNewFLayer = New FeatureLayer
'    Set pNewFLayer.FeatureClass = pNewFClass
'    pNewFLayer.Name = "Cluster Centroids, Run " & CStr(lngRun)
'    pNewFLayer.Visible = False
'    Set pGFLayer2 = pNewFLayer
'    Set pFLayerDef2 = pNewFLayer
'    Set pLayer = pNewFLayer
'    pFLayerDef2.DefinitionExpression = """Pass"" = " & CStr(lngNumIterations)
'    Set pLegInfo = pGFLayer2.Renderer
'    Set pLegGroup = pLegInfo.LegendGroup(0)
'    pLegGroup.Visible = False
'    pMxDoc.FocusMap.AddLayer pNewFLayer
'
'    Set pByClusterFLayer = PointByClusterLayer(dblXYZ, lngRun, pFClass, dblDistances, _
'        lngFinalClusterCount, dblSSE)
'    Set pByClusterFLayer = CreateAndApplyUVRenderer(pByClusterFLayer, "Cluster")
'    pByClusterFLayer.Visible = False
'    Set pGFLayer2 = pByClusterFLayer
'    Set pFLayerDef2 = pByClusterFLayer
'    Set pLayer = pByClusterFLayer
'    Set pLegInfo = pGFLayer2.Renderer
'    Set pLegGroup = pLegInfo.LegendGroup(0)
'    pLegGroup.Visible = False
'    pMxDoc.FocusMap.AddLayer pByClusterFLayer
    
    
'    Call MembershipsToText(dblDistances)
  Next lngRun
    
  Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pBestPoints, _
      pBestValsArray, pFieldArray)
  Set pNewFLayer = New FeatureLayer
  Set pNewFLayer.FeatureClass = pNewFClass
  pNewFLayer.Name = "Cluster Centroids, Run " & CStr(lngRun)
  pNewFLayer.Visible = False
  Set pGFLayer2 = pNewFLayer
  Set pFLayerDef2 = pNewFLayer
  Set pLayer = pNewFLayer
  pFLayerDef2.DefinitionExpression = """Pass"" = " & CStr(lngNumIterations)
  Set pLegInfo = pGFLayer2.Renderer
  Set pLegGroup = pLegInfo.LegendGroup(0)
  pLegGroup.Visible = False
  pMxDoc.FocusMap.AddLayer pNewFLayer
  
  Set pByClusterFLayer = PointByClusterLayer(dblXYZ, lngRun, pFClass, dblBestDistances, _
      lngBestClusterCount, dblBestSSE, varAttributes)
  Set pByClusterFLayer = CreateAndApplyUVRenderer(pByClusterFLayer, "Cluster")
  pByClusterFLayer.Visible = False
  Set pGFLayer2 = pByClusterFLayer
  Set pFLayerDef2 = pByClusterFLayer
  Set pLayer = pByClusterFLayer
  Set pLegInfo = pGFLayer2.Renderer
  Set pLegGroup = pLegInfo.LegendGroup(0)
  pLegGroup.Visible = False
  pMxDoc.FocusMap.AddLayer pByClusterFLayer
  
  Dim pFinalFLayer As IFeatureLayer
  Set pFinalFLayer = MakeFinalFeatureLayer(varAttributes, pOrigPoints, pAttFields)
  pMxDoc.FocusMap.AddLayer pFinalFLayer
  
  Call MembershipsToText(dblMemberships)
  ' Call AttributesToText(varAttributes)
  
  pMxDoc.ActiveView.Refresh
  pMxDoc.UpdateContents
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)
  

ClearMemory:
  Set pMxDoc = Nothing
  Set pFClass = Nothing
  Erase dblXYZ
  Erase varAttributes
  Set pAttFields = Nothing
  Set pOrigPoints = Nothing
  Set pFieldArray = Nothing
  Set pValsArray = Nothing
  Set pPoints = Nothing
  Erase dblClusterCentroids
  Erase dblDistances
  Erase dblMemberships
  Set pNewFClass = Nothing
  Set pNewFLayer = Nothing
  Set pByClusterFLayer = Nothing
  Set pGFLayer2 = Nothing
  Set pFLayerDef2 = Nothing
  Set pLegInfo = Nothing
  Set pLegGroup = Nothing
  Set pLayer = Nothing
  Erase dblBestDistances
  Erase dblBestCentroids
  Set pBestPoints = Nothing
  Set pBestValsArray = Nothing
  Set pFinalFLayer = Nothing


  
End Sub

Public Sub TestRandomlySelect()

  Dim pFLayer As IFeatureLayer
  Dim lngSelCount As Long
  Dim lngClusterCount As Long
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Springs_by_Cluster", pMxDoc.FocusMap)
  lngSelCount = 4
  lngClusterCount = 30
  
  Call SelectRandomlyByCluster(pFLayer, lngClusterCount, lngSelCount)
  
ClearMemory:
  Set pFLayer = Nothing
  Set pMxDoc = Nothing


End Sub
Public Sub SelectRandomlyByCluster(pFLayer As IFeatureLayer, lngClusterCount As Long, _
    lngMaxPerCluster As Long)

  Dim pFeatSel As IFeatureSelection
  Set pFeatSel = pFLayer
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  pFeatSel.Clear
  
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strPrefix As String
  Dim strSuffix As String
  Set pFClass = pFLayer.FeatureClass
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix)
  
  Dim pQueryFilt As IQueryFilter
  Dim lngCluster As Long
  Dim pEsriSet As esriSystem.ISet
  Dim varRandoms() As Variant
  Dim lngIndex As Long
  
  Set pQueryFilt = New QueryFilter
  
  For lngCluster = 1 To lngClusterCount
    pQueryFilt.WhereClause = strPrefix & "Cluster_ID" & strSuffix & " = " & CStr(lngCluster)
    
    If pFClass.FeatureCount(pQueryFilt) <= lngMaxPerCluster Then
      pFeatSel.SelectFeatures pQueryFilt, esriSelectionResultAdd, False
      
    Else
      Set pFCursor = pFClass.Search(pQueryFilt, False)
      Set pEsriSet = MyGeneralOperations.CursorToSet_Features(pFCursor)
      varRandoms = RandomlySelectFromSet(pEsriSet, lngMaxPerCluster)
      For lngIndex = 0 To UBound(varRandoms)
        Set pFeature = varRandoms(lngIndex)
        pFeatSel.Add pFeature
      Next lngIndex
    End If
  Next lngCluster
  
  pFeatSel.SelectionChanged
  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh
    
    
ClearMemory:
  Set pFeatSel = Nothing
  Set pMxDoc = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pQueryFilt = Nothing
  Set pEsriSet = Nothing
  Erase varRandoms



End Sub


Public Function RandomlySelectFromSet(pEsriSet As esriSystem.ISet, _
    lngSelectionCount As Long) As Variant()
  
  Dim pSelColl As Collection
  Set pSelColl = New Collection
  
  Dim lngIndex As Long
  Dim lngCounter As Long
  
  lngCounter = 0
  
  Dim pFeature As IFeature
  pEsriSet.Reset
  Set pFeature = pEsriSet.Next
  Do Until pFeature Is Nothing
    lngCounter = lngCounter + 1
    pSelColl.Add pFeature ', CStr(lngCounter) ' indexes 1-based
    
    Set pFeature = pEsriSet.Next
  Loop
    
  Dim varReturn() As Variant
  ReDim varReturn(lngSelectionCount - 1)
  
  Randomize
  Dim lngRandom As Long
  For lngIndex = 1 To lngSelectionCount
    lngRandom = Int(Rnd() * CDbl(pSelColl.Count)) + 1    ' RANDOMS NOW 1-BASED
    Set varReturn(lngIndex - 1) = pSelColl.Item(lngRandom)
    pSelColl.Remove lngRandom
  Next lngIndex
  
  RandomlySelectFromSet = varReturn
  
ClearMemory:
  Set pSelColl = Nothing
  Set pFeature = Nothing
  Erase varReturn




End Function



Public Function MakeFinalFeatureLayer(varAttributes As Variant, pOrigPoints As esriSystem.IArray, _
    pAttFields As esriSystem.IVariantArray) As IFeatureLayer

  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim pSubArray As esriSystem.IVariantArray
  Dim pValsArray As esriSystem.IVariantArray
  
  Set pValsArray = New esriSystem.VarArray
  
  For lngIndex1 = 0 To UBound(varAttributes, 2)
    Set pSubArray = New esriSystem.VarArray
    For lngIndex2 = 0 To UBound(varAttributes, 1)
      pSubArray.Add varAttributes(lngIndex2, lngIndex1)
    Next lngIndex2
    pValsArray.Add pSubArray
  Next lngIndex1
  
  Dim pNewFClass As IFeatureClass
  Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pOrigPoints, _
      pValsArray, pAttFields)
  
  Dim pFLayer As IFeatureLayer
  Dim pGFLayer2 As IGeoFeatureLayer
  Dim pFLayerDef2 As IFeatureLayerDefinition2
  Dim pLegInfo As ILegendInfo
  Dim pLegGroup As ILegendGroup
  Dim pLayer As ILayer
  
  Set pFLayer = New FeatureLayer
  Set pFLayer.FeatureClass = pNewFClass
  
  Set pFLayer = CreateAndApplyUVRenderer(pFLayer, "Cluster_ID")
  pFLayer.Visible = False
  Set pGFLayer2 = pFLayer
  Set pFLayerDef2 = pFLayer
  Set pLayer = pFLayer
  Set pLegInfo = pGFLayer2.Renderer
  Set pLegGroup = pLegInfo.LegendGroup(0)
  pLegGroup.Visible = False
  pFLayer.Name = "Springs_by_Cluster"
  
  Set MakeFinalFeatureLayer = pFLayer
  
  
ClearMemory:
  Set pSubArray = Nothing
  Set pValsArray = Nothing
  Set pNewFClass = Nothing
  Set pFLayer = Nothing
  Set pGFLayer2 = Nothing
  Set pFLayerDef2 = Nothing
  Set pLegInfo = Nothing
  Set pLegGroup = Nothing
  Set pLayer = Nothing



End Function

Public Sub FillMemberships(dblXYZ() As Double, dblClusterCentroids() As Double, _
    dblDistances() As Double, dblMemberships() As Double, dblPhi As Double, _
    dblMinX As Double, _
    dblMaxX As Double, dblMinY As Double, dblMaxY As Double, dblMinZ As Double, _
    dblMaxZ As Double)
    
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblDistance As Double
  Dim dblSortDistances() As Double
  Dim dblSortFuzzy() As Double
  Dim lngInsertIndex As Long
  Dim lngAssignCluster As Long
  Dim lngAssignFuzzyCluster As Long
  lngInsertIndex = UBound(dblClusterCentroids, 2)
  
  Dim pCentroidColl As New Collection
  Dim dblX() As Double
  Dim dblY() As Double
  Dim dblZ() As Double
  Dim varCoords As Variant
  Dim dblFuzzyDenominator As Double
  Dim dblFuzzyMembership As Double
  
  For lngIndex = 0 To lngInsertIndex
    varCoords = Array(dblX, dblY, dblZ)
    pCentroidColl.Add varCoords, CStr(lngIndex)
  Next lngIndex
  
  For lngIndex = 0 To UBound(dblXYZ, 2)
    ReDim dblSortDistances(1, UBound(dblClusterCentroids, 2))
    
    dblFuzzyDenominator = 0
    For lngIndex2 = 0 To UBound(dblClusterCentroids, 2)
      dblDistance = Sqr(((dblXYZ(0, lngIndex) - dblClusterCentroids(0, lngIndex2)) ^ 2) + _
                     ((dblXYZ(1, lngIndex) - dblClusterCentroids(1, lngIndex2)) ^ 2) + _
                     ((dblXYZ(2, lngIndex) - dblClusterCentroids(2, lngIndex2)) ^ 2))
      dblDistances(lngIndex2, lngIndex) = dblDistance
      dblSortDistances(0, lngIndex2) = dblDistance
      dblSortDistances(1, lngIndex2) = lngIndex2
      dblFuzzyDenominator = dblFuzzyDenominator + ((dblDistance ^ 2) ^ (-1 / (dblPhi - 1)))
    Next lngIndex2
    
    ReDim dblSortFuzzy(1, UBound(dblClusterCentroids, 2))
    For lngIndex2 = 0 To UBound(dblClusterCentroids, 2)
      dblDistance = dblDistances(lngIndex2, lngIndex)
      dblFuzzyMembership = ((dblDistance ^ 2) ^ (-1 / (dblPhi - 1))) / dblFuzzyDenominator
      dblMemberships(lngIndex2, lngIndex) = dblFuzzyMembership
      dblSortFuzzy(0, lngIndex2) = dblFuzzyMembership
      dblSortFuzzy(1, lngIndex2) = lngIndex2
    Next lngIndex2
    
    ' sorting on first column
    QuickSort.DoubleAscending_TwoDimensional _
        dblSortDistances, 0, UBound(dblSortDistances, 2), 0, 1
    QuickSort.DoubleAscending_TwoDimensional _
        dblSortFuzzy, 0, UBound(dblSortFuzzy, 2), 0, 1
    
    ' cluster id in second column
    lngAssignCluster = dblSortDistances(1, 0)
    dblDistances(lngInsertIndex + 1, lngIndex) = lngAssignCluster
    dblDistances(lngInsertIndex + 2, lngIndex) = dblSortDistances(1, 1)
    
    ' Fuzzy cluster id in second column
    lngAssignFuzzyCluster = dblSortFuzzy(1, UBound(dblSortFuzzy, 2))
    dblMemberships(lngInsertIndex + 1, lngIndex) = lngAssignFuzzyCluster
    dblMemberships(lngInsertIndex + 2, lngIndex) = dblSortDistances(1, UBound(dblSortFuzzy, 2) - 1)
    
    ' UPDATE CLUSTER COLLECTION
    varCoords = pCentroidColl.Item(CStr(lngAssignCluster))
    dblX = varCoords(0)
    dblY = varCoords(1)
    dblZ = varCoords(2)
    
    If MyGeneralOperations.IsDimmed(dblX) Then
      ReDim Preserve dblX(UBound(dblX) + 1)
      ReDim Preserve dblY(UBound(dblY) + 1)
      ReDim Preserve dblZ(UBound(dblZ) + 1)
    Else
      ReDim dblX(0)
      ReDim dblY(0)
      ReDim dblZ(0)
    End If
    
    dblX(UBound(dblX)) = dblXYZ(0, lngIndex)
    dblY(UBound(dblY)) = dblXYZ(1, lngIndex)
    dblZ(UBound(dblZ)) = dblXYZ(2, lngIndex)
    varCoords = Array(dblX, dblY, dblZ)
    pCentroidColl.Remove CStr(lngAssignCluster)
    pCentroidColl.Add varCoords, CStr(lngAssignCluster)
    
  Next lngIndex
  
  Dim dblNewCentroids() As Double
  ReDim dblNewCentroids(UBound(dblClusterCentroids, 1), UBound(dblClusterCentroids, 2))
  Dim dblMean As Double
  For lngIndex = 0 To lngInsertIndex  ' GOES TO HIGHEST CLUSTER INDEX
    varCoords = pCentroidColl.Item(CStr(lngIndex))
    dblX = varCoords(0)
    dblY = varCoords(1)
    dblZ = varCoords(2)
    
    If Not MyGeneralOperations.IsDimmed(dblX) Then
      dblClusterCentroids(0, lngIndex) = dblMinX + (Rnd() * (dblMaxX - dblMinX))
      dblClusterCentroids(1, lngIndex) = dblMinY + (Rnd() * (dblMaxY - dblMinY))
      dblClusterCentroids(2, lngIndex) = dblMinZ + (Rnd() * (dblMaxZ - dblMinZ))
    
    Else
      MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblX, False, , , , , dblMean
      dblClusterCentroids(0, lngIndex) = dblMean
      MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblY, False, , , , , dblMean
      dblClusterCentroids(1, lngIndex) = dblMean
      MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblZ, False, , , , , dblMean
      dblClusterCentroids(2, lngIndex) = dblMean
    End If
  Next lngIndex
  
ClearMemory:
  Erase dblSortDistances
  Erase dblNewCentroids


End Sub


Public Function ReturnNumberOfClusters(dblDistances() As Double, dblSSE As Double) As Long

  Dim lngIndex As Long
  Dim pColl As New Collection
  Dim lngClassIndex As Long
  lngClassIndex = UBound(dblDistances, 1) - 1
  Dim lngClass As Long
  dblSSE = 0
  Dim dblDistToCentroid As Double
  Dim lngDebug As Long
  
  For lngIndex = 0 To UBound(dblDistances, 2)
    lngClass = CLng(dblDistances(lngClassIndex, lngIndex))
    dblDistToCentroid = dblDistances(lngClass, lngIndex)
    dblSSE = dblSSE + (dblDistToCentroid ^ 2)
    
'    Debug.Print "Expecting Distance '" & Format(dblDistToCentroid, "0") & _
'          "' at Cluster '" & CStr(lngClass) & "'"
'    For lngDebug = 0 To UBound(dblDistances, 1)
'      Debug.Print IIf(lngClass = lngDebug, "--> ", "    ") & _
'          Format(dblDistances(lngDebug, lngIndex), "0")
'    Next lngDebug
    If Not MyGeneralOperations.CheckCollectionForKey(pColl, CStr(lngClass)) Then
      pColl.Add True, CStr(lngClass)
    End If
  Next lngIndex
  ReturnNumberOfClusters = pColl.Count
  
  Set pColl = Nothing
  
End Function

Public Function CreateAndApplyUVRenderer(pLayer As IFeatureLayer2, _
    strFieldName As String) As IFeatureLayer
     
     '** Paste into VBA
     '** Creates a UniqueValuesRenderer and applies it to first layer in the map.
     '** Layer must have "Name" field
 
     Dim pApp As Application
     Dim pDoc As IMxDocument
     Set pDoc = ThisDocument
     Dim pMap As IMap
     Set pMap = pDoc.FocusMap
 
     Dim pFLayer As IFeatureLayer
     Set pFLayer = pLayer
     Dim pLyr As IGeoFeatureLayer
     Set pLyr = pFLayer
     
     Dim pFeatCls As IFeatureClass
     Set pFeatCls = pFLayer.FeatureClass
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter 'empty supports: SELECT *
     Dim pFeatCursor As IFeatureCursor
     Set pFeatCursor = pFeatCls.Search(pQueryFilter, False)
 
     '** Make the color ramp we will use for the symbols in the renderer
     Dim rx As IRandomColorRamp
     Set rx = New RandomColorRamp
     rx.MinSaturation = 20
     rx.MaxSaturation = 80
     rx.MinValue = 45
     rx.MaxValue = 100
     rx.StartHue = 0
     rx.EndHue = 350
     rx.UseSeed = True
     rx.Seed = 43
     
     '** Make the renderer
     Dim pRender As IUniqueValueRenderer, n As Long
     Set pRender = New UniqueValueRenderer
'     Dim pOutline As ISimpleLineSymbol
     
     Dim symd As ISimpleMarkerSymbol
     Set symd = New SimpleMarkerSymbol
     symd.Style = esriSMSCircle
     symd.Outline = True
     symd.size = 13
          
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = strFieldName
     pRender.DefaultSymbol = symd
     pRender.UseDefaultSymbol = False
     
     Dim pFeat As IFeature
     n = pFeatCls.FeatureCount(pQueryFilter)
     '** Loop through the features
     Dim i As Integer
     i = 0
     Dim ValFound As Boolean
     Dim NoValFound As Boolean
     Dim uh As Integer
     Dim pFields As IFields
     Dim iField As Integer
     Set pFields = pFeatCursor.Fields
     iField = pFields.FindField(strFieldName)
     Do Until i = n
         Dim symx As ISimpleMarkerSymbol
         Set symx = New SimpleMarkerSymbol
         symx.Style = esriSMSCircle
         symx.Outline = True
         symx.size = 13
'         symx.Outline.Width = 0.4
         Set pFeat = pFeatCursor.NextFeature
         Dim X As String
         X = pFeat.Value(iField) '*new Cory*
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.Value(uh) = X Then
             NoValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             pRender.AddValue X, strFieldName, symx
             pRender.Label(X) = X
             pRender.Symbol(X) = symx
         End If
         i = i + 1
     Loop
     
     '** now that we know how many unique values there are
     '** we can size the color ramp and assign the colors.
     rx.size = pRender.ValueCount
     rx.CreateRamp (True)
     Dim RColors As IEnumColors, ny As Long
     Set RColors = rx.Colors
     RColors.Reset
     For ny = 0 To (pRender.ValueCount - 1)
         Dim xv As String
         xv = pRender.Value(ny)
         If xv <> "" Then
             Dim jsy As ISimpleMarkerSymbol
             Set jsy = pRender.Symbol(xv)
             jsy.Color = RColors.Next
             pRender.Symbol(xv) = jsy
         End If
     Next ny
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     pRender.ColorScheme = "Custom"
     pRender.fieldType(0) = True
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = strFieldName
     
     '** This makes the layer properties symbology tab show
     '** show the correct interface.
     Dim hx As IRendererPropertyPage
     Set hx = New UniqueValuePropertyPage
     pLyr.RendererPropertyPageClassID = hx.ClassID
 
     Set CreateAndApplyUVRenderer = pLyr
     '** Refresh the TOC
'     pDoc.ActiveView.ContentsChanged
'     pDoc.UpdateContents
     
     '** Draw the map
'     pDoc.ActiveView.Refresh
  
ClearMemory:
  Set pApp = Nothing
  Set pDoc = Nothing
  Set pMap = Nothing
  Set pFLayer = Nothing
  Set pLyr = Nothing
  Set pFeatCls = Nothing
  Set pQueryFilter = Nothing
  Set pFeatCursor = Nothing
  Set rx = Nothing
  Set symd = Nothing
  Set pFeat = Nothing
  Set pFields = Nothing
  Set symx = Nothing
  Set jsy = Nothing
  Set hx = Nothing

   
End Function

Public Sub MembershipsToText(dblMemberships() As Double)

  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim strReport As String
  Dim lngUBound2 As Long
  lngUBound2 = UBound(dblMemberships, 1)
  
  For lngIndex = 0 To UBound(dblMemberships, 2)
    For lngIndex2 = 0 To UBound(dblMemberships, 1)
      strReport = strReport & dblMemberships(lngIndex2, lngIndex) & _
            IIf(lngIndex2 = lngUBound2, vbCrLf, Chr(9))
    Next lngIndex2
  Next lngIndex
  
  Dim pDataObj As MSForms.DataObject
  Set pDataObj = New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard
  Set pDataObj = Nothing
  

End Sub
Public Sub AttributesToText(varAttributes As Variant)

  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim strReport As String
  Dim lngUBound2 As Long
  lngUBound2 = UBound(varAttributes, 1)
  
  Dim varVal As Variant
  Dim strVal As String
  
  For lngIndex = 0 To UBound(varAttributes, 2)
    For lngIndex2 = 0 To UBound(varAttributes, 1)
      varVal = varAttributes(lngIndex2, lngIndex)
      If IsNull(varVal) Then
        strVal = ""
      Else
        strVal = CStr(varVal)
      End If
      strReport = strReport & strVal & _
            IIf(lngIndex2 = lngUBound2, vbCrLf, Chr(9))
    Next lngIndex2
  Next lngIndex
  
  Dim pDataObj As MSForms.DataObject
  Set pDataObj = New MSForms.DataObject
  pDataObj.Clear
  pDataObj.SetText strReport
  pDataObj.PutInClipboard
  Set pDataObj = Nothing
  

End Sub

Public Function PointByClusterLayer(dblXYZ() As Double, lngRun As Long, _
    pFClass As IFeatureClass, dblMemberships() As Double, _
    lngNumClusters As Long, dblSSE As Double, varAttributes As Variant) As IFeatureLayer

  Dim lngIndex As Long
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  Dim dblClass As Double
  Dim lngClassIndex As Long
  lngClassIndex = UBound(dblMemberships, 1) - 1
  
  Dim pFieldArray As esriSystem.IVariantArray
  Dim pPoints As esriSystem.IArray
  Dim pValsArray As esriSystem.IVariantArray
  
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pSubArray As esriSystem.IVariantArray
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Dim pPoint As IPoint
  
  Set pFieldArray = New esriSystem.VarArray
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "X"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Y"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Z"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Cluster"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField

  Set pValsArray = New esriSystem.VarArray
  Set pPoints = New esriSystem.Array
  
  For lngIndex = 0 To UBound(dblXYZ, 2)
    dblX = dblXYZ(0, lngIndex)
    dblY = dblXYZ(1, lngIndex)
    dblZ = dblXYZ(2, lngIndex)
    dblClass = dblMemberships(lngClassIndex, lngIndex)
    
    Set pSubArray = New esriSystem.VarArray
    pSubArray.Add dblX
    pSubArray.Add dblY
    pSubArray.Add dblZ
    pSubArray.Add dblClass
    pValsArray.Add pSubArray
    
    Set pPoint = New Point
    Set pPoint.SpatialReference = pSpRef
    pPoint.PutCoords dblX, dblY
    pPoints.Add pPoint
    
    varAttributes(UBound(varAttributes, 1), lngIndex) = CLng(dblClass) + 1 ' MAKE 1-BASED
    
  Next lngIndex
  
  Dim pNewFLayer As IFeatureLayer
  Set pNewFLayer = New FeatureLayer
  Dim pNewFClass As IFeatureClass
  Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pPoints, _
      pValsArray, pFieldArray)
  Set pNewFLayer.FeatureClass = pNewFClass
  pNewFLayer.Name = "Run " & CStr(lngRun) & " [n = " & CStr(lngNumClusters) & _
      ", SSE = " & Format(dblSSE, "#,##0") & "]"
  Set PointByClusterLayer = pNewFLayer


ClearMemory:
  Set pFieldArray = Nothing
  Set pPoints = Nothing
  Set pValsArray = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSubArray = Nothing
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pPoint = Nothing
  Set pNewFLayer = Nothing
  Set pNewFClass = Nothing




End Function

Public Sub CumulativeFLayer(pMxDoc As IMxDocument, lngPass As Long, dblClusterCentroids() As Double, _
    pFClass As IFeatureClass, pFieldArray As esriSystem.IVariantArray, _
    pValsArray As esriSystem.IVariantArray, pPoints As esriSystem.IArray)

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pSubArray As esriSystem.IVariantArray
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Dim pPoint As IPoint
  
  If pFieldArray Is Nothing Then
    Set pFieldArray = New esriSystem.VarArray
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "X"
      .Type = esriFieldTypeDouble
    End With
    pFieldArray.Add pField
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Y"
      .Type = esriFieldTypeDouble
    End With
    pFieldArray.Add pField
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Z"
      .Type = esriFieldTypeDouble
    End With
    pFieldArray.Add pField
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Pass"
      .Type = esriFieldTypeInteger
    End With
    pFieldArray.Add pField
  
    Set pValsArray = New esriSystem.VarArray
    Set pPoints = New esriSystem.Array
  End If
  
  Dim lngIndex As Long
  For lngIndex = 0 To UBound(dblClusterCentroids, 2)
    Set pPoint = New Point
    Set pPoint.SpatialReference = pSpRef
    pPoint.PutCoords dblClusterCentroids(0, lngIndex), dblClusterCentroids(1, lngIndex)
    pPoints.Add pPoint
    Set pSubArray = New esriSystem.VarArray
    pSubArray.Add dblClusterCentroids(0, lngIndex)
    pSubArray.Add dblClusterCentroids(1, lngIndex)
    pSubArray.Add dblClusterCentroids(2, lngIndex)
    pSubArray.Add lngPass
    pValsArray.Add pSubArray
  Next lngIndex
  
  
  
ClearMemory:
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSubArray = Nothing
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pPoint = Nothing




End Sub

Public Sub AddFLayer(pMxDoc As IMxDocument, lngPass As Long, dblClusterCentroids() As Double, _
    pFClass As IFeatureClass)

  Dim pFieldArray As esriSystem.IVariantArray
  Set pFieldArray = New esriSystem.VarArray
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "X"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Y"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Z"
    .Type = esriFieldTypeDouble
  End With
  pFieldArray.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Pass"
    .Type = esriFieldTypeInteger
  End With
  pFieldArray.Add pField
  
  Dim pValsArray As esriSystem.IVariantArray
  Set pValsArray = New esriSystem.VarArray
  Dim pSubArray As esriSystem.IVariantArray
  Dim pPoints As esriSystem.IArray
  Set pPoints = New esriSystem.Array
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Dim pPoint As IPoint
  
  Dim lngIndex As Long
  For lngIndex = 0 To UBound(dblClusterCentroids, 2)
    Set pPoint = New Point
    Set pPoint.SpatialReference = pSpRef
    pPoint.PutCoords dblClusterCentroids(0, lngIndex), dblClusterCentroids(1, lngIndex)
    pPoints.Add pPoint
    Set pSubArray = New esriSystem.VarArray
    pSubArray.Add dblClusterCentroids(0, lngIndex)
    pSubArray.Add dblClusterCentroids(1, lngIndex)
    pSubArray.Add dblClusterCentroids(2, lngIndex)
    pSubArray.Add lngPass
    pValsArray.Add pSubArray
  Next lngIndex
  
  Dim pNewFClass As IFeatureClass
  Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass3(pPoints, _
      pValsArray, pFieldArray)
  Dim pNewFLayer As IFeatureLayer
  Set pNewFLayer = New FeatureLayer
  Set pNewFLayer.FeatureClass = pNewFClass
  pNewFLayer.Name = "Pass " & CStr(lngPass)
  pMxDoc.FocusMap.AddLayer pNewFLayer
  pMxDoc.UpdateContents
  
  
ClearMemory:
  Set pFieldArray = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pValsArray = Nothing
  Set pSubArray = Nothing
  Set pPoints = Nothing
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pPoint = Nothing
  Set pNewFClass = Nothing
  Set pNewFLayer = Nothing




End Sub


Public Function ReturnClusterMembership(dblDistToCentroid As Double, _
    dblDistToAllCentroids() As Double) As Double
    
  Dim lngIndex As Long
  Dim dblVal As Double
  
'  for l


End Function

Public Sub InitializeCentroids(dblXYZ() As Double, dblClusterCentroids() As Double, _
    dblMinX As Double, _
    dblMaxX As Double, dblMinY As Double, dblMaxY As Double, dblMinZ As Double, _
    dblMaxZ As Double)
  
  Dim lngClusterMax As Long
  lngClusterMax = UBound(dblClusterCentroids, 2)
  
  Dim pPointColl As New Collection
  Dim dblCoord() As Double
  ReDim dblCoord(2)
  Dim lngRandVal As Long
  Dim lngIndex As Long
  
  For lngIndex = 0 To UBound(dblXYZ, 2)
    dblCoord(0) = dblXYZ(0, lngIndex)
    dblCoord(1) = dblXYZ(1, lngIndex)
    dblCoord(2) = dblXYZ(2, lngIndex)
    pPointColl.Add dblCoord
  Next lngIndex
  
  Dim pCentroidColl As New Collection
  Dim dblX() As Double
  Dim dblY() As Double
  Dim dblZ() As Double
  Dim varCoords As Variant
  
  For lngIndex = 0 To lngClusterMax
    varCoords = Array(dblX, dblY, dblZ)
    pCentroidColl.Add varCoords, CStr(lngIndex)
  Next lngIndex
  
  Dim lngClusterCounter As Long
  lngClusterCounter = 0
  Randomize
  For lngIndex = 0 To UBound(dblXYZ, 2)
  
    lngRandVal = Int(Rnd() * pPointColl.Count)
    dblCoord = pPointColl.Item(lngRandVal + 1)
    pPointColl.Remove lngRandVal + 1
    
    lngClusterCounter = lngClusterCounter + 1
    If lngClusterCounter > lngClusterMax Then lngClusterCounter = 0
    varCoords = pCentroidColl.Item(CStr(lngClusterCounter))
    dblX = varCoords(0)
    dblY = varCoords(1)
    dblZ = varCoords(2)
    
    If MyGeneralOperations.IsDimmed(dblX) Then
      ReDim Preserve dblX(UBound(dblX) + 1)
      ReDim Preserve dblY(UBound(dblY) + 1)
      ReDim Preserve dblZ(UBound(dblZ) + 1)
    Else
      ReDim dblX(0)
      ReDim dblY(0)
      ReDim dblZ(0)
    End If
    
    dblX(UBound(dblX)) = dblCoord(0)
    dblY(UBound(dblY)) = dblCoord(1)
    dblZ(UBound(dblZ)) = dblCoord(2)
    varCoords = Array(dblX, dblY, dblZ)
    pCentroidColl.Remove CStr(lngClusterCounter)
    pCentroidColl.Add varCoords, CStr(lngClusterCounter)
    
  Next lngIndex
  
  Dim dblNewCentroids() As Double
  ReDim dblNewCentroids(UBound(dblClusterCentroids, 1), UBound(dblClusterCentroids, 2))
  Dim dblMean As Double
  For lngIndex = 0 To lngClusterMax
    varCoords = pCentroidColl.Item(CStr(lngIndex))
    dblX = varCoords(0)
    dblY = varCoords(1)
    dblZ = varCoords(2)
    MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblX, False, , , , , dblMean
    dblNewCentroids(0, lngIndex) = dblMean
    MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblY, False, , , , , dblMean
    dblNewCentroids(1, lngIndex) = dblMean
    MyGeneralOperations.BasicStatsFromArraySimpleFast2 dblZ, False, , , , , dblMean
    dblNewCentroids(2, lngIndex) = dblMean
  Next lngIndex
  
  dblClusterCentroids = dblNewCentroids
  
'  For lngIndex = 0 To UBound(dblClusterCentroids, 2)
'    dblClusterCentroids(0, lngIndex) = dblMinX + (Rnd() * (dblMaxX - dblMinX))
'    dblClusterCentroids(1, lngIndex) = dblMinY + (Rnd() * (dblMaxY - dblMinY))
'    dblClusterCentroids(2, lngIndex) = dblMinZ + (Rnd() * (dblMaxZ - dblMinZ))
'  Next lngIndex
  
End Sub

Public Function ReturnArrayOfXYZ(strLayerName As String, pMxDoc As IMxDocument, _
    pFClass As IFeatureClass, dblMinX As Double, dblMaxX As Double, _
    dblMinY As Double, dblMaxY As Double, dblMinZ As Double, _
    dblMaxZ As Double, varAttributes() As Variant, _
    pAttFields As esriSystem.IVariantArray, pOrigPoints As esriSystem.IArray) As Double()
  
  Dim pFLayer As IFeatureLayer
  Set pFLayer = MyGeneralOperations.ReturnLayerByName(strLayerName, pMxDoc.FocusMap)
  
  Dim lngXIndex As Long
  Dim lngYIndex As Long
  Dim lngZIndex As Long
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  
  Set pFClass = pFLayer.FeatureClass
  lngXIndex = pFClass.FindField("UTME")
  lngYIndex = pFClass.FindField("UTMN")
  lngZIndex = pFClass.FindField("ElevationM")
  
  Dim lngSiteIDIndex As Long
  Dim lngNameIndex As Long
  Dim lngLatitudeIndex As Long
  Dim lngLongitudeIndex As Long
  Dim lngDatumIndex As Long
  Dim lngElevIndex As Long
  Dim lngLandUnitIndex As Long
  Dim lngLandUnitDescIndex As Long
  Dim lngQuadIndex As Long
  Dim lngSurveyLevelIndex As Long
  Dim lngInvDevIndex As Long
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pClone As IClone
  
  lngSiteIDIndex = pFClass.FindField("SiteID")
  lngNameIndex = pFClass.FindField("SiteName")
  lngLatitudeIndex = pFClass.FindField("LatitudeDD")
  lngLongitudeIndex = pFClass.FindField("LongitudeDD")
  lngDatumIndex = pFClass.FindField("Datum")
  lngElevIndex = pFClass.FindField("ElevationM")
  lngLandUnitIndex = pFClass.FindField("LandUnit")
  lngLandUnitDescIndex = pFClass.FindField("LandUnitDetail")
  lngQuadIndex = pFClass.FindField("USGS_Quad")
  lngSurveyLevelIndex = pFClass.FindField("SurveyStatus")
  lngInvDevIndex = pFClass.FindField("InventoryLevel")
  
  Dim varIndexes As Variant
  varIndexes = Array(lngSiteIDIndex, lngNameIndex, lngLatitudeIndex, _
    lngLongitudeIndex, lngDatumIndex, lngElevIndex, lngLandUnitIndex, _
    lngLandUnitDescIndex, lngQuadIndex, lngSurveyLevelIndex, lngInvDevIndex)
      
  Set pAttFields = New esriSystem.VarArray
  Set pClone = pFClass.Fields.Field(lngSiteIDIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngNameIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngLatitudeIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngLongitudeIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngDatumIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngElevIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngLandUnitIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngLandUnitDescIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngQuadIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngSurveyLevelIndex)
  pAttFields.Add pClone.Clone
  Set pClone = pFClass.Fields.Field(lngInvDevIndex)
  pAttFields.Add pClone.Clone
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Sub_Basin"
    .Type = esriFieldTypeString
    .length = 60
  End With
  pAttFields.Add pField
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
    .Name = "Cluster_ID"
    .Type = esriFieldTypeInteger
  End With
  pAttFields.Add pField
    
  Dim lngCounter As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim dblReturn() As Double
  Dim lngAttIndex As Long
  
  Dim pSubBasinFLayer As IFeatureLayer
  Dim pSubBasinFClass As IFeatureClass
  Dim lngSubBasinIndex As Long
  Dim pSubBasinFCursor As IFeatureCursor
  Dim pSubBasinFeature As IFeature
  Dim pSpFilt As ISpatialFilter
  Dim pTransform As IGeoTransformation
  Dim pGeom As IGeometry2
  Dim pPoint As IPoint
  Dim pSpRef As ISpatialReference
  Dim pGeoDataset As IGeoDataset
  Dim strSubBasin As String
  
  Set pSpFilt = New SpatialFilter
  pSpFilt.SpatialRel = esriSpatialRelIntersects
  Set pSubBasinFLayer = MyGeneralOperations.ReturnLayerByName("Groundwater_Subbasin", pMxDoc.FocusMap)
  Set pSubBasinFClass = pSubBasinFLayer.FeatureClass
  lngSubBasinIndex = pSubBasinFClass.FindField("SUBBASIN_NAME_GWSI")
  Set pTransform = MyGeneralOperations.CreateNAD83_WGS84_GeoTransformationFlagstaff
  Set pGeoDataset = pSubBasinFClass
  Set pSpRef = pGeoDataset.SpatialReference
  Set pOrigPoints = New esriSystem.Array
  
  lngCounter = -1
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    lngCounter = lngCounter + 1
    
    ' GET GENERAL ATTRIBUTES
    ReDim Preserve varAttributes(UBound(varIndexes) + 2, lngCounter)
    For lngAttIndex = 0 To UBound(varIndexes)
      varAttributes(lngAttIndex, lngCounter) = pFeature.Value(CLng(varIndexes(lngAttIndex)))
    Next lngAttIndex
    
    ' GET SUBBASIN
    pOrigPoints.Add pFeature.ShapeCopy
    Set pPoint = pFeature.ShapeCopy
    Set pGeom = pPoint
    pGeom.ProjectEx pSpRef, esriTransformForward, pTransform, False, 0, 0
    Set pSpFilt.Geometry = pGeom
    Set pSubBasinFCursor = pSubBasinFClass.Search(pSpFilt, False)
    Set pSubBasinFeature = pSubBasinFCursor.NextFeature
    strSubBasin = pSubBasinFeature.Value(lngSubBasinIndex)
    varAttributes(UBound(varIndexes) + 1, lngCounter) = strSubBasin
    
    ReDim Preserve dblReturn(2, lngCounter)
    dblX = pFeature.Value(lngXIndex)
    dblY = pFeature.Value(lngYIndex)
    dblZ = pFeature.Value(lngZIndex)
    
    If lngCounter = 0 Then
      dblMinX = dblX
      dblMaxX = dblX
      dblMinY = dblY
      dblMaxY = dblY
      dblMinZ = dblZ
      dblMaxZ = dblZ
    Else
      dblMinX = MyGeometricOperations.MinDouble(dblX, dblMinX)
      dblMaxX = MyGeometricOperations.MaxDouble(dblX, dblMinX)
      dblMinY = MyGeometricOperations.MinDouble(dblY, dblMinY)
      dblMaxY = MyGeometricOperations.MaxDouble(dblY, dblMinY)
      dblMinZ = MyGeometricOperations.MinDouble(dblZ, dblMinZ)
      dblMaxZ = MyGeometricOperations.MaxDouble(dblZ, dblMinZ)
    End If
    
    dblReturn(0, lngCounter) = dblX
    dblReturn(1, lngCounter) = dblY
    dblReturn(2, lngCounter) = dblZ
     
    Set pFeature = pFCursor.NextFeature
  Loop

  ReturnArrayOfXYZ = dblReturn


ClearMemory:
  Set pFLayer = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pClone = Nothing
  varIndexes = Null
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase dblReturn
  Set pSubBasinFLayer = Nothing
  Set pSubBasinFClass = Nothing
  Set pSubBasinFCursor = Nothing
  Set pSubBasinFeature = Nothing
  Set pSpFilt = Nothing
  Set pTransform = Nothing
  Set pGeom = Nothing
  Set pPoint = Nothing
  Set pSpRef = Nothing
  Set pGeoDataset = Nothing



End Function

Public Function ReturnArrayOfXYZ_2(strLayerName As String, pMxDoc As IMxDocument, _
    pFClass As IFeatureClass, dblMinX As Double, dblMaxX As Double, _
    dblMinY As Double, dblMaxY As Double, dblMinZ As Double, _
    dblMaxZ As Double, varFieldNamesToReturn() As Variant, varAttributes() As Variant, _
    booUseSelected As Boolean, booIsProjected As Boolean, varOrigPoints() As Variant, _
    strQueryString As String) As Double()
  
  Dim pFLayer As IFeatureLayer
  Set pFLayer = MyGeneralOperations.ReturnLayerByName(strLayerName, pMxDoc.FocusMap)
  
  Dim lngIndex As Long
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  Dim pGeoDataset As IGeoDataset
  
  Set pFClass = pFLayer.FeatureClass
  Set pGeoDataset = pFClass
  booIsProjected = TypeOf pGeoDataset.SpatialReference Is IProjectedCoordinateSystem
  
  Dim lngIndexes() As Long
  Dim booGetAttributes As Boolean
  
  booGetAttributes = MyGeneralOperations.IsDimmed(varFieldNamesToReturn)
  If booGetAttributes Then
    ReDim lngIndexes(UBound(varFieldNamesToReturn))
    For lngIndex = 0 To UBound(varFieldNamesToReturn)
      lngIndexes(lngIndex) = pFClass.FindField(CStr(varFieldNamesToReturn(lngIndex)))
    Next lngIndex
  End If
        
  Dim lngCounter As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim dblReturn() As Double
  Dim lngAttIndex As Long
  Dim pPoint As IPoint
  Dim pGeomDef As IGeometryDef
  Dim pField As iField
  Set pField = pFClass.Fields.Field(pFClass.FindField(pFClass.ShapeFieldName))
  Set pGeomDef = pField.GeometryDef
  Dim pFeatSel As IFeatureSelection
  Dim pSelSet As ISelectionSet
    
  Dim booHasZ As Boolean
  booHasZ = pGeomDef.HasZ
  
  lngCounter = -1
  
  Dim pQueryFilt As IQueryFilter
  Set pQueryFilt = New QueryFilter
  
  If booUseSelected Then
    pQueryFilt.WhereClause = strQueryString
    Set pFCursor = pFClass.Search(pQueryFilt, False)
'    Set pFeatSel = pFLayer
'    Set pSelSet = pFeatSel.SelectionSet
'    If pSelSet.Count = 0 Then
'      Set pFCursor = pFClass.Search(Nothing, False)
'    Else
'      pSelSet.Search Nothing, False, pFCursor
'    End If
  Else
    Set pFCursor = pFClass.Search(Nothing, False)
  End If
  
  
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pPoint = pFeature.ShapeCopy
    
    If Not pPoint.IsEmpty Then
    
      
      lngCounter = lngCounter + 1
      
      ' GET GENERAL ATTRIBUTES
      ReDim Preserve varAttributes(UBound(lngIndexes), lngCounter)
      For lngAttIndex = 0 To UBound(lngIndexes)
        varAttributes(lngAttIndex, lngCounter) = pFeature.Value(CLng(lngIndexes(lngAttIndex)))
      Next lngAttIndex
      ReDim Preserve varOrigPoints(lngCounter)
      Set varOrigPoints(lngCounter) = pPoint
      
      ReDim Preserve dblReturn(3, lngCounter)
      dblX = pPoint.X
      dblY = pPoint.Y
      If booHasZ Then
        dblZ = pPoint.Z
      Else
        dblZ = 0
      End If
      
      If lngCounter = 0 Then
        dblMinX = dblX
        dblMaxX = dblX
        dblMinY = dblY
        dblMaxY = dblY
        dblMinZ = dblZ
        dblMaxZ = dblZ
      Else
        dblMinX = MyGeometricOperations.MinDouble(dblX, dblMinX)
        dblMaxX = MyGeometricOperations.MaxDouble(dblX, dblMinX)
        dblMinY = MyGeometricOperations.MinDouble(dblY, dblMinY)
        dblMaxY = MyGeometricOperations.MaxDouble(dblY, dblMinY)
        dblMinZ = MyGeometricOperations.MinDouble(dblZ, dblMinZ)
        dblMaxZ = MyGeometricOperations.MaxDouble(dblZ, dblMinZ)
      End If
      
      dblReturn(0, lngCounter) = dblX
      dblReturn(1, lngCounter) = dblY
      dblReturn(2, lngCounter) = dblZ
      dblReturn(3, lngCounter) = 0 ' - FOR CLUSTER ID
    End If
    
    Set pFeature = pFCursor.NextFeature
  Loop

  ReturnArrayOfXYZ_2 = dblReturn


ClearMemory:
  Set pFLayer = Nothing
  Set pGeoDataset = Nothing
  Erase lngIndexes
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase dblReturn
  Set pPoint = Nothing
  Set pGeomDef = Nothing
  Set pField = Nothing





End Function


Public Function ReturnArrayOfXYZ_3(pFClass As IFeatureClass, lngCount As Long, _
    dblMinX As Double, dblMaxX As Double, _
    dblMinY As Double, dblMaxY As Double, dblMinZ As Double, _
    dblMaxZ As Double, dblMinPHat As Double, dblMaxPHat As Double, _
    booIsProjected As Boolean, varOrigPoints() As Variant) As Double()
  
  Dim lngIndex As Long
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  Dim pGeoDataset As IGeoDataset
  Dim lngpHatIndex As Long
  Dim dblpHat As Double
  
  lngpHatIndex = pFClass.FindField("phat")
  Set pGeoDataset = pFClass
  booIsProjected = TypeOf pGeoDataset.SpatialReference Is IProjectedCoordinateSystem
  
  Dim lngIndexes() As Long
  Dim booGetAttributes As Boolean
          
  Dim lngCounter As Long
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim dblReturn() As Double
  Dim lngAttIndex As Long
  Dim pPoint As IPoint
  Dim pGeomDef As IGeometryDef
  Dim pField As iField
  Set pField = pFClass.Fields.Field(pFClass.FindField(pFClass.ShapeFieldName))
  Set pGeomDef = pField.GeometryDef
  Dim pFeatSel As IFeatureSelection
  Dim pSelSet As ISelectionSet
    
  Dim booHasZ As Boolean
  booHasZ = pGeomDef.HasZ
  
  lngCounter = -1
  Set pFCursor = pFClass.Search(Nothing, False)
    
  ' FIRST SORT OUT BY pHat
  Dim dblpHats() As Double
  Dim varPoints() As Variant
  ReDim dblpHats(pFClass.FeatureCount(Nothing) - 1)
  ReDim varPoints(UBound(dblpHats))
  
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pPoint = pFeature.ShapeCopy
    dblpHat = pFeature.Value(lngpHatIndex)
            
    lngCounter = lngCounter + 1
    Set varPoints(lngCounter) = pPoint
    dblpHats(lngCounter) = dblpHat
    Set pFeature = pFCursor.NextFeature
  Loop
  
  QuickSort.DoubleAscendingWithObjects dblpHats, varPoints, 0, UBound(dblpHats)
  
  dblMinPHat = dblpHats(UBound(dblpHats) - (lngCount - 1))
  dblMaxPHat = dblpHats(UBound(dblpHats))
  lngCounter = -1
'  For lngIndex = UBound(dblpHats) To (UBound(dblpHats) - lngCount) Step -1
'    Set pPoint = varPoints(lngIndex)
  For lngIndex = 1 To lngCount
    Set pPoint = varPoints(UBound(dblpHats) - (lngIndex - 1))
    lngCounter = lngCounter + 1
    
'    Debug.Print CStr(lngCounter + 1) & "] Long Index = " & Format(lngIndex, "#,##0") & ": " & _
'        "Counting from " & Format(UBound(dblpHats), "#,##0") & " to " & _
'        Format(UBound(dblpHats) - lngCount, "#,##0")
    
    ReDim Preserve varOrigPoints(lngCounter)
    Set varOrigPoints(lngCounter) = pPoint
      
    ReDim Preserve dblReturn(3, lngCounter)
    dblX = pPoint.X
    dblY = pPoint.Y
    If booHasZ Then
      dblZ = pPoint.Z
    Else
      dblZ = 0
    End If
    
    If lngCounter = 0 Then
      dblMinX = dblX
      dblMaxX = dblX
      dblMinY = dblY
      dblMaxY = dblY
      dblMinZ = dblZ
      dblMaxZ = dblZ
    Else
      dblMinX = MyGeometricOperations.MinDouble(dblX, dblMinX)
      dblMaxX = MyGeometricOperations.MaxDouble(dblX, dblMinX)
      dblMinY = MyGeometricOperations.MinDouble(dblY, dblMinY)
      dblMaxY = MyGeometricOperations.MaxDouble(dblY, dblMinY)
      dblMinZ = MyGeometricOperations.MinDouble(dblZ, dblMinZ)
      dblMaxZ = MyGeometricOperations.MaxDouble(dblZ, dblMinZ)
    End If
      
    dblReturn(0, lngCounter) = dblX
    dblReturn(1, lngCounter) = dblY
    dblReturn(2, lngCounter) = dblZ
    dblReturn(3, lngCounter) = 0 ' - FOR CLUSTER ID
    
  Next lngIndex
  
  
  ReturnArrayOfXYZ_3 = dblReturn


ClearMemory:
  Set pGeoDataset = Nothing
  Erase lngIndexes
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase dblReturn
  Set pPoint = Nothing
  Set pGeomDef = Nothing
  Set pField = Nothing
  Set pFeatSel = Nothing
  Set pSelSet = Nothing
  Erase dblpHats
  Erase varPoints



End Function




