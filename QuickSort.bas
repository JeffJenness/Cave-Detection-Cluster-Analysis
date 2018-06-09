Attribute VB_Name = "QuickSort"
Option Explicit
Option Compare Binary

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
' see http://vbnet.mvps.org/index.html?code/sort/qsvariations.htm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODIFIED JAN. 4, 2006 BY JEFF JENNESS, TO SIMPLIFY IMPLEMENTATION IN ARCGIS
'-----------------------------------------------------------------------------
Public Sub AngleSort(pAngleInDegrees As esriSystem.IDoubleArray, Optional booSortClockwise As Boolean = True, _
    Optional dblCentralAngle As Double = -999)
    
  Dim dblAngles() As Double
  Dim lngIndex As Long
  ReDim dblAngles(pAngleInDegrees.Count - 1)
  
  For lngIndex = 0 To pAngleInDegrees.Count - 1
    dblAngles(lngIndex) = pAngleInDegrees.Element(lngIndex)
  Next lngIndex
  
  DoubleAscending dblAngles, 0, UBound(dblAngles)
  
  
'  Debug.Print
'  Debug.Print "Sorted ..."
'  Dim dblDebugGap As Double
'  Dim strDebugGap As String
'  For lngIndex = 0 To UBound(dblAngles)
'    If lngIndex = 0 Then
'      dblDebugGap = dblAngles(0) + 360 - dblAngles(UBound(dblAngles))
'    Else
'      dblDebugGap = dblAngles(lngIndex) - dblAngles(lngIndex - 1)
'    End If
'    strDebugGap = CStr(Format(dblDebugGap, "0"))
'    Debug.Print CStr(lngIndex + 1) & "]  " & CStr(Format(dblAngles(lngIndex), "0.00")) & "       Gap = " & strDebugGap
'  Next lngIndex
  
  Dim pTempArray As esriSystem.IDoubleArray
  Set pTempArray = New esriSystem.DoubleArray
  
  Dim lngSplitIndex As Long
  
  If dblCentralAngle <> -999 Then
    
    Dim dblSplitAngle As Double
    dblSplitAngle = dblCentralAngle - 180
    If dblSplitAngle < 0 Then
      dblSplitAngle = dblSplitAngle + 360
    End If
    
    lngSplitIndex = 0
    Do While lngSplitIndex <= UBound(dblAngles)
      If dblAngles(lngSplitIndex) > dblSplitAngle Then
        Exit Do
      End If
      lngSplitIndex = lngSplitIndex + 1
    Loop
    
    If lngSplitIndex <= UBound(dblAngles) Then
      For lngIndex = lngSplitIndex To UBound(dblAngles)
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
    If lngSplitIndex > 0 Then
      For lngIndex = 0 To lngSplitIndex - 1
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
  Else
  
    Dim dblLargestGap As Double
    dblLargestGap = dblAngles(0) + 360 - dblAngles(UBound(dblAngles))
    lngSplitIndex = 0
    Dim dblTempGap As Double
    
    For lngIndex = 0 To UBound(dblAngles) - 1
      dblTempGap = dblAngles(lngIndex + 1) - dblAngles(lngIndex)
      If dblTempGap > dblLargestGap Then
        dblLargestGap = dblTempGap
        lngSplitIndex = lngIndex + 1
      End If
    Next lngIndex
    
    
    If lngSplitIndex <= UBound(dblAngles) Then
      For lngIndex = lngSplitIndex To UBound(dblAngles)
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
    
    If lngSplitIndex > 0 Then
      For lngIndex = 0 To lngSplitIndex - 1
        pTempArray.Add dblAngles(lngIndex)
      Next lngIndex
    End If
  
  End If
  
'  Debug.Print
'  Debug.Print "Sorted and Wrapped..."
'  For lngIndex = 0 To UBound(dblAngles)
'    If lngIndex = 0 Then
'      dblDebugGap = pTempArray.Element(0) + 360 - pTempArray.Element(UBound(dblAngles))
'    Else
'      dblDebugGap = pTempArray.Element(lngIndex) - pTempArray.Element(lngIndex - 1)
'    End If
'    If dblDebugGap > 360 Then
'      dblDebugGap = dblDebugGap - 360
'    ElseIf dblDebugGap < 0 Then
'      dblDebugGap = dblDebugGap + 360
'    End If
'    strDebugGap = CStr(Format(dblDebugGap, "0"))
'    Debug.Print CStr(lngIndex + 1) & "]  " & CStr(Format(pTempArray.Element(lngIndex), "0.00")) & "       Gap = " & strDebugGap
'  Next lngIndex
  
  pAngleInDegrees.RemoveAll
  If booSortClockwise Then
    For lngIndex = 0 To pTempArray.Count - 1
      pAngleInDegrees.Add pTempArray.Element(lngIndex)
    Next lngIndex
  Else
    For lngIndex = pTempArray.Count - 1 To 0 Step -1
      pAngleInDegrees.Add pTempArray.Element(lngIndex)
    Next lngIndex
  End If

End Sub
Public Sub ByteAscending(narray() As Byte, inLow As Long, inHi As Long)

   Dim pivot As Byte
   Dim tmpSwap As Byte
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then ByteAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then ByteAscending narray(), tmpLow, inHi

End Sub


Public Sub ByteDescending(narray() As Byte, inLow As Long, inHi As Long)

   Dim pivot As Byte
   Dim tmpSwap As Byte
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then ByteDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then ByteDescending narray(), tmpLow, inHi

End Sub
Public Sub LongAscending(narray() As Long, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then LongAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then LongAscending narray(), tmpLow, inHi

End Sub


Public Sub LongDescending(narray() As Long, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then LongDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then LongDescending narray(), tmpLow, inHi

End Sub




Public Sub SingleAscending(narray() As Single, inLow As Long, inHi As Long)

   Dim pivot As Single
   Dim tmpSwap As Single
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then SingleAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then SingleAscending narray(), tmpLow, inHi

End Sub


Public Sub SingleDescending(narray() As Single, inLow As Long, inHi As Long)

   Dim pivot As Single
   Dim tmpSwap As Single
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then SingleDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then SingleDescending narray(), tmpLow, inHi

End Sub
Public Sub DoubleAscendingWithObjects(narray() As Double, varObjArray() As Variant, inLow As Long, inHi As Long)

  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizeSwap As Variant
  Dim tmpLow As Long
  Dim tmpHi  As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray((inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
       tmpSwap = narray(tmpLow)
       narray(tmpLow) = narray(tmpHi)
       narray(tmpHi) = tmpSwap
       
       tmpSizeSwap = varObjArray(tmpLow)
       varObjArray(tmpLow) = varObjArray(tmpHi)
       varObjArray(tmpHi) = tmpSizeSwap
       
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DoubleAscendingWithObjects narray(), varObjArray(), inLow, tmpHi
  If (tmpLow < inHi) Then DoubleAscendingWithObjects narray(), varObjArray(), tmpLow, inHi

End Sub

Public Sub StringsAscendingWithObjects(sarray() As String, varObjArray() As Variant, inLow As Long, inHi As Long)

   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpSizeSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
       
         tmpSizeSwap = varObjArray(tmpLow)
         varObjArray(tmpLow) = varObjArray(tmpHi)
         varObjArray(tmpHi) = tmpSizeSwap
        
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then StringsAscendingWithObjects sarray(), varObjArray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsAscendingWithObjects sarray(), varObjArray(), tmpLow, inHi

End Sub


Public Sub DoubleAscendingWithSizes(narray() As Double, nSizeArray() As Double, inLow As Long, inHi As Long)

  Dim pivot As Double
  Dim tmpSwap As Double
  Dim tmpSizeSwap As Double
  Dim tmpLow As Long
  Dim tmpHi  As Long
   
  tmpLow = inLow
  tmpHi = inHi
   
  pivot = narray((inLow + inHi) / 2)
  While (tmpLow <= tmpHi)
       
    While (narray(tmpLow) < pivot And tmpLow < inHi)
       tmpLow = tmpLow + 1
    Wend
  
    While (pivot < narray(tmpHi) And tmpHi > inLow)
       tmpHi = tmpHi - 1
    Wend
    If (tmpLow <= tmpHi) Then
       tmpSwap = narray(tmpLow)
       narray(tmpLow) = narray(tmpHi)
       narray(tmpHi) = tmpSwap
       
       tmpSizeSwap = nSizeArray(tmpLow)
       nSizeArray(tmpLow) = nSizeArray(tmpHi)
       nSizeArray(tmpHi) = tmpSizeSwap
       
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
     
  Wend
    
  If (inLow < tmpHi) Then DoubleAscendingWithSizes narray(), nSizeArray(), inLow, tmpHi
  If (tmpLow < inHi) Then DoubleAscendingWithSizes narray(), nSizeArray(), tmpLow, inHi

End Sub
Public Sub DoubleAscending(narray() As Double, inLow As Long, inHi As Long)

   Dim pivot As Double
   Dim tmpSwap As Double
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)

   While (tmpLow <= tmpHi)
       
      While (narray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DoubleAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DoubleAscending narray(), tmpLow, inHi

End Sub


Public Sub DoubleDescending(narray() As Double, inLow As Long, inHi As Long)

   Dim pivot As Double
   Dim tmpSwap As Double
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = narray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
        
      While (narray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > narray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DoubleDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DoubleDescending narray(), tmpLow, inHi

End Sub

Public Sub StringsAscending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then StringsAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsAscending sarray(), tmpLow, inHi

End Sub


Public Sub StringsDescending(sarray() As String, inLow As Long, inHi As Long)
  
   Dim pivot As String
   Dim tmpSwap As String
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
      
      While (sarray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
    
      While (pivot > sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
  
   Wend
  
   If (inLow < tmpHi) Then StringsDescending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then StringsDescending sarray(), tmpLow, inHi

End Sub


Public Sub VariantAscending(sarray() As Variant, inLow As Long, inHi As Long)
  
   Dim pivot As Variant
   Dim tmpSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
  
   While (tmpLow <= tmpHi)
   
      While (sarray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then VariantAscending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then VariantAscending sarray(), tmpLow, inHi

End Sub


Public Sub VariantDescending(sarray() As Variant, inLow As Long, inHi As Long)
  
   Dim pivot As Variant
   Dim tmpSwap As Variant
   Dim tmpLow As Long
   Dim tmpHi As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = sarray((inLow + inHi) / 2)
   
   While (tmpLow <= tmpHi)
      
      While (sarray(tmpLow) > pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
    
      While (pivot > sarray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = sarray(tmpLow)
         sarray(tmpLow) = sarray(tmpHi)
         sarray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
  
   Wend
  
   If (inLow < tmpHi) Then VariantDescending sarray(), inLow, tmpHi
   If (tmpLow < inHi) Then VariantDescending sarray(), tmpLow, inHi

End Sub

Public Sub DatesDescending(narray() As Date, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = DateToJulian(narray((inLow + inHi) / 2))
   
   While (tmpLow <= tmpHi)
        
      While DateToJulian(narray(tmpLow)) > pivot And (tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot > DateToJulian(narray(tmpHi))) And (tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DatesDescending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DatesDescending narray(), tmpLow, inHi

End Sub


Public Sub DatesAscending(narray() As Date, inLow As Long, inHi As Long)

   Dim pivot As Long
   Dim tmpSwap As Long
   Dim tmpLow As Long
   Dim tmpHi  As Long
   
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = DateToJulian(narray((inLow + inHi) / 2))

   While (tmpLow <= tmpHi)
       
      While (DateToJulian(narray(tmpLow)) < pivot) And (tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
   
      While (pivot < DateToJulian(narray(tmpHi))) And (tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
      
         tmpSwap = narray(tmpLow)
         narray(tmpLow) = narray(tmpHi)
         narray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
         
      End If
      
   Wend
    
   If (inLow < tmpHi) Then DatesAscending narray(), inLow, tmpHi
   If (tmpLow < inHi) Then DatesAscending narray(), tmpLow, inHi

End Sub

Private Function DateToJulian(MyDate As Date) As Long

  'Return a numeric value representing
  'the passed date
   DateToJulian = DateValue(MyDate)

End Function

Public Sub SortStringArray(pStringArray As esriSystem.IStringArray)

  Dim strArray() As String
  ReDim strArray(pStringArray.Count - 1)
  Dim lngIndex As Long
  For lngIndex = 0 To pStringArray.Count - 1
    strArray(lngIndex) = pStringArray.Element(lngIndex)
  Next lngIndex
  QuickSort.StringsAscending strArray, LBound(strArray), UBound(strArray)
  pStringArray.RemoveAll
  For lngIndex = 0 To UBound(strArray)
    pStringArray.Add strArray(lngIndex)
  Next lngIndex

End Sub
Public Sub SortDoubleArray(pDoubleArray As esriSystem.IDoubleArray)

  Dim dblArray() As Double
  ReDim dblArray(pDoubleArray.Count - 1)
  Dim lngIndex As Long
  For lngIndex = 0 To pDoubleArray.Count - 1
    dblArray(lngIndex) = pDoubleArray.Element(lngIndex)
  Next lngIndex
  QuickSort.DoubleAscending dblArray, LBound(dblArray), UBound(dblArray)
  pDoubleArray.RemoveAll
  For lngIndex = 0 To UBound(dblArray)
    pDoubleArray.Add dblArray(lngIndex)
  Next lngIndex

End Sub


