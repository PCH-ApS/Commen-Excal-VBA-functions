Attribute VB_Name = "CompareSourceWithTargetArray"
Option Explicit
'@Folder "Code Common"
Public Function CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArrayV2(ByRef sourceArray As Variant, ByRef targetArray As Variant, ByRef headerArray As Variant, ByVal dictionaryColumn As String) As Variant

' **** Get pos of dictionary column in target array
    Dim headerRowCount As Long
    Dim dictionaryColumnNo As Long
    For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
        If dictionaryColumn = headerArray(headerRowCount, 1) Then
            dictionaryColumnNo = headerArray(headerRowCount, 2)
        End If
    Next headerRowCount

' **** Compare arrays and add rows from importArray
    Dim newRowCount As Long
    newRowCount = 1
    Dim idArray() As Long
    ReDim Preserve idArray(1 To newRowCount)

' **** Create Dictionary of dates from targetArray
    Dim targetDates As Object
    Set targetDates = CreateObject("Scripting.Dictionary")
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 1) + 1 To UBound(targetArray, 1)
        If Not targetDates.Exists(targetArray(targetRowCount, dictionaryColumnNo)) Then
            targetDates.Add targetArray(targetRowCount, dictionaryColumnNo), 1
        End If
    Next targetRowCount

' **** Get pos of dictionary column in source array
    For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
        If dictionaryColumn = headerArray(headerRowCount, 1) Then
            dictionaryColumnNo = headerArray(headerRowCount, 3)
        End If
    Next headerRowCount

' **** Loop though sourceArray, skip header row
    Dim importRowCount As Long
    Dim sourceCompareString As String
    Dim importColumnCount As Long
    For importRowCount = LBound(sourceArray, 1) + 1 To UBound(sourceArray, 1)
    ' **** Test if the date is found in targetArray, if not found, then assume row should be added
        If targetDates.Exists(sourceArray(importRowCount, dictionaryColumnNo)) Then
            sourceCompareString = vbNullString
            For importColumnCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                For headerRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
                    If Not InStr(headerArray(headerRowCount, 1), "*") > 0 Then
                        If headerArray(headerRowCount, 1) = sourceArray(1, importColumnCount) Then
                            sourceCompareString = sourceCompareString & ";" & Trim$(CStr(sourceArray(importRowCount, importColumnCount)))
                            Exit For
                        End If
                    End If
                Next headerRowCount
            Next importColumnCount
            If SourceSearchStringFoundInTarget.SourceSearchStringFoundInTargetArrayV2(sourceCompareString, targetArray, headerArray, sourceArray(importRowCount, 1)) = False Then
                ReDim Preserve idArray(1 To newRowCount)
                idArray(newRowCount) = importRowCount
                newRowCount = newRowCount + 1
            End If
        Else
            ReDim Preserve idArray(1 To newRowCount)
            idArray(newRowCount) = importRowCount
            newRowCount = newRowCount + 1
        End If
    Next importRowCount
    
    CompareSourceWithTargetArrayAndCountRowsToAddFromSourceArrayV2 = idArray

End Function
