Attribute VB_Name = "GetArrayColumnHeaders"
Option Explicit
'@Folder "Code Common"
Public Function GetArrayColumnHeadersFromSourceAndTargetArray(ByRef sourceArray As Variant, ByRef targetArray As Variant) As Variant

    Dim headerArray() As Variant
    ReDim headerArray(1 To UBound(targetArray, 2), 1 To 3)
    Dim targetRowCount As Long
    For targetRowCount = LBound(targetArray, 2) To UBound(targetArray, 2)
        headerArray(targetRowCount, 1) = targetArray(1, targetRowCount)
        headerArray(targetRowCount, 2) = targetRowCount
    Next targetRowCount
    
    Dim importRowCount As Long
    Dim headerArrayRowCount As Long
    For importRowCount = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        For headerArrayRowCount = LBound(headerArray, 1) To UBound(headerArray, 1)
            If sourceArray(1, importRowCount) = headerArray(headerArrayRowCount, 1) Then
                headerArray(headerArrayRowCount, 3) = importRowCount
                Exit For
            End If
        Next headerArrayRowCount
    Next importRowCount
    
    GetArrayColumnHeadersFromSourceAndTargetArray = headerArray

End Function
