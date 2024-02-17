Attribute VB_Name = "ConvertTextToNumberInArray"
Option Explicit
'@Folder "Code Common"
Public Function ConvertTextToNumberInArray(ByRef arr As Variant, Optional ByRef decimalSeperator As String) As Variant
   
    Dim columnCount As Long
    Dim rowCount As Long
    Dim cellValue As Variant
    Dim tmpCellValue As Variant
    For columnCount = LBound(arr, 2) To UBound(arr, 2)
        For rowCount = LBound(arr, 1) To UBound(arr, 1)
            cellValue = arr(rowCount, columnCount)
            tmpCellValue = 0
            If IsNumeric(cellValue) Then
                If Not decimalSeperator = vbNullString Then
                    If InStr(cellValue, decimalSeperator) > 0 And (Len(cellValue) - InStr(cellValue, decimalSeperator) <= 2) Then
                        tmpCellValue = Replace(cellValue, decimalSeperator, ",")
                    End If
                    If InStr(cellValue, decimalSeperator) > 0 And (Len(cellValue) - InStr(cellValue, decimalSeperator) <= 1) Then
                        cellValue = cellValue & "0"
                        tmpCellValue = Replace(cellValue, decimalSeperator, ",")
                    End If
                    If InStr(cellValue, decimalSeperator) = 0 Then
                        cellValue = cellValue & ",00"
                        tmpCellValue = Replace(cellValue, decimalSeperator, ",")
                    End If
                    tmpCellValue = CDec(tmpCellValue)
                Else
                    If InStr(cellValue, ",") > 0 And (Len(cellValue) - InStr(cellValue, ",") <= 1) Then
                        cellValue = cellValue & "0"
                    End If
                    If InStr(cellValue, ",") = 0 Then
                        cellValue = cellValue & ",00"
                    End If
                    tmpCellValue = CDec(cellValue)
                End If
                arr(rowCount, columnCount) = tmpCellValue
            End If
        Next rowCount
    Next columnCount
    
    ConvertTextToNumberInArray = arr

End Function
