Attribute VB_Name = "ConvertTextToDateInArrayV2"
Option Explicit
'@Folder "Code Common"
Public Function ConvertTextToDateInArrayV2(ByRef arr As Variant, ByVal dateColumnName As String, ByVal dateFormat As String, ByVal oldDateSeperator As String, ByVal newDateSeperator As String) As Variant
   
    Dim columnCount As Long
    Dim columnName As String
    Dim rowCount As Long
    Dim cellValue As Variant
    Dim tmpCellValue As Variant
    
    Dim splitDay As Integer
    Dim splitMonth As Integer
    Dim splitYear As Integer
    
    Dim parts() As String
    Dim datePart As String
              
    For columnCount = LBound(arr, 2) To UBound(arr, 2)
        columnName = arr(1, columnCount)
        If columnName = dateColumnName Then
            For rowCount = LBound(arr, 1) + 1 To UBound(arr, 1)
                cellValue = arr(rowCount, columnCount)
                parts = split(cellValue, " ")
                datePart = parts(0)
                parts = split(datePart, "/")
                splitDay = CInt(parts(1))
                splitMonth = CInt(parts(0))
                splitYear = CInt(parts(2))
                tmpCellValue = DateSerial(splitYear, splitMonth, splitDay)
                If Not IsDate(cellValue) Then
                    tmpCellValue = Replace(cellValue, oldDateSeperator, newDateSeperator)
                    If IsDate(tmpCellValue) Then
                        arr(rowCount, columnCount) = tmpCellValue
                        arr(rowCount, columnCount) = CDate(arr(rowCount, columnCount))
                    End If
                Else
                    If IsDate(tmpCellValue) Then
                        tmpCellValue = Replace(tmpCellValue, oldDateSeperator, newDateSeperator)
                        arr(rowCount, columnCount) = CDate(tmpCellValue)
                    End If
                End If
            Next rowCount
        End If
    Next columnCount
    
    ConvertTextToDateInArrayV2 = arr
    
End Function
