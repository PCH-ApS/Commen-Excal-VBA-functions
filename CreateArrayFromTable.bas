Attribute VB_Name = "CreateArrayFromTable"
Option Explicit
'@Folder "Code Common"
Public Function CreateArrayFromTable(ByVal newTbl As ListObject) As Variant

    Dim numRows As Long
    Dim numCols As Long
    numRows = newTbl.Range.rows.Count
    numCols = newTbl.Range.columns.Count

    Dim myArray() As Variant
    ReDim myArray(1 To numRows, 1 To numCols)

    Dim rows As Long
    Dim columns As Long
    For rows = 1 To numRows
        For columns = 1 To numCols
            myArray(rows, columns) = newTbl.Range.Cells.Item(rows, columns).Value
        Next columns
    Next rows
    
    CreateArrayFromTable = myArray
    
End Function
