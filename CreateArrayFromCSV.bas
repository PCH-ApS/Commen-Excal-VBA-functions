Attribute VB_Name = "CreateArrayFromCSV"
Option Explicit
'@Folder "Code Common"
Public Function CreateArrayFromCSV(ByVal reportFile As String, ByRef delimitter As String) As Variant

    Open reportFile For Input As #1
    Dim numCols As Long
    Dim numRows As Long
    Dim rowFromFile As String
    Dim columnFromRow() As String
        
    Do Until EOF(1)
        Line Input #1, rowFromFile
        columnFromRow = split(rowFromFile, delimitter)
        numRows = numRows + 1
        'Array starts from 0 thus + 1. Check all rows to find max number of columns
        If UBound(columnFromRow) + 1 > numCols Then numCols = numCols + 1
    Loop
    
    Dim myArray() As Variant
    ReDim myArray(1 To numRows, 1 To numCols)
    Close #1
    
    Open reportFile For Input As #1
    numRows = 0
    Dim numColumns As Long
    Do Until EOF(1)
        Line Input #1, rowFromFile
        columnFromRow = split(rowFromFile, delimitter)
        numRows = numRows + 1
        For numColumns = 0 To UBound(columnFromRow)
            myArray(numRows, numColumns + 1) = columnFromRow(numColumns)
        Next numColumns
    Loop
    Close #1
    
    CreateArrayFromCSV = myArray
    
End Function
