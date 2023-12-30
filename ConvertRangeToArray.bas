Public Function ConvertRangeToArray(ByVal range As range) As Variant

    Dim resultArray() As String
    Dim resultRows As Long
    Dim resultColumns As Long
    
    resultRows = range.rows.Count
    resultColumns = range.columns.Count
    
    ReDim resultArray(resultRows, resultColumns)
    resultArray = range
    
    ConvertRangeToArray = resultArray
    
End Function