Sub CountTablesAndReferences()
    Dim tbl As Table
    Dim rng As Range
    Dim strText As String
    Dim tableCount As Integer
    Dim referenceCount As Integer
    
    tableCount = ActiveDocument.Tables.Count
    
    For Each rng In ActiveDocument.StoryRanges
        strText = rng.Text
        referenceCount = referenceCount + UBound(Split(strText, "Таблица "))
        referenceCount = referenceCount + UBound(Split(strText, "Табл. "))
    Next rng

    MsgBox "Количество таблиц: " & tableCount & "," & vbNewLine & "количество упоминаний 'Таблица' или 'Табл.': " & referenceCount, vbOKOnly, "Таблицы"
End Sub
