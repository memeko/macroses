Sub FormatSelectedText()
    Dim rng As Range
    Set rng = Selection.Range 

    With rng.Font
        .Name = "Courier New"
        .Size = 12
        .Color = wdColorBlack
    End With
    
    Dim findRange As Range
    Set findRange = rng.Duplicate 
    
    With findRange.Find
        .ClearFormatting
        .Text = "#*^13" 
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True 
        .Execute
        
        Do While .Found
            Dim highlightRange As Range
            Set highlightRange = findRange.Duplicate
            highlightRange.MoveStart wdCharacter, 1 
            highlightRange.End = findRange.End 
            
            highlightRange.HighlightColorIndex = wdGray25
            
            findRange.Collapse wdCollapseEnd
            .Execute
        Loop
    End With
    
    MsgBox "Форматирование завершено!", vbInformation
End Sub
