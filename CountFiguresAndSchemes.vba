Sub CountFiguresAndSchemes()
    Dim i As Integer
    Dim figureCount As Integer
    Dim schemeCount As Integer
    figureCount = 0
    schemeCount = 0
    picCount = 0

    For i = 1 To ActiveDocument.InlineShapes.Count
        If ActiveDocument.InlineShapes(i).Type = wdInlineShapePicture Then
            figureCount = figureCount + 1
            If InStr(1, ActiveDocument.InlineShapes(i).AlternativeText, "Схема", 1) > 0 _
                Or InStr(1, ActiveDocument.InlineShapes(i).AlternativeText, "схема", 1) > 0 Then
                schemeCount = schemeCount + 1
            End If
          If InStr(1, ActiveDocument.InlineShapes(i).AlternativeText, "Рисунок", 1) > 0 _
                Or InStr(1, ActiveDocument.InlineShapes(i).AlternativeText, "Рис.", 1) > 0 Then
                picCount = picCount + 1
            End If
        End If
    Next i

    ' Отображаем результат в строке состояния
     MsgBox "Количество картинок в документе: " & figureCount & ". " & vbNewLine & "Из них подписано как рисунки:" & picCount & ". " & vbNewLine & "Количество схем: " & schemeCount, vbOKOnly, "Иллюстрации"

End Sub
