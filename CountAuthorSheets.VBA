Sub CountAuthorSheets()
    Dim charCount As Long
    Dim authorSheets As Double

    ' Подсчитать количество символов с пробелами
    charCount = ActiveDocument.Characters.Count

    ' Расчет авторских листов (по 40000 символов с пробелами на лист)
    authorSheets = charCount / 40000

    ' Вывод результатов в сообщении
    MsgBox "В документе " & Round(authorSheets, 2) & " авторских лист(а/ов).", vbOKOnly, "Подсчет авторских листов"
    
End Sub
