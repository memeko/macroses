
Sub FormatWebReference()
    ' Request resource information
    Dim resourceTitle As String: resourceTitle = InputBox("Введите название ресурса:")
    Dim resourceType As String: resourceType = InputBox("Введите тип ресурса (например, Сайт, Блог и т.д.):")
    Dim accessMode As String: accessMode = InputBox("Введите режим доступа (например, Свободный):")
    Dim url As String: url = InputBox("Введите URL ресурса:")
    Dim accessDate As String: accessDate = InputBox("Введите дату обращения (в формате ДД.ММ.ГГГГ):")

    Dim biblioStr As String
    biblioStr = resourceTitle & " : " & resourceType & ". " & "URL: " _
                & url & " (" & accessMode & "). " & "Дата обращения: " & accessDate & "."
       
    Dim rng As Range
    Set rng = ActiveDocument.Content
    rng.InsertAfter biblioStr
End Sub
