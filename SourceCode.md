    Option Explicit

    Public Function Levenshtein(s1 As String, s2 As String) As Long
    'Levenshtein edit (text) distance
    'Distance between two strings in terms of number of insertions, deletions or substitutions

    Dim i As Long
    Dim j As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim d() As Long
    Dim min1 As Long
    Dim min2 As Long

    l1 = Len(s1)
    l2 = Len(s2)
    ReDim d(l1, l2)
    For i = 0 To l1
        d(i, 0) = i
    Next
    For j = 0 To l2
        d(0, j) = j
    Next
    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = d(i - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                d(i, j) = min1
            End If
        Next
    Next
    Levenshtein = d(l1, l2)
    End Function


    Public Function VLOOKUPfuzzy(s1 As String, rng As Range)
        Dim cllNumber As Long
        Dim shortestDistance As Long
        Dim shortestDistanceCell As Long
        Dim vlookupRes
        Set rng = rng.Columns(1)

        On Error GoTo Levenshtein
        vlookupRes = Application.WorksheetFunction.VLookup(s1, rng, 1, 0)
        VLOOKUPfuzzy = vlookupRes
        Exit Function

    Levenshtein:
        On Error GoTo 0
        shortestDistance = Levenshtein(s1, rng.Cells(1, 1))
        shortestDistanceCell = 1

        For cllNumber = 2 To rng.Cells.Count
            If Levenshtein(s1, rng.Cells(cllNumber, 1)) < shortestDistance Then
                shortestDistance = Levenshtein(s1, rng.Cells(cllNumber, 1))
                shortestDistanceCell = cllNumber
            End If
        Next cllNumber

        VLOOKUPfuzzy = rng.Cells(shortestDistanceCell, 1)

    End Function

    Public Function ВПРнечеткий(s1 As String, rng As Range)
        ВПРнечеткий = VLOOKUPfuzzy(s1, rng)
    End Function

    Public Function VLOOKUPfuzzy_help() As String
        VLOOKUPfuzzy_help = "Hit Wrap Text on Home tab for better view. " & Chr(10) & _
        "The VLOOKUPfuzzy function was designed to be like VLOOKUP but to return similar values if there is no exact match. " & Chr(10) & _
        "Parameters:" & Chr(10) _
            & "  s1 as String   :   Specifies the String to serach for (could be cell reference). " & Chr(10) _
            & "  rng as String  :   Specifies the range (column) where to search and to return values from. " & Chr(10) _
            & "Example : ""=VLOOKUPfuzzy(A1,B1:B10)""" & Chr(10) _
            & "WARNING! It will always return something. No NA errors! Сontact amchercashin@gmail.com for more information."
    End Function

    Public Function ВПРнечеткий_помощь() As String
        ВПРнечеткий_помощь = "Нажмите ""Переносить текст"" на вкладке Главная, что бы было лучше видно. " & Chr(10) & _
        "Формула ВПРнечеткий работает как ВПР, только возвращает похожие значения, если нет точного совпадения. " & Chr(10) & _
        "Параметры:" & Chr(10) _
            & "  s1 as String   :   Искомая строка (может быть ссылкой на ячейку). " & Chr(10) _
            & "  rng as String  :   Столбец - область поиска, из него возвращается найденное значение ." & Chr(10) _
            & "Пример : ""=ВПРнечеткий(A1;B1:B10)""" & Chr(10) _
            & "ВНИМАНИЕ! Формула всегда возвращает какое-то значение. Ошибок Н/Д не будет! Пишите на amchercashin@gmail.com если есть вопросы."
    End Function
