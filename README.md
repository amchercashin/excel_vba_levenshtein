# VLOOKUPfuzzy
It's an Excel add-in with User Defined Function: =VLOOKUPfuzzy

This function is like regular VLOOKUP, but if it finds no exact match, it uses Levenshtein distance metric to find the closest (most similar) string.

VLOOKUPfuzzy could be userfull for normalising lists.

Use =VLOOKUPfuzzy_help() for help on parameters.

----------------------------------------------
# ВПРнечеткий
Это надстрока Excel с пользовательской функцией: =ВПРнечеткий

Это функция работает как стандартная ВПР, но, если она не находит точное совпадение, то возващает наиболее "похожее" значение используя "Расстояние Ливенштейна" как метрику "похожести".

ВПРнечеткий может быть полезен при нормализации данных по справочникам.

Введите =ВПРнечеткий() для помощи по параметрам.


## Пример работы
ВНИМАНИЕ! Формула всегда выдает "какой-то" ответ, даже если с человеческой точки зрения он совсем не к месту.

Например, если:

Если искать текст "кошак" в столбце с такими значениями: "кит", "тигр", "кошка", "крошка", то ВПРнечеткий выдаст "кошка".

Если искать аналогичный текст в столбце с: "кит", "тигр", "крошка", то ВПРнечеткий выдаст "крошка".

Если искать среди: "кит", "тигр", то ВПРнечеткий выдаст "кит".
