Да, это как раз легко и очень полезно 👍
Самый простой и надёжный способ — через Timer.

Как это работает

* Timer возвращает секунды с полуночи
* разница = длительность выполнения макроса

⸻

Добавь в макрос

1. В начало (после объявления переменных):

Dim startTime As Double
Dim elapsedTime As Double
startTime = Timer

⸻

2. В самый конец (перед MsgBox):

elapsedTime = Timer - startTime
' если макрос прошёл через полночь (редко, но правильно учесть)
If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400

⸻

3. В MsgBox добавь:

"Время выполнения: " & _
Int(elapsedTime / 60) & " мин " & _
Round(elapsedTime Mod 60, 1) & " сек"

⸻

Итоговый кусок MsgBox будет такой:

MsgBox "Готово" & vbCrLf & _
       "Год: " & targetYear & vbCrLf & _
       "Неделя: " & targetWeek & vbCrLf & _
       "Ключ: " & yearWeekKey & vbCrLf & _
       "Обновлено: " & changedCount & vbCrLf & _
       "Пропущено: " & skippedCount & vbCrLf & vbCrLf & _
       "Время выполнения: " & _
       Int(elapsedTime / 60) & " мин " & _
       Round(elapsedTime Mod 60, 1) & " сек"

⸻

💡 Маленький инсайт

С твоими OLAP-сводными ты сейчас фактически можешь:
👉 измерять, сколько реально занимает refresh
👉 и оптимизировать (например, с/без RefreshTable)

Если хочешь — можем дальше:

* сравнить 2 варианта (с RefreshTable / без)
* или сделать лог в Excel (замер времени по дням)

Это уже прям уровень “аналитик оптимизирует Excel как систему” 😄
