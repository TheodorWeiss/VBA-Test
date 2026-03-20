Да. Теперь уже можно переходить к практической схеме.

С учётом твоего уточнения про фиксированный вертикальный диапазон по каждой сети, это как раз удобно: для каждого блока строк ты просто задаёшь свой диапазон поиска.

Что будет в итоге

Для колонки H делаем так:
	•	K — первая ли это строка группы в пределах текущего блока сети
	•	L — итог по группе для H: 1 или 0
	•	внизу блока — SUMME(L...)

⸻

1) VBA: две вспомогательные функции

Ниже код для стандартного модуля VBA.
Я специально сделала так, чтобы:
	•	поиск шёл только внутри переданного диапазона
	•	B=1 и B=2 были разными группами
	•	дробные значения сравнивались аккуратно
	•	строки могли идти не подряд

VBA-код

Option Explicit

Private Function SameB(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    If IsNumeric(v1) And IsNumeric(v2) Then
        SameB = (Abs(CDbl(v1) - CDbl(v2)) < 0.000001)
    Else
        SameB = (CStr(v1) = CStr(v2))
    End If
End Function

Private Function IsSimilarArticle(ByVal s1 As Variant, ByVal s2 As Variant) As Boolean
    ' !!! ЗАМЕНИ эту строку на вызов своей готовой функции похожести !!!
    ' Пример:
    ' IsSimilarArticle = SehrAehnlichArtikel(CStr(s1), CStr(s2))
    
    IsSimilarArticle = SehrAehnlichArtikel(CStr(s1), CStr(s2))
End Function

Public Function HasSimilarAbove( _
    ByVal curArtCell As Range, _
    ByVal curBCell As Range, _
    ByVal artRange As Range, _
    ByVal bRange As Range) As Boolean
    
    Dim i As Long
    Dim curRow As Long
    
    If artRange.Rows.Count <> bRange.Rows.Count Then
        HasSimilarAbove = False
        Exit Function
    End If
    
    curRow = curArtCell.Row
    HasSimilarAbove = False
    
    For i = 1 To artRange.Rows.Count
        If artRange.Cells(i, 1).Row < curRow Then
            If SameB(curBCell.Value, bRange.Cells(i, 1).Value) Then
                If IsSimilarArticle(curArtCell.Value, artRange.Cells(i, 1).Value) Then
                    HasSimilarAbove = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Function AllSimilarHaveH1( _
    ByVal curArtCell As Range, _
    ByVal curBCell As Range, _
    ByVal artRange As Range, _
    ByVal bRange As Range, _
    ByVal hRange As Range) As Boolean
    
    Dim i As Long
    Dim foundAny As Boolean
    Dim sameGroup As Boolean
    
    If artRange.Rows.Count <> bRange.Rows.Count _
       Or artRange.Rows.Count <> hRange.Rows.Count Then
        AllSimilarHaveH1 = False
        Exit Function
    End If
    
    foundAny = False
    
    For i = 1 To artRange.Rows.Count
        sameGroup = False
        
        If SameB(curBCell.Value, bRange.Cells(i, 1).Value) Then
            If artRange.Cells(i, 1).Row = curArtCell.Row Then
                sameGroup = True
            ElseIf IsSimilarArticle(curArtCell.Value, artRange.Cells(i, 1).Value) Then
                sameGroup = True
            End If
        End If
        
        If sameGroup Then
            foundAny = True
            
            If hRange.Cells(i, 1).Value <> 1 Then
                AllSimilarHaveH1 = False
                Exit Function
            End If
        End If
    Next i
    
    AllSimilarHaveH1 = foundAny
End Function


⸻

2) Что нужно поменять у тебя

В этом месте:

IsSimilarArticle = SehrAehnlichArtikel(CStr(s1), CStr(s2))

нужно оставить имя твоей реальной функции похожести.

Если у неё другой набор аргументов, просто подставь его туда.

Например, если у тебя функция такая:

MySimilarity(cell1, cell2)

то будет:

IsSimilarArticle = MySimilarity(CStr(s1), CStr(s2))


⸻

3) Формулы в Excel

Ниже я предполагаю, что:
	•	колонка с артикулом/названием = A
	•	вес = B
	•	флаг = H
	•	вспомогательные колонки = K и L

Если артикул у тебя не в A, просто заменишь букву.

⸻

Блок 1: строки 7:44

K7

Проверка: первая ли это строка группы внутри блока 7:44

=WENN(HasSimilarAbove($A7;$B7;$A$7:$A$44;$B$7:$B$44);0;1)

Протянуть до K44.

⸻

L7

Итог по группе для H

=WENN($K7=1;WENN(AllSimilarHaveH1($A7;$B7;$A$7:$A$44;$B$7:$B$44;$H$7:$H$44);1;0);0)

Протянуть до L44.

⸻

Сумма по блоку

Например, внизу:

=SUMME($L$7:$L$44)


⸻

Блок 2: строки 46:99

Для следующей сети просто меняешь диапазон.

K46

=WENN(HasSimilarAbove($A46;$B46;$A$46:$A$99;$B$46:$B$99);0;1)

L46

=WENN($K46=1;WENN(AllSimilarHaveH1($A46;$B46;$A$46:$A$99;$B$46:$B$99;$H$46:$H$99);1;0);0)

И протягиваешь вниз до 99.

⸻

4) Как это работает на твоей логике

Для каждой строки в блоке:

K

проверяет:
	•	есть ли выше в пределах блока
	•	похожая строка
	•	с тем же B

Если есть → 0
Если нет → 1

То есть K помечает первую строку группы.

L

считает только для первых строк:
	•	если все похожие строки с тем же B внутри блока имеют H=1 → 1
	•	если хотя бы одна имеет H<>1 → 0

⸻

5) Что важно

B=1 и B=2

Они не смешиваются, потому что сравнение идёт по точному B.

То есть:
	•	1 — одна группа
	•	2 — другая группа

Это уже учтено.

Неполные группы

Тоже учтены:
	•	найден один дробный артикль
	•	у него H=1
	•	похожих с H=0 не найдено

→ группа даст 1

Именно так, как ты хотел.

⸻

6) Ограничение, которое надо понимать

Если твоя функция похожести несимметрична или ведёт себя нестабильно, возможны странности.

То есть желательно, чтобы для двух строк:
	•	если A похож на B,
	•	то и B похож на A.

Иначе группировка может чуть “плыть”.

⸻

7) Что я бы советовала практически

Сначала проверить на маленьком блоке 7:44:
	•	руками отфильтровать K=1
	•	посмотреть, совпадают ли “первые строки групп” с твоим ожиданием
	•	потом проверить L

Если всё ок, просто копируешь схему на следующие блоки сети с заменой диапазона.

⸻

8) Если хочешь чуть удобнее

Можно сделать ещё третий служебный столбец, чисто для визуального контроля, например:
	•	M = номер/маркер группы или текст "first" / "skip"

Но для расчёта это не обязательно.

⸻

9) Мой совет по следующему шагу

Сейчас лучше всего:
	1.	вставить VBA,
	2.	протестировать на блоке 7:44,
	3.	если нужно, я помогу сразу адаптировать под твои реальные колонки, если ты пришлёшь:

	•	где именно артикул,
	•	где B,
	•	где H,
	•	и как точно называется твоя функция похожести.
