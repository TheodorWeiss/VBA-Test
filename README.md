Да, логика такая: каждая CUBE-формула потенциально обращается к кубу, хотя Excel частично кэширует запросы. Поэтому много формул = тормоза.

Для твоей первой цели нужен один CUBEMENGE / CUBESET, который сразу попросит у куба только Top/Flop категории.

Цель

Не все категории, а только:

Top 5 Kategorien nach:
[Abweichung Feld 1] + [Abweichung Feld 2]
Flop 5 Kategorien nach:
[Abweichung Feld 1] + [Abweichung Feld 2]

Немецкие функции

Английская	Немецкая
CUBESET	CUBEMENGE
CUBERANKEDMEMBER	CUBERANGEELEMENT
CUBEVALUE	CUBEWERT

1. Flop 5 категорий

В отдельную ячейку, например B2:

=CUBEMENGE(
"Название_соединения";
"BOTTOMCOUNT(
    [Kategorie].[Kategorie].[Kategorie].MEMBERS,
    5,
    ([Measures].[Abweichung_1] + [Measures].[Abweichung_2])
)";
"Flop 5 Kategorien"
)

2. Top 5 категорий

Например E2:

=CUBEMENGE(
"Название_соединения";
"TOPCOUNT(
    [Kategorie].[Kategorie].[Kategorie].MEMBERS,
    5,
    ([Measures].[Abweichung_1] + [Measures].[Abweichung_2])
)";
"Top 5 Kategorien"
)

3. Вывести элементы из набора

Под Flop 5:

=CUBERANGEELEMENT("Название_соединения";$B$2;1)

ниже:

=CUBERANGEELEMENT("Название_соединения";$B$2;2)

и так до 5.

Для Top 5 аналогично, только ссылка на $E$2.

4. Вывести значение суммы Abweichung

Рядом с категорией:

=CUBEWERT(
"Название_соединения";
A4;
"[Measures].[Abweichung_1]"
)
+
CUBEWERT(
"Название_соединения";
A4;
"[Measures].[Abweichung_2]"
)

Где A4 — ячейка с категорией из CUBERANGEELEMENT.

Что нужно заменить

Тебе надо взять из уже созданных CUBE-формул:

1. Название соединения
    Обычно выглядит как "ThisWorkbookDataModel" или название Cloud NDW connection.
2. MDX-путь категории
    Что-то вроде:

[Artikelhierarchie].[Kategorie].[Kategorie].MEMBERS

3. Имена двух measures Abweichung
    Например:

[Measures].[Abweichung Umsatz]
[Measures].[Abweichung Absatz]

Важно

Если твои Abweichung уже считаются как measures в кубе, то выражение:

([Measures].[Abweichung_1] + [Measures].[Abweichung_2])

должно работать прямо внутри TOPCOUNT / BOTTOMCOUNT.

Начни именно с Flop 5 Kategorien. Когда это заработает, Top 5 делается почти копированием с заменой BOTTOMCOUNT на TOPCOUNT.
