Да, вижу. Для категории правильная иерархия, судя по формуле сверху, такая:

[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe]

Попробуй сначала самый простой тест CUBEMENGE:

=CUBEMENGE("Cloud NDW Prod111";"{[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe].&[26 - WURST]}";"Test")

Потом рядом вытащи элемент:

=CUBERANGEELEMENT("Cloud NDW Prod111";A1;1)

где A1 — ячейка с CUBEMENGE.

Если это заработает, тогда список всех Hauptwarengruppe:

=CUBEMENGE("Cloud NDW Prod111";"[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe].MEMBERS";"Alle Hauptwarengruppen")

А уже Flop 5 по сумме двух Abweichung-полей должен выглядеть примерно так:

=CUBEMENGE(
"Cloud NDW Prod111";
"BOTTOMCOUNT(
[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe].MEMBERS;
5;
([Measures].[Abweichung_1] + [Measures].[Abweichung_2])
)";
"Flop 5 Hauptwarengruppen"
)

Но внутри MDX, возможно, нужны запятые, а не ;. Тогда вариант:

=CUBEMENGE(
"Cloud NDW Prod111";
"BOTTOMCOUNT(
[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe].MEMBERS,
5,
([Measures].[Abweichung_1] + [Measures].[Abweichung_2])
)";
"Flop 5 Hauptwarengruppen"
)

Top 5:

=CUBEMENGE(
"Cloud NDW Prod111";
"TOPCOUNT(
[Artikel].[Nettowarengruppenstruktur].[Hauptwarengruppe].MEMBERS,
5,
([Measures].[Abweichung_1] + [Measures].[Abweichung_2])
)";
"Top 5 Hauptwarengruppen"
)

Тебе надо заменить только:

[Measures].[Abweichung_1]
[Measures].[Abweichung_2]

на точные имена твоих двух готовых полей.

Начни с самого первого теста с {...[26 - WURST]}. Если даже он даёт #NV, значит CUBEMENGE в твоём кубе не принимает такой MDX-set, и тогда надо идти через Pivot-Wertfilter, а не через CUBESET.
