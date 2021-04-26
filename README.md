# Excel - Inventory Management

Einfaches Inventar-Management-System based on MS Excel und VBA Makros zum Ausleihen und Zurückgeben von Materialien, wie z.Bsp. iPads.

Die Ausleihe und die Rückgabe erfolgt automatisch nach Eingabe der Seriennummer. Dies ermöglicht eine schnelle Ausleihe und Rückgabe mit Bar-Code-Scannern.

## Ausleihe

Um eine Materialie auszuleihen muss der Name der ausleihenden Person und die Seriennummmer eingegeben werden. Nach Eingabe der Seriennummer wird das Material aus der Tabelle der verfügbaren Materialien in die Tabelle der ausgeliehenen Materialien übertragen. Zusätzlich wird ein Eintrag in der Tabelle Ausleih-Historie erstellt.


## Rückgabe

Um eine Materialie zurückzugeben muss der Name der rückgebenden Person und die Seriennummmer eingegeben werden. Nach Eingabe der Seriennummer wird das Material aus der Tabelle der ausgeliehenen Materialien in die Tabelle der verfügbaren Materialien übertragen. Zusätzlich wird ein Eintrag in der Tabelle Ausleih-Historie erstellt.

## Defekt melden

Zusätzlich kann ein Defekt einer Materialie gemeldet werden. Um einen Defekt zu melden, muss der Name der meldenden Person, die Seriennummer und eine Beschreibung des Defekts eingegeben. Danach muss der Button "Defekt melden" gedrückt werden.

## Administration

Im Arbeitsblatt "Administration" können die Namen der Personen und die verfügbaren Materialien eingestellt werden. Das Arbeitsblatt enthält versteckte Buttons, die durch die Eingabe eines Wertes in die Zelle "D1" sichtbar gemacht werden könnnen.

Mit dem Reset-Button werden alle Tabellen in den Ursprungszustand zurückgesetzt.
