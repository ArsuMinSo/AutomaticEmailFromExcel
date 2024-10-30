# Automatic Email from Excel

## Automatické generování a zasílání e-mailů z dat v Excelu

Toto makro automatizuje vytvoření e-mailu s objednávkou v aplikaci Outlook. Obsahuje objednávkový formulář, detaily objednávky, kontaktní informace a interaktivní tlačítka, která příjemce umožní potvrdit přijetí, zahájení nebo dokončení objednávky.

---

### Obsah
1. [Předpoklady](#předpoklady)
2. [Struktura kódu](#struktura-kódu)
3. [Popis funkcí a proměnných](#popis-funkcí-a-proměnných)
4. [Postup generování e-mailu](#postup-generování-e-mailu)
5. [Pata e-mailu](#pata-e-mailu)

---

## Předpoklady
- **Data**: Makro předpokládá, že data o odesílateli a příjemci jsou na listech `InfoOdesilatel`, `InfoPrijemce` a `Seznam požadavků`.
- **Outlook**: Pro úspěšné spuštění makra musí být na počítači nainstalována aplikace Outlook.

---

## Struktura kódu
Makro sestává ze tří hlavních sekcí:
1. **Inicializace proměnných a načtení dat** - Proměnné obsahují textové řetězce s údaji o objednávce a kontaktní informace odesílatele a příjemce.
2. **Sestavení těla e-mailu** - HTML formátování pro tělo e-mailu s interaktivními tlačítky pro zpětnou vazbu.
3. **Odeslání e-mailu** - Nastavení a vytvoření nové zprávy v Outlooku.

---

## Popis funkcí a proměnných

### 1. Inicializace proměnných a načtení dat
Makro načítá klíčové informace o objednávce a záznamy o odesílateli a příjemci:
- `CisloObjednavky`, `Ulice`, `CisloObjektu`, `CisloBytu`, `Kategorie`, `PopisObjednavky`, `KontaktHlaseneho`, `Hlasil` – údaje o objednávce.
- `TitulOdesilatele`, `JmenoOdesilatele`, `PrijmeniOdesilatele`, `HodnostOdesilatele` atd. – informace o odesílateli.
- `EmailPrijemce` – e-mail příjemce, který je identifikován podle záznamu `Prijemce`.

**Příklad načítání z listu:**
```vba
Set InfoListOdesilatel = Sheets("InfoOdesilatel")
Set ShodaOdesilatel = InfoListOdesilatel.Columns(2).Find(What:=Zapsal, LookAt:=xlWhole)
```

### 2. Sestavení těla e-mailu
Makro sestaví tělo e-mailu s použitím HTML formátování a podmínek:
- **Tělo zprávy** - Proměnná `TelozpravyMisto` upravuje obsah na základě toho, zda je vyplněno číslo bytu.
- **Podpis a kontaktní údaje** - Přidávají kontaktní údaje odesílatele, případně jeho webovou stránku.
- **Interaktivní tlačítka** - Umožňují příjemci potvrdit přijetí, zahájení nebo dokončení objednávky. Každé tlačítko obsahuje vlastní odkaz s odlišným textem v předmětu i těle zprávy.

**Příklad tlačítka:**
```vba
TeloTlacitka = "<a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-PŘIJATO'>Potvrdit přijetí objednávky</a>"
```

### 3. Odeslání e-mailu
Po sestavení e-mailu je tento zobrazen v Outlooku jako nový e-mail k odeslání.
- **Vytvoření e-mailu** – pomocí `OutlookApp.CreateItem(0)` a následného nastavení příjemce, předmětu a těla zprávy.
- **Metoda `.Display`** – zobrazí e-mail před odesláním, aby jej uživatel mohl upravit.

---

## Postup generování e-mailu
1. **Načtení dat**: Makro načítá informace z aktivního řádku a odpovídajících listů.
2. **Sestavení těla e-mailu**: Generuje obsah na základě údajů o objednávce a kontaktních údajů.
3. **Přidání tlačítek**: Do e-mailu se vloží tlačítka pro potvrzení.
4. **Odeslání e-mailu**: E-mail je zobrazen v aplikaci Outlook.

---
