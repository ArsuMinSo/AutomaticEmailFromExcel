# Automatic Email from Excel

## Automatické generování a zasílání e-mailů z dat v Excelu

Toto makro slouží k automatizovanému vytvoření a odeslání e-mailu prostřednictvím aplikace Microsoft Outlook. E-mail obsahuje informace o objednávce a interaktivní tlačítka, která příjemci umožňují potvrdit přijetí a realizaci objednávky.

## Popis funkcionality

1. **Získání dat z Excelu**: 
   - Makro pracuje s daty uloženými v různých listech Excelu, konkrétně z listů **"InfoOdesilatel"**, **"InfoPrijemce"** a **"Seznam požadavků"**. 
   - Na základě aktuálně vybrané buňky (řádku) makro načte klíčové informace o objednávce a zpracovává je do e-mailové zprávy.

2. **Dynamické generování těla e-mailu**: 
   - E-mail je automaticky sestaven z předdefinovaných šablon, přičemž se do zprávy doplňují hodnoty z buněk Excelu (např. číslo objednávky, popis, místo objednávky, odesilatel a příjemce).
   - Pokud je vyplněno číslo bytu, do textu zprávy se zahrne i toto číslo.

3. **Generování tlačítek**:
   - Součástí e-mailu jsou dvě tlačítka: jedno pro potvrzení přijetí objednávky a druhé pro potvrzení realizace objednávky.
   - Tlačítka jsou hypertextové odkazy, které otevřou nový e-mail v Outlooku s předvyplněným předmětem a tělem zprávy.

4. **Načtení informací o odesílateli a příjemci**:
   - Na základě iniciál zadaných v listu "Seznam požadavků" makro hledá odpovídající údaje o odesílateli a příjemci v listech "InfoOdesilatel" a "InfoPrijemce".
   - Pokud není odesilatel nebo příjemce nalezen, makro zobrazí chybovou hlášku.

5. **Vytvoření a zobrazení e-mailu**: 
   - Vytvořený e-mail se zobrazí v okně Outlooku pro kontrolu. E-mail je možné buď okamžitě odeslat, nebo před odesláním upravit.

## Vstupní data

Makro čte následující data z listů Excelu:

### Seznam požadavků (hlavní list)
- **Číslo objednávky**: sloupec 3
- **Ulice**: sloupec 4
- **Číslo objektu**: sloupec 5
- **Číslo bytu**: sloupec 6
- **Kategorie**: sloupec 7
- **Popis objednávky**: sloupec 8
- **Zapsal (iniciály odesílatele)**: sloupec 11
- **Příjemce**: sloupec 17

### InfoOdesilatel (list s údaji o odesílateli)
- **Iniciály**: sloupec 2
- **Titul**: sloupec 3
- **Jméno**: sloupec 4
- **Příjmení**: sloupec 5
- **Hodnost**: sloupec 6
- **Odbor**: sloupec 7
- **Oddělení**: sloupec 8
- **Město úřadu**: sloupec 9
- **Adresa úřadu**: sloupec 10
- **Telefon**: sloupec 11
- **E-mail**: sloupec 12
- **Web**: sloupec 13

### InfoPrijemce (list s údaji o příjemci)
- **Příjemce (jméno)**: sloupec 2
- **E-mail**: sloupec 3

## Výstupní data

### Tělo e-mailu
- Obsahuje informace o objednávce včetně:
  - Číslo objednávky
  - Popis objednávky
  - Místo objednávky (ulice, číslo domu a bytu)
- Interaktivní tlačítka pro potvrzení přijetí a realizace objednávky.
- Podpis odesílatele s kontaktními údaji.

### Předmět e-mailu
- "Nová objednávka [Číslo objednávky] [Ulice a číslo domu/bytu]"

## Příklad použití

Toto makro je vhodné pro scénáře, kde je potřeba automatizovat komunikaci ohledně objednávek. Uživatel zadá informace do tabulky a makro na základě těchto informací automaticky vygeneruje a předvyplní e-mail, který je připraven k odeslání.

## Chybová hlášení

- Pokud není nalezen odesílatel podle zadaných iniciál, zobrazí se chybová zpráva: 
  `Nebylo nalezeno žádné příjmení začínající na [iniciály].`
  
- Pokud není nalezen příjemce, zobrazí se upozornění: 
  `Nebyl nalezen žádný příjemce [jméno příjemce].`

## Omezení a předpoklady

- Makro předpokládá, že všechny potřebné údaje jsou správně vyplněny v příslušných listech.
- Pokud nejsou data v listech kompletní, může dojít k chybám nebo nesprávnému fungování makra.
  
## Závěr

Toto VBA makro poskytuje efektivní způsob, jak automatizovat proces vytváření objednávkových e-mailů s využitím dat z Excelu a aplikace Outlook. Správným nastavením dat a šablon e-mailů se zajišťuje konzistentní a rychlá komunikace mezi odesílatelem a příjemcem objednávek.
