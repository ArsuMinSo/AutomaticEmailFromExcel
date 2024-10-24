# Automatic Email From Excel

## Automatické generování a zasílání emailů z dat v Excelu

Tento VBA skript automatizuje proces vytváření e-mailu pomocí aplikace Microsoft Outlook. E-mail obsahuje informace o nové objednávce, včetně adresy a popisu, a poskytuje příjemci možnost potvrdit zahájení nebo ukončení práce prostřednictvím kliknutelných tlačítek.

### Předpoklady

1. **Listy Excelu**:
    - **`InfoOdesilatel`**: Obsahuje informace o odesílatelích (iniciály, jméno, příjmení, e-mail atd.).
    - **`InfoPrijemce`**: Obsahuje informace o příjemcích (jméno a e-mail).
    - **`Seznam požadavků`**: List, který obsahuje jednotlivé objednávky, z něhož skript čte data.

2. **Outlook**: Skript vyžaduje aplikaci Microsoft Outlook pro zasílání e-mailů.

### Funkce

#### `VytvorEmailSObjednavkouATlacitky`

Tato funkce automaticky generuje e-mail s HTML tělem, který obsahuje následující části:

- **Číslo objednávky**: Zobrazuje číslo objednávky v těle e-mailu.
- **Popis objednávky**: Krátký popis objednané práce.
- **Adresa**: Včetně ulice, čísla objektu a bytu (pokud je relevantní).
- **Tlačítka pro akce**: Tři tlačítka umožňují příjemci potvrdit různé akce:
  1. Potvrzení přijetí objednávky.
  2. Potvrzení zahájení realizace objednávky.
  3. Třetí možnost pro přizpůsobení (např. doplňková akce).
  
#### Příjemce

- E-mail je automaticky adresován příjemci, jehož e-mailová adresa je načtena z listu `InfoPrijemce`.

### Postup

1. **Příprava dat**:
   - Otevřete sešit Excel, který obsahuje listy `InfoOdesilatel`, `InfoPrijemce` a `Seznam požadavků`.
   
2. **Výběr objednávky**:
   - Vyberte příslušnou objednávku v listu `Seznam požadavků` (aktivní buňka v řádku musí obsahovat data objednávky).

3. **Spuštění makra**:
   - Spusťte makro `VytvorEmailSObjednavkouATlacitky` (například z Editoru VBA nebo přiřazeného tlačítka).

4. **Vytvoření a zobrazení e-mailu**:
   - Makro automaticky vytvoří e-mail s předvyplněnými údaji a zobrazí ho v aplikaci Outlook. E-mail bude připraven k manuálnímu odeslání. 
   - Pokud chcete e-mail automaticky odeslat, můžete změnit `.Display` na `.Send`.

### Poznámky

- Pokud nejsou nalezeni příjemce nebo odesílatel podle zadaných údajů, skript zobrazí upozornění a ukončí se.
- **Přizpůsobení**: Makro lze přizpůsobit podle potřeby, například přidáním dalších informací do těla e-mailu nebo přidáním dalších funkcionalit.
  
