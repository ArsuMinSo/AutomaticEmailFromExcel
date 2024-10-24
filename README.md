# Automatic Email from Excel

## Automatické generování a zasílání e-mailů z dat v Excelu

Tento VBA skript automatizuje proces vytváření e-mailu pomocí aplikace Microsoft Outlook. E-mail obsahuje informace o nové objednávce, včetně adresy a popisu, a umožňuje příjemci potvrdit zahájení nebo ukončení práce prostřednictvím kliknutelných tlačítek.

## Předpoklady

### Listy Excelu:
- **`InfoOdesilatel`**: Obsahuje informace o odesílatelích (iniciály, jméno, příjmení, e-mail, telefon atd.).
- **`InfoPrijemce`**: Obsahuje informace o příjemcích (jméno a e-mail).
- **`Seznam požadavků`**: List obsahující seznam objednávek, ze kterého skript čte data potřebná k sestavení e-mailu.

### Microsoft Outlook:
- Tento skript vyžaduje aplikaci Microsoft Outlook pro vytváření a odesílání e-mailů.

## Funkce

### `VytvorEmailSObjednavkouATlacitky`

Tato procedura automaticky generuje e-mail s HTML formátováním, který obsahuje:
- **Číslo objednávky**: Unikátní identifikátor objednávky.
- **Popis objednávky**: Stručný popis práce nebo úkolu, který je objednáván.
- **Adresa**: Informace o adrese, včetně ulice, čísla objektu a případně čísla bytu.
- **Tlačítka pro potvrzení**: HTML tlačítka, která příjemci umožňují potvrdit přijetí objednávky, zahájení práce nebo ukončení práce.
  - **Potvrdit přijetí objednávky**
  - **Potvrdit realizaci objednávky**

### Přehled klíčových kroků

1. **Načtení údajů o odesílateli a příjemci**:
   - Z listu `InfoOdesilatel` se načtou iniciály odesílatele, podle kterých se vyhledá kompletní jméno, e-mail a další kontaktní informace.
   - Z listu `InfoPrijemce` se načte jméno příjemce a podle něj se vyhledá e-mailová adresa.

2. **Načtení údajů o objednávce**:
   - Skript čte číslo objednávky, popis objednávky a adresu z aktivního řádku v listu `Seznam požadavků`.

3. **Vytvoření e-mailu**:
   - Na základě načtených dat se sestaví HTML tělo e-mailu obsahující všechny relevantní informace.
   - E-mail také obsahuje tři tlačítka pro potvrzení různých akcí spojených s objednávkou.

### HTML Struktura E-mailu

- **Hlavička**: Obsahuje pozdrav a základní informace o nové objednávce.
- **Tělo zprávy**: Zahrnuje adresu a popis objednávky.
- **Tlačítka**: Kliknutelné HTML tlačítka, která automaticky vytvoří odpověď na e-mail s příslušným předmětem, včetně příslušného stavu objednávky.
  - **Zelené tlačítko**: Potvrzení přijetí objednávky.
  - **Modré tlačítko**: Potvrzení realizace objednávky.
  - **Červené tlačítko**: Třetí možnost akce.
- **Podpis**: Kontaktní údaje odesílatele.
- **Patička**: Upozornění o automatickém generování e-mailu.

## Použití

### Postup pro vytvoření e-mailu

1. **Otevřete Excel**: Ujistěte se, že máte otevřený sešit s listy `InfoOdesilatel`, `InfoPrijemce` a `Seznam požadavků`.
   
2. **Vyberte řádek s objednávkou**: Klikněte na buňku v listu `Seznam požadavků`, která obsahuje objednávku, pro kterou chcete vygenerovat e-mail.

3. **Spusťte makro**: Spusťte makro `VytvorEmailSObjednavkouATlacitky`.

4. **Zkontrolujte e-mail**: E-mail bude automaticky vytvořen a zobrazen v aplikaci Microsoft Outlook. Můžete ho před odesláním zkontrolovat nebo rovnou odeslat.

### Poznámky k používání

- **Chybové hlášky**: Pokud není nalezen odesílatel nebo příjemce, zobrazí se chybová zpráva a proces se zastaví.
- **Odesílání e-mailu**: E-mail je výchozí nastaven na zobrazení v Outlooku (`.Display`). Pokud chcete, aby byl automaticky odeslán, změňte tuto funkci na `.Send`.
  
## Výhody

- **Automatizace**: Šetří čas tím, že automaticky vytváří e-maily z dat v Excelu.
- **Interaktivní e-maily**: Umožňuje příjemcům snadno potvrdit různé fáze objednávky pomocí kliknutelných tlačítek.
- **Personalizace**: Každý e-mail je automaticky personalizován podle údajů o objednávce, příjemci a odesílateli.

## Příklad použití

Předpokládejme, že v listu `Seznam požadavků` máte následující objednávku:

| Číslo objednávky | Ulice      | Číslo objektu | Číslo bytu | Popis objednávky      | Iniciály odesílatele | Příjemce |
|------------------|------------|---------------|------------|-----------------------|----------------------|----------|
| 12345            | Dlouhá     | 12            | 5          | Oprava osvětlení      | JN                   | Novák    |

Po spuštění makra se vytvoří následující e-mail:

- **Předmět**: Nová objednávka 12345 Dlouhá 12_5
- **Tělo**:
  - Objednávka je na ulici Dlouhá v bytovém domě s číslem 12 a číslem bytu 5.
  - Stručný popis objednávky: Oprava osvětlení.
  - Tlačítka pro potvrzení přijetí a realizace objednávky.

