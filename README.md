# Automatic Email From Excel
# Automatické generování a zasílání emailů z dat v excelu

Tento VBA skript automatizuje proces vytváření e-mailu pomocí aplikace Microsoft Outlook, který obsahuje informace o nové objednávce, včetně adresy a popisu, a poskytuje příjemci možnost potvrdit začátek nebo ukončení práce prostřednictvím kliknutelných tlačítek.

## Předpoklady

1. **Listy Excelu**:
    - **`InfoOdesilatel`**: Obsahuje informace o odesílatelích (iniciály, jméno, příjmení, e-mail).
    - **`InfoPrijemce`**: Obsahuje informace o příjemcích (jméno a e-mail).
    - **`Seznam požadavků`**: List obsahující objednávky. Skript čte data z tohoto listu.

2. **Outlook**: Skript vyžaduje aplikaci Microsoft Outlook k odesílání e-mailů.

## Funkce

### `VytvorEmailSObjednavkouATlacitky`

Automaticky generuje e-mail s HTML tělem zprávy, které obsahuje:
- Číslo objednávky
- Popis objednávky
- Adresu (ulice, číslo objektu a bytu)
- Tři tlačítka pro potvrzení akce (začátek práce, ukončení práce, třetí možnost)

### Příjemce

- E-mail je zaslán příjemci, jehož e-mailová adresa je načtena z listu `InfoPrijemce`.

## Postup

1. Otevřete sešit Excel obsahující všechny potřebné listy.
2. Vyberte buňku s objednávkou v listu `Seznam požadavků`.
3. Spusťte makro `VytvorEmailSObjednavkouATlacitky`.
4. E-mail bude automaticky vytvořen a zobrazen, připraven k odeslání.

## Poznámky

- Pokud nejsou nalezeni příjemce nebo odesílatel, zobrazí se upozornění.
- Makro lze upravit, aby e-mail nebyl zobrazen, ale přímo odeslán změnou `.Display` na `.Send`.
