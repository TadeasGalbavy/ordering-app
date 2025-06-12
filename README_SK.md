# Objednávací nástroj pre Excel (GUI aplikácia)

## Popis

Tento nástroj slúži na automatizované spracovanie objednávok zo špecificky štruktúrovaného Excel súboru. 
Umožňuje jednoducho a rýchlo vypočítať množstvo produktov na objednanie podľa aktuálnej skladovej situácie, typu produktu a obchodnej logiky. 
Aplikácia má prehľadné GUI rozhranie a dva režimy výpočtu – „vykrytie objednávok“ a „objednávanie na sklad“.

---

## Funkcie

- Dva režimy výpočtu:
  - Vykrytie objednávok
  - Objednávanie na sklad podľa koeficientov
- Zohľadnenie:
  - Bestsellerov
  - Dopredaja
  - Individuálnych koeficientov
    
- Automatické zvýraznenie dát vo výstupe (farby, orámovanie) - pre lepšiu orientáciu v súbore
- Výpočet odporúčanej objednávky podľa nastavenej logiky
- Možnosť zadania vlastných parametrov
- Prehľadné GUI pre koncových používateľov bez potreby programovania
- Možnosť exportu upraveného Excel súboru

---

## Použitie

1. Spusť aplikáciu:
   ```bash
   python objednavanie.py
   ```

2. Vyber si režim:
   - **Vykrytie objednávok** (s voliteľným zahrnutím bestsellerov)
   - **Objednávanie na sklad** (zadaj koeficienty pre bestseller a bežný tovar)

3. Vyber vstupný `.xlsx` súbor so štruktúrovanými dátami.

4. Po spracovaní budeš vyzvaný na uloženie výstupného súboru.

> Upozornenie: Aplikácia pracuje so špecifickým formátom vstupu prispôsobeným konkrétnemu e-shopu. Nie je určená ako univerzálne riešenie pre akékoľvek dáta.

---

## Závislosti

- `pandas`
- `openpyxl`
- `tkinter` (súčasť štandardnej knižnice pre Python na Windows)

Odporúčaný spôsob inštalácie:

```bash
pip install -r requirements.txt
```

---

## Autor a poznámka

Aplikáciu som sám navrhol, backend vytvoril s pomocou AI a kompletne otestoval. Je plne funkčná a používa sa na objednávanie pre e-shop, ktorý aktuálne pôsobí v 6 krajinách.

Projekt slúži ako ukážka reálneho využitia automatizácie v e-commerce prostredí.
