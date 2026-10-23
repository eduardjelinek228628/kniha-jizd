# Kniha jízd

Jednoduchá responzivní stránka pro zapisování jízd autem. Data se ukládají do Google Tabulky (nový list „Kniha jízd“, vytvořený kopií vzoru „108P-Opicake(Honza)“), export do .xlsx drží formát vzoru.

- `index.html` + `config.js`: stránka (GitHub Pages)
- `apps-script/Code.gs`, `apps-script/appsscript.json`: backend v Google Apps Scriptu

## Nasazení backendu (jednorázově, cca 5 minut)

1. Otevři tabulku → **Rozšíření → Apps Script**.
2. Obsah `Code.gs` vlož do souboru `Code.gs` (přepiš výchozí kód).
3. V nastavení projektu (ozubené kolo) zapni **Zobrazit soubor manifestu appsscript.json** a přepiš ho obsahem `appsscript.json`.
4. **Nasadit → Nové nasazení → Webová aplikace**: Spustit jako **já**, Přístup **kdokoli**. Potvrď oprávnění (při varování „Google nezkontroloval“ klikni na *Upřesnit → Přejít*).
5. Zkopíruj adresu končící na `/exec` a vlož ji do `config.js` (`window.KNIHA_JIZD_API`), nebo ji vlož přímo na stránce při prvním otevření.

Po změně kódu skriptu je potřeba **Nasadit → Spravovat nasazení → upravit → Nová verze**.

## Poznámky

- Zdrojem pravdy je skrytý list `_kniha_data`; list „Kniha jízd“ se z něj vykresluje při každé změně. Ručně do „Knihy jízd“ nepiš, přepíše se.
- Sloupec **Řidič** je přidaný vpravo za tabulku vzoru (sloupec P).
- Kdo zná odkaz na stránku, může zapisovat (bez hesla, tak to bylo zadáno).
