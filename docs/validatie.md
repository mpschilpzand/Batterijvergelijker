# Validatie Pythonmodel tegen Excel

## Stand 1 juli 2026

De formules uit `Huishoudprofiel verrekenprijs 2025.xlsx` zijn direct naar
Python vertaald. Microsoft Excel is gebruikt om afzonderlijke werkkopieën
volledig te herberekenen.

Gecontroleerde scenario's:

1. standaardwaarden;
2. jaarverbruik 5.200 kWh en zonne-opwek 3.000 kWh;
3. batterijcapaciteit 13,5 kWh, vermogen 3 kW, beginlading 2 kWh,
   dagexport 1,5 kWh en LEND-verkoopprijs € 0,16/kWh;
4. gewijzigde energiebelasting, btw, leveranciersopslag en vaste kosten.

Per scenario zijn twaalf waarden vergeleken:

- netafname, teruglevering en marktkosten zonder batterij;
- netafname, teruglevering en marktkosten met de simpele batterij;
- netafname, teruglevering en marktkosten met de LEND-batterij;
- totale OPEX inclusief btw voor alle drie de scenario's.

Resultaat: alle 48 vergelijkingen hebben een maximaal absoluut verschil van
`0` met de door Excel berekende waarden.

## Reproduceerbare controles

De onafhankelijke regressietests:

```shell
python3 -m unittest discover -s tests -v
```

Een door Excel herberekende werkmap vergelijken:

```shell
python3 tools/validate_against_excel.py output/excel-validation.xlsx
```

## Afbakening

- Batterijverlies (`Aannames!B32`) wordt in Excel nergens in een rekenformule
  toegepast. Python doet dat daarom eveneens niet.
- Het profiel, de EPEX-prijzen, de simpele dagpieken en de LEND-piekselecties
  zijn vaste reeksen in het huishoudmodel en worden daaruit gelezen.
- `Scenarios verrekenprijs batterij.xlsx` is de bron van een deel van deze
  vaste LEND-referentiegegevens. De afzonderlijke Heeten- en
  Aardenhuizen-simulaties zijn nog niet als zelfstandig Pythonmodel vertaald;
  ze zijn niet nodig voor de huidige huishoudtool.

## Vaste LEND-prijzen

Het huishoudmodel bevat nu twee afzonderlijke vaste verkoopprijzen:

- `Aannames!B34`: batterij-export in het LEND-scenario;
- `Aannames!B38`: directe zonne-export in het LEND-scenario.

De scenario's zonder batterij en met de privébatterij blijven voor beide
exportstromen de EPEX-uurprijs gebruiken.
