# Batterijvergelijking

Workspace voor een rekentool die de jaarlijkse energiekosten van een huishouden
met en zonder thuisbatterij vergelijkt.

## Excelmodellen

De bronbestanden staan in `model/`:

- `Huishoudprofiel verrekenprijs 2025.xlsx`
- `Scenarios verrekenprijs batterij.xlsx`

Excelbestanden worden niet in GitHub opgeslagen. Voor deployment gebruikt de
app daarom `model/runtime_model.json`, een export met alleen de benodigde
uurreeksen en aannames. Genereer die opnieuw uit je lokale Excelbestand met:

```shell
python3 tools/export_runtime_model.py
```

## Python-rekenmodel

`model/calculator.py` bevat een directe Pythonvertaling van:

- het uurprofiel zonder batterij;
- de simpele batterijstrategie;
- de LEND-batterijstrategie;
- het kostenoverzicht inclusief belasting en btw.

Het profiel, de EPEX-prijzen en de vooraf geselecteerde piekuren worden
voorlopig uit de eerste werkmap gelezen. Excel is niet nodig om de berekening
uit te voeren.

Voorbeeld:

```python
from model.calculator import calculate

result = calculate(
    "model/Huishoudprofiel verrekenprijs 2025.xlsx",
    {"annual_usage_kwh": 5000, "battery_capacity": 12},
)
print(result.lend_costs.opex_inc_vat)
```

## Controle

```shell
python3 -m unittest discover -s tests -v
python3 tools/validate_against_excel.py output/excel-validation.xlsx
```

Zie `docs/validatie.md` voor de gecontroleerde scenario's en beperkingen.

## Mappen

- `app/`: gebruikersinterface en koppeling met het Excelmodel
- `config/`: configuratie en koppeling van invoervelden aan Excelcellen
- `docs/`: functionele en technische afspraken
- `model/`: later de definitieve versie van het Excelmodel
- `output/`: gegenereerde rapporten en resultaatbestanden
- `tests/`: controleberekeningen en regressietests

## Eerstvolgende stap

Start het Streamlit-scherm:

```shell
python3 -m pip install -r requirements.txt
streamlit run app/app.py
```

De Pythonberekening is onafhankelijk van een Excel-installatie.
