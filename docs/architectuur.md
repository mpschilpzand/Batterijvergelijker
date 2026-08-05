# Architectuur

```text
Gebruikersinvoer
      |
      v
Streamlit-scherm
      |
      v
Python-rekenmodel <---- uurprofiel en vaste referentiereeksen uit XLSX
      |
      v
Resultaten, grafieken en export
```

## Voorgestelde techniek

- interface: Streamlit
- rekenmodel: Python (`model/calculator.py`)
- XLSX-lezer: Python-standaardbibliotheek (`model/xlsx_reader.py`)
- validatie: Microsoft Excel, uitsluitend tijdens ontwikkeling
- rapport: Streamlit-tabellen en grafieken

De webapp heeft geen spreadsheet-engine en geen Microsoft Office-licentie
nodig. Excel blijft voorlopig het referentiemodel voor regressietests.

## Uitgangspunt

De applicatie leest het bronmodel alleen. Gebruikersinvoer wordt als overrides
aan de Pythonfunctie doorgegeven en wijzigt de werkmap niet.
