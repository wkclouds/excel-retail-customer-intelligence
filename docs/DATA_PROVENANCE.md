# Data provenance and license

- Dataset: **Online Retail II**.
- Creator: **Daqing Chen**.
- Repository/publisher: **UCI Machine Learning Repository**.
- Record: https://archive.ics.uci.edu/dataset/502/online+retail+ii
- DOI: https://doi.org/10.24432/C5CG6D
- Direct download: https://archive.ics.uci.edu/static/public/502/online%2Bretail%2Bii.zip
- Download date on the user's laptop: **2026-09-25**, Asia/Karachi.
- Source workbook: `online_retail_II.xlsx`.
- Dataset coverage: **2009-12-01 to 2011-12-09**; precise observed timestamps appear in `data/processed/analysis.json`.
- Currency: pounds sterling (GBP), as documented by UCI.
- License: **Creative Commons Attribution 4.0 International**, https://creativecommons.org/licenses/by/4.0/ .

Citation: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D.

UCI explicitly permits sharing and adaptation with attribution. Original source data and derived data remain under CC BY 4.0. The MIT code license does not relicense the source data. This project normalizes field names and identifiers, removes exact cross-sheet import overlap using all eight fields plus occurrence number, defines purchase/return populations, aggregates sales and creates customer scores/cohorts. It retains other repeated lines. Neither UCI nor the dataset creator endorses this project.

## Integrity

Original ZIP SHA-256:

```text
572e36277c2390fbfde10664750731e0a86f55e33470d91919085f0408e67bfb
```

Extracted source workbook SHA-256:

```text
bcbe73b35f5b7babf197fb0cb983a11f5d9ff929078d4aa53d171b1f2df2e980
```

The ZIP integrity check passed. The workbook contains eight columns and two sheets, totaling 1,067,371 rows. Exact worksheet ranges, repetition checks and original header spellings are saved with the processed outputs.

## Download and storage

Download the official ZIP, then extract only `online_retail_II.xlsx` into `data/raw`. Alternatively run `py scripts/prepare_data.py --download`. The script requires the original filename. It reads both sheets and checks the expected total before calculating results.

The original source workbook and ZIP are excluded from version control. This reduces repository size and avoids storing two copies of a large source. Processed tables and the workbook are redistributed with the attribution above. No private customer identities are added to the public dataset's numeric identifiers.
