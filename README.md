
## `copy_paste_finder.py`

Detects duplicated, copied, and fabricated data in Excel datasets.

### Strategies

Runs eight strategies:

- **A. Duplicate rows** – pairs of rows sharing ≥2 high-entropy values  
- **B. Repeated column sequences** – same run of values at two positions in a column  
- **C. Terminal digit test** – per-column uniformity test on last significant digit  
- **D. Periodic row duplication** – fixed-lag block copying (e.g. every 101 rows)  
- **E. Cosine similarity** – near-identical rows on repetitive columns only  
- **F. Fingerprint gap** – dominant gap between recurring row fingerprints  
- **G. Collinearity matrix** – column pairs with |r| ≥ 0.98 (explains E false positives)  
- **H. Modular block count** – counts exact-match row pairs per candidate period; writes `output.pdf` with heatmaps and autocorrelogram
- **I. Excel format forensics** – file-layer analysis independent of data values: internal metadata (creator/modifier/timestamps/revision),...

###  Benchmark files recognized
```
  Südhof / Neurexin2 pnas.2300363120.sd01.xlsx 
  Pruitt / Spider    Dumicola familiarity wide.xlsx
  Hawk / Owl         Dryad_dataset.xlsx
  Gino / Tax         Tax_Study_STUDY_1_2010-07-13.xlsx
  Attar / PREVENT    TAHA8.xlsx
```

### Installation

```bash
pip install openpyxl matplotlib numpy
```

### Usage

#### Heuristic mode

```bash
python copy_paste_finder.py <file.xlsx>
```

#### With visualization

```bash
python copy_paste_finder.py <file.xlsx> --plot
```

#### With specific columns for Strategy H

```bash
python copy_paste_finder.py <file.xlsx> --plot --plot-cols WBC,Hb,Plt,BUN,Cr,Na
```

#### All options

```bash
python copy_paste_finder.py <file.xlsx> \
    [--sheet SHEET] \
    [--min-suspicion low|medium|high] \
    [--plot] \
    [--plot-cols COL1,COL2,...] \
    [--plot-period 101] \
    [--min-period 50] \
    [--max-period 250] \
    [--max-lag 300] \
    [--out output.pdf] \
    [--forensics]
```


The [entropy calculation](https://github.com/markusenglund/copy-paste-detective?ref=sciencedetective.org) depends on an idea of Markus Eglund and may not be used without his permission.  
All other modules have been developed during the [PREVENT-TAHA8](https://www.bmj.com/content/391/bmj-2024-083382/rapid-responses) trial as discussed at the [blog](https://www.wjst.de/blog/sciencesurf/2025/11/is-there-a-data-agnostic-method-to-find-repetitive-data-in-clinical-trials/) of the author and are freely available.
