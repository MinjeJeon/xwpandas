# xwpandas

xwpandas is high performance Excel IO tools for pandas DataFrame. xp.save function in xwpandas saves large(100k~ rows) DataFrame to xlsx format 2x~3x faster than xlsxwriter engine in pandas. 

## Installation
xwpandas can be installed using `pip` package manager.

```bash
$ pip install xwpandas
```

## Usage

### Read data from Excel file

```python
import xwpandas as xp
df = xp.read('path/to/file.xlsx')
df
```

### View data in Excel window

```python
xp.view(df)
```

### Save DataFrame to Excel file

```python
xp.save(df, 'path/to/file.xlsx')
```
