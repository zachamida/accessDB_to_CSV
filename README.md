# accdbtools

accdbtools is a Python package that provides utilities for converting accdb tables to CSV files.

## Installation

You can install the package using pip:

```bash
pip install accdbtools

from accdbtools import accdb2csv

# Set the path to your accdb file
accdb_file = "path/to/your/database.accdb"
accdb2csv(accdb_file)
```
## Requirements

Before using the `accdbtools` package, make sure you have the following requirement installed:

- [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-US/download/details.aspx?id=54920)