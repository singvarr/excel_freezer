# excel_freezer

A python app that copies excel table and replaces formulas by calculated values.

## Motivation

### Problem

Imagine that you have multiple reports submitted to external customers and they rely on data calculated by formulas from internal workbooks. It allows to automate data generation and avoid errors. However, it cannot be submitted without replacement this formulas by values because your customers have no access to your workbooks.
This app replaces formulas by values and preserves styles and dimensions of cells from original tables.

## Installation

1. Install dependencies from [requirements.txt](./requirements.txt).
2. Create `.env` file. See example [here](./env-example).
3. Add configuration for table (see below).
4. Launch the app with `python -m excel_freezer.main`.

### Configuring the application

In order to customize freezing process a configuration should be created. It's possible to preserve formulae in the specified ranges or exclude sheets with auxiliary data. Here is an example:

```json
[
  {
    "name": "Sheet1",
    "preserve_sheet": true,
    "cells_with_preserved_formulae": ["G1:I2"]
  },
  {
    "name": "Sheet2",
    "preserve_sheet": false,
    "cells_with_preserved_formulae": []
  }
]
```

## Roadmap of changes

### Functional

- [ ] Refresh table calculation by script
- [x] Remove external links in script
- [x] Add ability to remove specific worksheets on build
- [x] Add ability to preserve formulas in specified ranges and cells

### Non-functional

- [x] Add opportunity to specify table path and output folder via CLI args
- [ ] Add mypy checks
- [ ] Add linting and formatting
