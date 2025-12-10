
# tabbitR

<img src="man/figures/logo.png" align="right" height="60" />

Weighted crosstabulation tables exported to Excel.

<!-- badges: start -->

[![R-CMD-check](https://github.com/smmcandrew/tabbitR/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/smmcandrew/tabbitR/actions/workflows/R-CMD-check.yaml)
[![License:
MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.md)
<!-- badges: end -->

## Overview

**tabbitR** automates the production of large sets of weighted
crosstabulation tables and exports them directly to Excel. It is
designed for situations where analysts need many tables at once:

- multiple outcome variables
- multiple breakdown (explanatory) variables
- weighted percentages
- unweighted counts
- transparent treatment of missing values.

This is a common and time-consuming task in survey research, monitoring
and evaluation work, and exploratory data analysis. Doing it manually is
slow, repetitive, and error-prone.

**tabbitR** solves this by generating a full and consistent set of
tables in a single command. For each outcome × breakdown pair, it
writes:

- a weighted percentage table
- a matching unweighted N table
- clearly labelled reporting of missing values
- light formatting (headers, borders, bold labels) using {openxlsx}.

User options control:

- number of decimal places
- row vs.column percentages
- whether to show overall % columns or total rows
- how missing responses should be reported
- whether to create one sheet per breakdown variable.

## Installation

You can install the development version of **tabbitR** from GitHub with
**pak**:

``` r
install.packages("pak")
pak::pak("smmcandrew/tabbitR")
```

or with **remotes**:

``` r
remotes::install_github("smmcandrew/tabbitR")
```

## Example

Below is a minimal example using toy data.

``` r
library(tabbitR)

df <- data.frame(
  outcome = factor(c("A", "B", "A", NA, "C", NA)),
  sex     = factor(c("Male", "Male", "Female", "Female", "Prefer not to say", "Male")),
  weight  = c(1, 2, 1, 1, 0.75, 3)
)

tabbit_excel(
  data        = df,
  vars        = "outcome",
  breakdown   = "sex",
  wtvar       = "weight",
  file        = "example_output.xlsx",
  decimals    = 2
)
```

This will create an Excel workbook with:

- a weighted percentage table, and
- an unweighted N table

… each automatically formatted and labelled.

## Development notes

This README is generated from README.Rmd. Use:

``` r
devtools::build_readme()
```
