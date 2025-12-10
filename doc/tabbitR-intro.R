## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>",
  fig.width = 7,
  fig.height = 5
)

## -----------------------------------------------------------------------------
library(tabbitR)

df <- data.frame(
  outcome = factor(c("A", "B", "A", NA, "C", NA)),
  sex     = factor(c("Male", "Male", "Female", "Female",
                     "Prefer not to say", "Male")),
  weight  = c(1, 2, 1, 1, 0.75, 3)
)

tmp <- tempfile(fileext = ".xlsx")

tabbit_excel(
  data        = df,
  vars        = "outcome",
  breakdown   = "sex",
  wtvar       = "weight",
  file        = tmp,
  decimals    = 1
)

tmp

## -----------------------------------------------------------------------------
### Example toy survey data
set.seed(123)

survey_df <- data.frame(
  outcome1       = factor(sample(c("Agree", "Neutral", "Disagree"), 200, replace = TRUE)),
  outcome2       = factor(sample(c("Often", "Sometimes", "Never"), 200, replace = TRUE)),
  outcome3       = factor(sample(c("Yes", "No"), 200, replace = TRUE)),
  sex            = factor(sample(c("Male", "Female"), 200, replace = TRUE)),
  age            = factor(sample(c("18-34", "35-54", "55+"), 200, replace = TRUE)),
  region         = factor(sample(c("North", "Midlands", "South"), 200, replace = TRUE)),
  survey_weight  = runif(200, 0.5, 2)
)

vars   <- c("outcome1", "outcome2", "outcome3")
breaks <- c("sex", "age", "region")

tmp2 <- tempfile(fileext = ".xlsx")

tabbit_excel(
  data        = survey_df,
  vars        = vars,
  breakdown   = breaks,
  wtvar       = "survey_weight",
  file        = tmp2,
  by_breakdown = TRUE,
  decimals    = 1
)

tmp2


## -----------------------------------------------------------------------------
sessionInfo()

