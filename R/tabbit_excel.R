#' Write Weighted Crosstab Tables To An Excel Workbook
#'
#' `tabbit_excel()` produces weighted percentage tables and unweighted counts
#' for one or more outcome variables, optionally including missing responses
#' and/or row percentages. Output is written to an Excel file with one sheet per
#' breakdown variable by default (or all results in a single sheet if
#' `by_breakdown = FALSE`).
#'
#' For each outcome variable in `vars` and each breakdown variable in
#' `breakdown`, `tabbit_excel()`:
#'
#' \itemize{
#'   \item computes weighted percentages (by column or by row),
#'   \item optionally adds an overall distribution across breakdowns,
#'   \item optionally adds a "Total percent" row (column mode only),
#'   \item handles missing outcomes either as a separate "Response missing" row
#'         or by excluding them, with or without a separate "Missing percent"
#'         line, and
#'   \item writes a corresponding unweighted N table including (or excluding)
#'         missing responses.
#' }
#'
#' A light formatting layer is applied using the \pkg{openxlsx} package:
#'
#' \itemize{
#'   \item table headers: bold, with top and bottom borders,
#'   \item row labels: bold,
#'   \item total rows ("Total percent" and "Column totals"): bold, with top and
#'         bottom borders, and
#'   \item missing-percentage row: italic.
#' }
#'
#' @param data A data frame.
#' @param vars Character vector of outcome variable names.
#' @param breakdown Character vector of breakdown variables.
#' @param file Path to the Excel file to create.
#' @param wtvar Name of the weight variable (string). Must be present in `data`.
#' @param row_pct Logical. If `FALSE` (default), tables show column
#'   percentages. If `TRUE`, tables show row percentages.
#' @param decimals Integer. Number of decimal places for percentages (0-6;
#'   default 1).
#' @param nooverall Logical. If `TRUE`, suppress the "Overall" column (or
#'   "Overall column percentage (valid responses)" table).
#' @param nototal Logical. If `TRUE`, suppress the "Total percent" row in the
#'   percentage table (column mode only).
#' @param missingasrow Logical. If `TRUE`, include "Response missing" as an
#'   explicit row in the main percentage table.
#' @param nomissing Logical. If `TRUE`, drop missing responses from the
#'   unweighted N table (and from the percentage table unless
#'   `missingasrow = TRUE`).
#' @param by_breakdown Logical. If `TRUE` (default), create one sheet per
#'   breakdown variable. If `FALSE`, stack all results into a single sheet
#'   called `sheet_base`.
#' @param sheet_base Sheet name to use when `by_breakdown = FALSE`.
#' @param ... For future extension.
#'
#' @return Invisibly, the file path of the created workbook (a character string).
#' @importFrom stats xtabs
#'
#' @examples
#' out_file <- tempfile(fileext = ".xlsx")
#'
#' df <- data.frame(
#'   courteous = factor(c("Definitely true","Mostly true", NA, "Mostly false")),
#'   listener  = factor(c("Often","Sometimes","Never", NA)),
#'   sex       = factor(c("Male","Female","Female","Male")),
#'   agegrp1   = factor(c("18-34","35-54","18-34","55+")),
#'   weight = c(1, 1.5, 0.8, 1.2)
#' )
#'
#' tabbit_excel(
#'   data         = df,
#'   vars         = c("courteous", "listener"),
#'   breakdown    = c("sex", "agegrp1"),
#'   file         = out_file,
#'   wtvar        = "weight",
#'   row_pct      = FALSE,
#'   decimals     = 1L,
#'   nooverall    = FALSE,
#'   nototal      = FALSE,
#'   missingasrow = FALSE,
#'   nomissing    = FALSE
#' )
#'
#' @export
tabbit_excel <- function(
    data,
    vars,
    breakdown,
    file,
    wtvar,
    row_pct      = FALSE,
    decimals     = 1L,
    nooverall    = FALSE,
    nototal      = FALSE,
    missingasrow = FALSE,
    nomissing    = FALSE,
    by_breakdown = TRUE,
    sheet_base   = "Frequencies",
    ...
) {
  # Validate Inputs
  if (!is.data.frame(data)) {
    stop("`data` must be a data.frame")
  }

  vars       <- as.character(vars)
  breakdown  <- as.character(breakdown)
  wtvar      <- as.character(wtvar)

  if (!all(vars %in% names(data))) {
    stop("Some `vars` not found in `data`")
  }
  if (!all(breakdown %in% names(data))) {
    stop("Some `breakdown` variables not found in `data`")
  }
  if (!wtvar %in% names(data)) {
    stop("Weight variable `wtvar` not found in `data`")
  }

  # Process Numeric Controls
  decimals <- max(0, min(6, as.integer(decimals)))

  if (missingasrow && nomissing) {
    stop("Options `missingasrow` and `nomissing` cannot both be TRUE.")
  }

  # Create Styles
  header_style <- openxlsx::createStyle(
    textDecoration = "bold",
    border         = "TopBottom"
  )

  rowlabel_style <- openxlsx::createStyle(
    textDecoration = "bold"
  )

  total_style <- openxlsx::createStyle(
    textDecoration = "bold",
    border         = "TopBottom"
  )

  missing_style <- openxlsx::createStyle(
    textDecoration = "italic"
  )

  title_style <- openxlsx::createStyle(
    textDecoration = "bold"
  )

  # Create Workbook
  wb <- openxlsx::createWorkbook()

  # Process Each Breakdown Variable
  for (bvar in breakdown) {

    # Determine Sheet Name
    sheetname <- if (by_breakdown) bvar else sheet_base
    sheetname <- substr(sheetname, 1, 31)

    openxlsx::addWorksheet(wb, sheetname)
    row <- 1

    # Write Sheet Title
    openxlsx::writeData(
      wb, sheetname,
      paste0("Outcome breakdowns by ", bvar),
      startRow = row, startCol = 1
    )
    openxlsx::addStyle(
      wb, sheet = sheetname,
      style = title_style,
      rows  = row,
      cols  = 1,
      gridExpand = TRUE
    )
    row <- row + 1

    # Write Descriptor Line
    openxlsx::writeData(
      wb, sheetname,
      "Weighted %, unweighted N",
      startRow = row, startCol = 1
    )
    row <- row + 1

    # Write Mode Description
    if (!row_pct) {
      openxlsx::writeData(
        wb, sheetname,
        "Weighted column percentages exclude missing responses; unweighted Ns include missing within column totals",
        startRow = row, startCol = 1
      )
    } else {
      openxlsx::writeData(
        wb, sheetname,
        "Weighted row percentages exclude missing responses; unweighted Ns include missing within totals",
        startRow = row, startCol = 1
      )
    }

    # Add Blank Row Before First Variable
    row <- row + 2

    # Process Each Outcome Variable For This Breakdown
    for (v in vars) {
      tab <- .make_weighted_table(
        df           = data,
        outcome      = v,
        breakdown    = bvar,
        wtvar        = wtvar,
        decimals     = decimals,
        row_pct      = row_pct,
        missingasrow = missingasrow,
        nomissing    = nomissing
      )
      if (is.null(tab)) next

      pct         <- tab$pct
      col_base    <- tab$col_base
      row_base    <- tab$row_base
      row_labels  <- tab$row_labels
      b_levels    <- tab$col_labels
      overall_pct <- tab$overall_pct

      v_label <- attr(data[[v]], "label", exact = TRUE)
      if (is.null(v_label) || !nzchar(v_label)) v_label <- v

      # Write Variable Header
      openxlsx::writeData(
        wb, sheetname,
        paste0("Variable: ", v),
        startRow = row, startCol = 1
      )
      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = title_style,
        rows  = row,
        cols  = 1,
        gridExpand = TRUE
      )
      row <- row + 1

      openxlsx::writeData(
        wb, sheetname,
        v_label,
        startRow = row, startCol = 1
      )
      row <- row + 1

      # Build Header Row For Weighted Percentage Table
      if (!row_pct) {
        if (!nooverall) {
          header <- c("Response (valid only)", b_levels, "Overall % (valid)")
        } else {
          header <- c("Response (valid only)", b_levels)
        }
      } else {
        if (!nooverall) {
          header <- c(
            "Response (valid only)",
            b_levels,
            "Total row percentage %",
            "Overall column percentage (valid %)"
          )
        } else {
          header <- c(
            "Response (valid only)",
            b_levels,
            "Total row percentage %"
          )
        }
      }

      header_row <- row

      openxlsx::writeData(
        wb, sheetname,
        t(as.matrix(header)),
        startRow = header_row, startCol = 1,
        colNames = FALSE, rowNames = FALSE
      )

      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = header_style,
        rows  = header_row,
        cols  = seq_along(header),
        gridExpand = TRUE
      )

      row <- header_row + 1

      # Write Weighted And Unweighted Bases
      wbase_row_idx <- row

      openxlsx::writeData(
        wb, sheetname,
        "Weighted base W (valid)",
        startRow = wbase_row_idx, startCol = 1,
        colNames = FALSE, rowNames = FALSE
      )

      wbase_nums <- if (!nooverall) {
        t(c(round(col_base), NA_real_))
      } else {
        t(round(col_base))
      }

      openxlsx::writeData(
        wb, sheetname,
        wbase_nums,
        startRow = wbase_row_idx, startCol = 2,
        colNames = FALSE, rowNames = FALSE
      )

      row <- wbase_row_idx + 1

      Nvalid <- sapply(b_levels, function(lv) {
        sum(!is.na(data[[v]]) & data[[bvar]] == lv, na.rm = TRUE)
      })

      Nvalid_row_idx <- row

      openxlsx::writeData(
        wb, sheetname,
        "Unweighted base N (valid)",
        startRow = Nvalid_row_idx, startCol = 1,
        colNames = FALSE, rowNames = FALSE
      )

      Nvalid_nums <- if (!nooverall) {
        t(c(Nvalid, NA_real_))
      } else {
        t(Nvalid)
      }

      openxlsx::writeData(
        wb, sheetname,
        Nvalid_nums,
        startRow = Nvalid_row_idx, startCol = 2,
        colNames = FALSE, rowNames = FALSE
      )

      row <- Nvalid_row_idx + 1

      # Prepare Weighted Percentage Matrix
      if (!row_pct) {
        if (!nooverall) {
          pct_with_extra <- cbind(pct, overall_pct)
        } else {
          pct_with_extra <- pct
        }
      } else {
        row_totals <- rowSums(pct, na.rm = TRUE)
        row_totals <- round(row_totals, decimals)
        if (!nooverall) {
          pct_with_extra <- cbind(pct, row_totals, overall_pct)
        } else {
          pct_with_extra <- cbind(pct, row_totals)
        }
      }

      pct_start_row <- row
      n_rows_pct    <- nrow(pct_with_extra)

      # Write Row Labels
      openxlsx::writeData(
        wb, sheetname,
        row_labels,
        startRow = pct_start_row, startCol = 1,
        colNames = FALSE, rowNames = FALSE
      )

      # Write Weighted Percentages As Numeric
      openxlsx::writeData(
        wb, sheetname,
        pct_with_extra,
        startRow = pct_start_row, startCol = 2,
        colNames = FALSE, rowNames = FALSE
      )

      # Style Row Labels
      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = rowlabel_style,
        rows  = seq(pct_start_row, length.out = n_rows_pct),
        cols  = 1,
        gridExpand = TRUE
      )

      row <- pct_start_row + n_rows_pct

      # Write Total Percentage Row If Requested
      if (!row_pct && !nototal) {
        col_tot_pct   <- colSums(pct, na.rm = TRUE)
        overall_total <- sum(overall_pct, na.rm = TRUE)

        total_row <- row

        openxlsx::writeData(
          wb, sheetname,
          "Total %",
          startRow = total_row, startCol = 1,
          colNames = FALSE, rowNames = FALSE
        )

        if (!nooverall) {
          total_nums <- t(c(col_tot_pct, overall_total))
        } else {
          total_nums <- t(col_tot_pct)
        }

        openxlsx::writeData(
          wb, sheetname,
          total_nums,
          startRow = total_row, startCol = 2,
          colNames = FALSE, rowNames = FALSE
        )

        openxlsx::addStyle(
          wb, sheet = sheetname,
          style = total_style,
          rows  = total_row,
          cols  = 1:(1 + ncol(total_nums)),
          gridExpand = TRUE
        )

        row <- total_row + 1
      }

      # Write Missing Percentage Row If Requested
      if (!missingasrow && !nomissing) {
        b_all     <- data[[bvar]]
        y_raw_all <- data[[v]]
        y_all <- if (inherits(y_raw_all, "haven_labelled")) {
          haven::as_factor(y_raw_all)
        } else {
          y_raw_all
        }
        w_all <- data[[wtvar]]

        miss_pct <- sapply(b_levels, function(lv) {
          idx <- !is.na(b_all) & b_all == lv
          w_miss <- sum(w_all[idx & is.na(y_all)], na.rm = TRUE)
          w_tot  <- sum(w_all[idx], na.rm = TRUE)
          if (w_tot > 0) round(100 * w_miss / w_tot, decimals) else NA_real_
        })

        idx_any      <- !is.na(b_all)
        w_miss_all   <- sum(w_all[idx_any & is.na(y_all)], na.rm = TRUE)
        w_tot_all    <- sum(w_all[idx_any], na.rm = TRUE)
        overall_miss <- if (w_tot_all > 0) {
          round(100 * w_miss_all / w_tot_all, decimals)
        } else {
          NA_real_
        }

        miss_row_idx <- row

        openxlsx::writeData(
          wb, sheetname,
          "Missing as column percentage of total weighted base (%)",
          startRow = miss_row_idx, startCol = 1,
          colNames = FALSE, rowNames = FALSE
        )

        miss_nums <- if (!nooverall) {
          t(c(miss_pct, overall_miss))
        } else {
          t(miss_pct)
        }

        openxlsx::writeData(
          wb, sheetname,
          miss_nums,
          startRow = miss_row_idx, startCol = 2,
          colNames = FALSE, rowNames = FALSE
        )

        openxlsx::addStyle(
          wb, sheet = sheetname,
          style = missing_style,
          rows  = miss_row_idx,
          cols  = 1:(1 + ncol(miss_nums)),
          gridExpand = TRUE
        )

        row <- miss_row_idx + 2

      } else {
        row <- row + 1
      }

      # Write Unweighted N Table
      title_N <- if (!nomissing) {
        "Unweighted N (responses incl. missing):"
      } else {
        "Unweighted N (responses, missing excluded):"
      }

      openxlsx::writeData(
        wb, sheetname,
        title_N,
        startRow = row, startCol = 1
      )
      row <- row + 1

      b_full <- data[[bvar]]
      y_raw2 <- data[[v]]
      y2 <- if (inherits(y_raw2, "haven_labelled")) {
        haven::as_factor(y_raw2)
      } else {
        y_raw2
      }

      if (nomissing) {
        ok_N <- !is.na(b_full) & !is.na(y2)
      } else {
        ok_N <- !is.na(b_full)
      }

      b_sub <- b_full[ok_N]
      y_sub <- y2[ok_N]

      b_char <- as.character(b_sub)
      y_char <- as.character(y_sub)

      if (!nomissing) {
        y_char[is.na(y_sub)] <- "Response missing"
      }

      if (nomissing) {
        resp_all <- as.character(row_labels)
      } else if (missingasrow) {
        resp_all <- as.character(row_labels)
      } else {
        resp_all <- c(as.character(row_labels), "Response missing")
      }

      b_levels_char <- as.character(b_levels)

      Nmat <- matrix(
        0L,
        nrow = length(resp_all),
        ncol = length(b_levels_char),
        dimnames = list(resp_all, b_levels_char)
      )

      for (j in seq_along(b_levels_char)) {
        lvl  <- b_levels_char[j]
        idxb <- b_char == lvl
        for (i in seq_along(resp_all)) {
          rlab <- resp_all[i]
          Nmat[i, j] <- sum(idxb & (y_char == rlab), na.rm = TRUE)
        }
      }

      RowTotN <- rowSums(Nmat)
      ColTotN <- colSums(Nmat)
      GrandN  <- sum(ColTotN)

      header_vec <- c("Response", b_levels_char, "Total N")
      header_row_N <- row

      openxlsx::writeData(
        wb, sheetname,
        t(as.matrix(header_vec)),
        startRow = header_row_N, startCol = 1,
        colNames = FALSE, rowNames = FALSE
      )

      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = header_style,
        rows  = header_row_N,
        cols  = seq_along(header_vec),
        gridExpand = TRUE
      )

      row <- header_row_N + 1

      first_resp_row <- row

      for (i in seq_along(resp_all)) {
        resp      <- resp_all[i]
        counts    <- as.numeric(Nmat[i, ])
        total_n   <- RowTotN[i]

        row_df <- data.frame(
          Response = resp,
          t(counts),
          `Total N` = total_n,
          check.names = FALSE
        )
        colnames(row_df) <- c("Response", b_levels_char, "Total N")

        openxlsx::writeData(
          wb, sheetname,
          row_df,
          startRow = row, startCol = 1,
          colNames = FALSE, rowNames = FALSE
        )
        row <- row + 1
      }

      last_resp_row <- row - 1

      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = rowlabel_style,
        rows  = seq(first_resp_row, last_resp_row),
        cols  = 1,
        gridExpand = TRUE
      )

      totals_row <- row

      openxlsx::writeData(
        wb, sheetname,
        "Column totals",
        startRow = totals_row, startCol = 1
      )

      col_tot_vec <- c(as.numeric(ColTotN), GrandN)

      openxlsx::writeData(
        wb, sheetname,
        t(as.matrix(col_tot_vec)),
        startRow = totals_row, startCol = 2,
        colNames = FALSE, rowNames = FALSE
      )

      n_tot_cols <- 1 + length(col_tot_vec)

      openxlsx::addStyle(
        wb, sheet = sheetname,
        style = total_style,
        rows  = totals_row,
        cols  = 1:n_tot_cols,
        gridExpand = TRUE
      )

      row <- totals_row + 2

      openxlsx::writeData(
        wb, sheetname,
        "==============================================================",
        startRow = row, startCol = 1
      )
      row <- row + 2
    }
  }


  # Save Workbook
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  invisible(file)
}

#' Internal Helper To Compute Weighted Tables
#' @noRd
.make_weighted_table <- function(df, outcome, breakdown, wtvar,
                                 decimals = 1, row_pct = FALSE,
                                 missingasrow = FALSE, nomissing = FALSE) {
  # Extract Variables
  y_raw <- df[[outcome]]
  b_raw <- df[[breakdown]]
  w     <- df[[wtvar]]

  # Convert Haven Labelled Variables If Needed
  if (inherits(y_raw, "haven_labelled")) {
    y_lab <- haven::as_factor(y_raw)
  } else {
    y_lab <- y_raw
  }

  # Define Valid Cases For Weight And Breakdown
  ok    <- !is.na(b_raw) & !is.na(w)
  y_ok  <- y_lab[ok]
  b_ok  <- b_raw[ok]
  w_ok  <- w[ok]

  # Handle Missing Outcome According To Options
  if (nomissing) {
    valid_idx <- !is.na(y_ok)
    if (!any(valid_idx)) return(NULL)
    y_valid <- y_ok[valid_idx]
    b_valid <- b_ok[valid_idx]
    w_valid <- w_ok[valid_idx]

    y_fac <- factor(y_valid)
    b_fac <- factor(b_valid)

    row_labels <- levels(y_fac)
    col_labels <- levels(b_fac)

  } else if (missingasrow) {
    y_char <- as.character(y_ok)

    nonmiss <- !is.na(y_ok)
    if (any(nonmiss)) {
      base_levels <- levels(factor(y_ok[nonmiss]))
    } else {
      base_levels <- character(0)
    }
    row_labels <- c(base_levels, "Response missing")

    y_char[is.na(y_ok)] <- "Response missing"
    y_fac <- factor(y_char, levels = row_labels)
    b_fac <- factor(b_ok)
    col_labels <- levels(b_fac)

    w_valid <- w_ok

  } else {
    valid_idx <- !is.na(y_ok)
    if (!any(valid_idx)) return(NULL)
    y_valid <- y_ok[valid_idx]
    b_valid <- b_ok[valid_idx]
    w_valid <- w_ok[valid_idx]

    y_fac <- factor(y_valid)
    b_fac <- factor(b_valid)

    row_labels <- levels(y_fac)
    col_labels <- levels(b_fac)
  }

  # Build Weighted Matrix
  Mw <- xtabs(w_valid ~ y_fac + b_fac)
  Mw <- as.matrix(Mw)

  col_base <- colSums(Mw)
  row_base <- rowSums(Mw)

  # Compute Percentages
  if (!row_pct) {
    pct <- sweep(Mw, 2, col_base,
                 function(x, s) ifelse(s > 0, 100 * x / s, NA_real_))
  } else {
    pct <- sweep(Mw, 1, row_base,
                 function(x, s) ifelse(s > 0, 100 * x / s, NA_real_))
  }
  pct <- round(pct, decimals)

  # Compute Overall Percentages
  overall_pct <- 100 * row_base / sum(Mw)
  overall_pct <- round(overall_pct, decimals)

  list(
    pct         = pct,
    col_base    = col_base,
    row_base    = row_base,
    row_labels  = row_labels,
    col_labels  = col_labels,
    overall_pct = overall_pct
  )
}

