#' @importFrom dplyr select
#' @importFrom dplyr mutate
#' @importFrom dplyr rename
#' @importFrom dplyr tibble
#' @importFrom stringr str_c
#' @importFrom stringr str_remove_all
#' @importFrom stringr str_replace_all
#' @importFrom stringr str_glue
#' @importFrom lubridate now
#' @importFrom psych describe
#' @import openxlsx

## describe function

describe <- function(data) {
  desc <- data |>
    dplyr::select(where(is.numeric)) |>
    psych::describe()

  desc$rownames <- base::rownames(desc)

  maxChar <- desc$rownames |>
    base::nchar() |>
    base::max()

  desc <- desc |>
    dplyr::tibble() |>
    dplyr::select(
      rownames, n, mean, sd, min, max
    ) |>
    dplyr::mutate(
      min = dplyr::if_else(
        min %% 1 != 0,
        base::round(min, digits = 4) |>
          base::format(nsmall = 4) |>
          stringr::str_c() |>
          stringr::str_remove_all(" "),
        stringr::str_c(min)
      ),
      max = dplyr::if_else(
        max %% 1 != 0,
        base::round(max, digits = 4) |>
          base::format(nsmall = 4) |>
          stringr::str_c() |>
          stringr::str_remove_all(" "),
        stringr::str_c(max)
      )
    ) |>
    dplyr::rename(
      "\u5909\u6570\u540d" = rownames,
      "\u89b3\u6e2c\u6570" = n,
      "\u5e73\u5747\u5024" = mean,
      "\u6a19\u6e96\u504f\u5dee" = sd,
      "\u6700\u5c0f\u5024" = min,
      "\u6700\u5927\u5024" = max
    )


  nrow <- desc |> base::nrow()


  workBook <- openxlsx::createWorkbook()
  workSheet <- openxlsx::addWorksheet(wb = workBook,
                                      sheetName = "\u8a18\u8ff0\u7d71\u8a08",
                                      tabColour = "blue")

  openxlsx::writeData(
    wb = workBook,
    sheet = 1,
    x = desc,
    xy = c(2, 4),
    borders = "surrounding",
    borderColour = "black",
    borderStyle = "thin"
  )

  openxlsx::deleteData(
    wb = workBook,
    sheet = 1,
    cols = 2,
    rows = 4
  )

  styleMarginTop <- openxlsx::createStyle(
    fontName = "Century",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "center",
    valign = "center",
    numFmt = "TEXT",
    border = "bottom",
    borderColour = "black",
    borderStyle = "double"
  )

  styleMarginBottom <- openxlsx::createStyle(
    fontName = "Century",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "center",
    valign = "center",
    numFmt = "TEXT",
    border = "top",
    borderColour = "black",
    borderStyle = "double"
  )

  styleHeading <- openxlsx::createStyle(
    fontName = "MS Mincho",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "center",
    valign = "center",
    numFmt = "TEXT",
    border = "bottom",
    borderColour = "black",
    borderStyle = "thin"
  )

  styleValue <- openxlsx::createStyle(
    fontName = "century",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "right",
    valign = "center",
    numFmt = "0.0000"
  )

  styleN <- openxlsx::createStyle(
    fontName = "century",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "right",
    valign = "center",
    numFmt = "0"
  )

  styleNameMS <- openxlsx::createStyle(
    fontName = "MS Mincho",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "left",
    valign = "center",
    numFmt = "TEXT"
  )

  styleNameCentury <- openxlsx::createStyle(
    fontName = "Century",
    fontSize = 11,
    fontColour = "black",
    fgFill = "white",
    halign = "left",
    valign = "center",
    numFmt = "TEXT"
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleMarginTop,
    rows = 3,
    cols = 2:7,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleMarginBottom,
    rows = (nrow + 5),
    cols = 2:7,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleHeading,
    rows = 4,
    cols = 2:7,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleValue,
    rows = 5:(nrow + 4),
    cols = 4:7,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleN,
    rows = 5:(nrow + 4),
    cols = 3,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleNameMS,
    rows = 5:(nrow + 4),
    cols = 2,
    gridExpand = TRUE
  )

  openxlsx::addStyle(
    wb = workBook,
    sheet = 1,
    style = styleNameCentury,
    rows = 5:(nrow + 4),
    cols = 2,
    gridExpand = TRUE
  )

  openxlsx::setColWidths(
    wb = workBook,
    sheet = 1,
    cols = 3:7,
    widths = 12
  )

  openxlsx::setColWidths(
    wb = workBook,
    sheet = 1,
    cols = 2,
    widths = (maxChar * 2.25) |>
      base::round()
  )

  openxlsx::setRowHeights(
    wb = workBook,
    sheet = 1,
    rows = 4:(nrow + 4),
    heights = 18
  )

  openxlsx::setRowHeights(
    wb = workBook,
    sheet = 1,
    rows = c(3, nrow + 5),
    heights = 4.5
  )

  fileName <-
    stringr::str_glue("\u8a18\u8ff0\u7d71\u8a08_{lubridate::now()}.xlsx") |>
    stringr::str_replace_all(pattern = "(-|:)",
                             replacement = "") |>
    stringr::str_replace_all(pattern = " ",
                             replacement = "_")

  openxlsx::saveWorkbook(wb = workBook,
                         file = fileName,
                         overwrite = TRUE)

}

