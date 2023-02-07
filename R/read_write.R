#' Read a file
#'
#' Read a dataset file.
#' @param file A path to a file.
#' @param skip Number of lines to skip before reading data.
#' @param na Character vector of strings to interpret as missing values.
#' @export
#' @importFrom readr read_csv
#' @importFrom readr read_tsv
#' @importFrom readr read_rds
#' @importFrom readxl read_excel
#' @importFrom arrow read_parquet
#' @importFrom arrow read_feather
#' @importFrom fst read_fst
#' @importFrom haven read_spss
#' @importFrom haven read_stata
#' @importFrom haven read_sas

read <- function (file, skip = 0, na = c("", "NA"))
{
  encoding_res <- readr::guess_encoding(file)[[1]][1] |> as.character()

  if (grepl("(?i)\\.csv$", file)) {
    file <-
      readr::read_csv(
        file,
        locale = readr::locale(encoding = encoding_res),
        skip = skip,
        na = na
      )
  }
  else if (grepl("(?i)\\.tsv$", file)) {
    file <-
      readr::read_tsv(
        file,
        locale = readr::locale(encoding = encoding_res),
        skip = skip,
        na = na
      )
  }
  else if (grepl("(?i)((\\.xlsx$)|(\\.xls$))", file)) {
    file <- readxl::read_excel(file, skip = skip, na = na)
  }
  else if (grepl("(?i)\\.parquet$", file)) {
    file <- arrow::read_parquet(file)
  }
  else if (grepl("(?i)\\.feather$", file)) {
    file <- arrow::read_feather(file)
  }
  else if (grepl("(?i)\\.rds$", file)) {
    file <- readr::read_rds(file)
  }
  else if (grepl("(?i)\\.fst$", file)) {
    file <- fst::read_fst(file)
  }
  else if (grepl("(?i)((\\.sav$)|(\\.por$))", file)) {
    file <- haven::read_spss(file, skip = skip)
  }
  else if (grepl("(?i)\\.dta$", file)) {
    file <- haven::read_stata(file, skip = skip)
  }
  else if (grepl("(?i)((\\.sas7bdat$)|(\\.sas7bcat$))", file)) {
    file <- haven::read_sas(file, skip = skip)
  }
  else if (!file.exists(file)) {
    stop("No such file")
  }
  return(file)
}


#' Write a file
#'
#' Write a dataset file.
#' @param data A data frame or tibble to write to disk.
#' @param file File to write to.
#' @export
#' @importFrom readr write_csv
#' @importFrom readr write_rds
#' @importFrom arrow write_parquet
#' @importFrom arrow write_feather
#' @importFrom fst write_fst
#' @importFrom haven write_sav
#' @importFrom haven write_dta
#' @importFrom haven write_sas


write <- function (data, file)
{
  if (grepl("(?i)\\.csv$", file)) {
    readr::write_excel_csv(data, file)
  }
  else if (grepl("(?i)\\.parquet$", file)) {
    arrow::write_parquet(data, file)
  }
  else if (grepl("(?i)\\.feather$", file)) {
    arrow::write_feather(data, file)
  }
  else if (grepl("(?i)\\.rds$", file)) {
    readr::write_rds(data, file)
  }
  else if (grepl("(?i)\\.fst$", file)) {
    fst::write_fst(data, file)
  }
  else if (grepl("(?i)((\\.sav$)|(\\.por$))", file)) {
    haven::write_sav(data, file)
  }
  else if (grepl("(?i)\\.dta$", file)) {
    names(data) <- gsub("\\.", "_", names(data))
    haven::write_dta(data, file)
  }
  else if (grepl("(?i)((\\.sas7bdat$)|(\\.sas7bcat$))", file)) {
    names(data) <- gsub("\\.", "_", names(data))
    haven::write_sas(data, file)
  }

}
