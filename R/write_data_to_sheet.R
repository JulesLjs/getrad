#' Écrire un data frame dans une feuille Excel cellule par cellule
#'
#' Les valeurs manquantes et les chaînes vides donnent des cellules vides
#' dans les en-têtes et les données.
#'
#' @param sheet Feuille Excel (objet Apache POI).
#' @param data Data frame à écrire.
#' @param start_row Ligne de départ, à partir de 1.
#' @param start_col Colonne de départ, à partir de 1.
#' @param colnames Booléen indiquant si les noms de colonnes doivent être écrits.
#'
#' @return NULL, invisiblement. La fonction modifie directement la feuille Excel.
#' @export
write_data_to_sheet <- function(
    sheet,
    data,
    start_row,
    start_col,
    colnames = TRUE
) {
  # Écriture éventuelle des noms de colonnes
  if (colnames) {
    header_index <- as.integer(start_row - 1)
    header_row <- sheet$getRow(header_index)

    if (is.null(header_row)) {
      header_row <- sheet$createRow(header_index)
    }

    headers <- base::colnames(data)

    for (j in seq_len(ncol(data))) {
      col_index <- as.integer(start_col + j - 2)
      cell <- header_row$getCell(col_index)

      if (is.null(cell)) {
        cell <- header_row$createCell(col_index)
      }

      header <- headers[j]

      if (is.na(header) || header == "") {
        cell$setCellType(3L)
      } else {
        cell$setCellValue(header)
      }
    }
  }

  # Écriture des données
  for (i in seq_len(nrow(data))) {
    row_index <- as.integer(start_row + i - 2 + as.integer(colnames))
    row <- sheet$getRow(row_index)

    if (is.null(row)) {
      row <- sheet$createRow(row_index)
    }

    for (j in seq_len(ncol(data))) {
      col_index <- as.integer(start_col + j - 2)
      cell <- row$getCell(col_index)

      if (is.null(cell)) {
        cell <- row$createCell(col_index)
      }

      value <- data[[j]][i]

      if (is.na(value) || (is.character(value) && value == "")) {
        cell$setCellType(3L)
      } else if (is.numeric(value)) {
        cell$setCellValue(as.numeric(value))
      } else {
        cell$setCellValue(as.character(value))
      }
    }
  }

  invisible(NULL)
}
