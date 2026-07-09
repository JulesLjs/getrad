#' Écrire une dataframe dans une feuille Excel cellule par cellule
#'
#' Fonction utilitaire pour écrire une dataframe dans une feuille Excel
#' via l'API Java Apache POI (via le package xlsx).
#'
#' @param sheet feuille Excel (objet sheet Apache POI)
#' @param data dataframe à écrire
#' @param start_row ligne de départ
#' @param start_col colonne de départ
#'
#' @return Rien. La fonction modifie directement la feuille Excel.
#' @export
write_data_to_sheet <- function(sheet, data, start_row, start_col) {

  # Écriture des noms de colonnes
  header_row <- sheet$getRow(as.integer(start_row - 1))
  if (is.null(header_row)) {
    header_row <- sheet$createRow(as.integer(start_row - 1))
  }

  for (j in seq_len(ncol(data))) {
    col_index <- as.integer(start_col + j - 2)

    cell <- header_row$getCell(col_index)
    if (is.null(cell)) {
      cell <- header_row$createCell(col_index)
    }

    cell$setCellValue(colnames(data)[j])
  }

  # Écriture des données
  for (i in seq_len(nrow(data))) {

    row_index <- as.integer(start_row + i - 1)

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

      value <- data[i, j]

      if (is.numeric(value)) {
        cell$setCellValue(as.numeric(value))
      } else {
        cell$setCellValue(as.character(value))
      }
    }
  }

}
