#' Écrire une valeur dans une cellule Excel
#'
#' Fonction utilitaire pour écrire une valeur dans une cellule
#' d'une feuille Excel (API Apache POI via le package xlsx).
#'
#' @param sheet Feuille Excel (objet Apache POI)
#' @param row_idx Numéro de ligne
#' @param col_idx Numéro de colonne
#' @param value Valeur à écrire
#'
#' @return NULL, invisiblement. La cellule est modifiée dans la feuille.
#' @export
set_cell_value <- function(sheet, row_idx, col_idx, value) {
  row_index <- as.integer(row_idx - 1L)
  col_index <- as.integer(col_idx - 1L)

  row <- sheet$getRow(row_index)
  if (is.null(row)) {
    row <- sheet$createRow(row_index)
  }

  cell <- row$getCell(col_index)
  if (is.null(cell)) {
    cell <- row$createCell(col_index)
  }

  cell$setCellValue(value)
  invisible(NULL)
}
