#' Écrire une valeur dans une cellule Excel
#'
#' Fonction utilitaire pour écrire une valeur dans une cellule
#' d'une feuille Excel (API Apache POI via le package xlsx).
#'
#' @param sheet feuille Excel (objet sheet)
#' @param row_idx index de la ligne
#' @param col_idx index de la colonne
#' @param value valeur à écrire
#'
#' @return Rien. La cellule est modifiée dans la feuille.
#' @export
set_cell_value <- function(sheet, row_idx, col_idx, value) {

  row <- sheet$getRow(as.integer(row_idx))
  if (is.null(row)) {
    row <- sheet$createRow(as.integer(row_idx))
  }

  cell <- row$getCell(as.integer(col_idx))
  if (is.null(cell)) {
    cell <- row$createCell(as.integer(col_idx))
  }

  cell$setCellValue(value)

}
