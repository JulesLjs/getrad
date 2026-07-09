#' Convertir un fichier Excel `.xlsx` en fichier OpenDocument `.ods`
#'
#' Cette fonction convertit un fichier Excel au format `.xlsx` en fichier
#' OpenDocument Spreadsheet (`.ods`) en utilisant l'exécutable `soffice.com`
#' de LibreOffice.
#'
#' Le fichier d'entrée est recherché dans `path_input` et le fichier converti
#' est enregistré dans `path_output`. Si `file_output` n'est pas renseigné,
#' le fichier de sortie reprend automatiquement le nom du fichier d'entrée
#' sans l'extension `.xlsx`.
#'
#' La fonction :
#' \itemize{
#'   \item vérifie l'existence de `soffice.com` ;
#'   \item vérifie l'existence du fichier `.xlsx` d'entrée ;
#'   \item crée le dossier de sortie s'il n'existe pas ;
#'   \item lance la conversion via LibreOffice ;
#'   \item renomme le fichier généré si nécessaire.
#' }
#'
#' @param path_input Chaîne de caractères. Chemin du dossier contenant le
#'   fichier Excel d'entrée.
#' @param file_input Chaîne de caractères. Nom du fichier Excel d'entrée,
#'   avec ou sans l'extension `.xlsx`.
#' @param path_output Chaîne de caractères. Chemin du dossier de sortie.
#'   Par défaut, identique à `path_input`.
#' @param file_output Chaîne de caractères ou `NULL`. Nom du fichier de
#'   sortie, sans extension. Si `NULL`, le nom de `file_input` est utilisé.
#' @param path_to_soffice Chaîne de caractères. Chemin du dossier contenant
#'   l'exécutable `soffice.com` de LibreOffice.
#'
#' @return Invisiblement, le chemin complet du fichier `.ods` généré.
#'
#' @examples
#' \dontrun{
#' convert_xlsx_to_ods(
#'   path_input = "C:/mes_fichiers",
#'   file_input = "tableau.xlsx"
#' )
#'
#' convert_xlsx_to_ods(
#'   path_input = "C:/mes_fichiers",
#'   file_input = "tableau",
#'   path_output = "C:/sortie",
#'   file_output = "tableau_converti"
#' )
#' }
#'
#' @export
convert_xlsx_to_ods <- function(path_input,
                                file_input,
                                path_output = path_input,
                                file_output = NULL,
                                path_to_soffice = "C:/Program Files/LibreOffice/program/") {

  file_input <- sub("\\.xlsx$", "", file_input, ignore.case = TRUE)

  if (is.null(file_output)) {
    file_output <- file_input
  }

  soffice_path <- file.path(path_to_soffice, "soffice.com")
  input_file <- file.path(path_input, paste0(file_input, ".xlsx"))
  output_dir <- path_output
  output_file <- file.path(output_dir, paste0(file_output, ".ods"))
  generated_file <- file.path(output_dir, paste0(file_input, ".ods"))

  if (!file.exists(soffice_path)) {
    stop("soffice.com introuvable : ", soffice_path, call. = FALSE)
  }

  if (!file.exists(input_file)) {
    stop("Fichier d'entrée introuvable : ", input_file, call. = FALSE)
  }

  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  command_conversion <- paste0(
    '"', soffice_path, '" ',
    '--convert-to "ods:calc8" ',
    '--outdir "', output_dir, '" ',
    '"', input_file, '"'
  )

  status <- system(command_conversion)

  if (status != 0) {
    stop("Erreur lors de la conversion LibreOffice.", call. = FALSE)
  }

  if (!file.exists(generated_file)) {
    stop("Le fichier .ods n'a pas été généré.", call. = FALSE)
  }

  if (normalizePath(generated_file, winslash = "\\", mustWork = FALSE) !=
      normalizePath(output_file, winslash = "\\", mustWork = FALSE)) {
    file.rename(generated_file, output_file)
  }

  return(invisible(output_file))
}
