
# Chargement du package et des données
devtools::load_all()
date_revu <- as.Date("2026-06-01")

data_fig_2 <- data.frame(
  "Janvier 26 / Décembre 25 (en %)" = c(0.0, 0.4, -7.3, 0.9), # Nom de colonnes mobiles écrit dans R
  "Février 26 / Janvier 26 (en %)" = c(0.5, -1.9, 0.0, 1.5),
  "Mars 26 / Février 26 (en %)" = c(2.2, -1.4, -2.9, 4.3),
  "Avril 26 / Mars 26 (en %)" = c(-6.3, -3.3, -4.5, -7.6),
  "Avril 26" = c(99569, 25033, 8372, 66164),
  row.names = c(
    "Ensemble",
    "Sociétés",
    "Entreprises individuelles classiques",
    "Micro-entrepreneurs"
  ),
  check.names = FALSE
)

data_fig_3 <- data.frame(
  "sur 3 mois" = c(11.8, 4.7, 3.2, 15.9),
  "sur 12 mois" = c(9.3, 7.3, 3.0, 10.9),
  "2025" = c(393883, 104636, 34808, 254439),
  "2026" = c(438292, 111880, 35697, 290715),
  row.names = c(
    "Ensemble",
    "Sociétés",
    "Entreprises individuelles classiques",
    "Micro-entrepreneurs"
  ),
  check.names = FALSE
)

data_fig_4 <- data.frame(
  "Cumul sur 1 mois (Avril 2026)" = c(100.0, 25.6, 7.8, 66.6),
  "Cumul sur 3 mois (Février 2026 à Avril 2026)" = c(100.0, 25.4, 7.9, 66.7),
  "Cumul sur 12 mois (Mai 2025 à Avril 2026)" = c(100.0, 25.5, 8.8, 65.7),
  "Cumul sur 12 mois, un an auparavant (Mai 2024 à Avril 2025)" = c(100.0, 26.0, 9.4, 64.7),
  row.names = c(
    "Ensemble",
    "Sociétés",
    "Entreprises individuelles classiques",
    "Micro-entrepreneurs"
  ),
  check.names = FALSE
)

data_fig_7 <- data.frame(
  "Février 26" = c(
    103953, 6117, 4311, 7305, 28897, 15358, 9396,
    4143, 7614, 2408, 3435, 27678, 9301, 11198
  ),
  "Mars 26" = c(
    106274, 6120, 4351, 7451, 29139, 15670, 9169,
    4300, 8020, 2500, 3386, 28691, 9477, 11490
  ),
  "Avril 26" = c(
    99569, 5651, 4031, 7481, 27180, 14921, 8110,
    4149, 7784, 2346, 3185, 26750, 8730, 10462
  ),
  row.names = c(
    "CREAT", "BE", "C", "FZ", "GI", "G", "H", "I", "JZ", "KZ", "LZ", "MN", "OQ", "RU"
  ),
  check.names = FALSE
)

data_fig_8 <- data.frame(
  "Avril 26 / Mars 26" = c(
    -6.3, -7.7, -7.4, 0.4, -6.7, -4.8, -11.5,
    -3.5, -2.9, -6.2, -5.9, -6.8, -7.9, -8.9
  ),
  "sur 3 mois" = c(
    11.8, 6.8, 16.0, 8.0, 6.8, 12.7, -2.5,
    7.6, 37.9, -3.4, 5.7, 17.7, 8.9, 9.9
  ),
  "sur 12 mois" = c(
    9.3, 6.3, 14.5, 3.1, 7.9, 14.4, -0.8,
    7.4, 22.1, -9.8, 9.2, 15.0, 6.2, 6.2
  ),
  row.names = c(
    "CREAT", "BE", "C", "FZ", "GI", "G", "H", "I", "JZ", "KZ", "LZ", "MN", "OQ", "RU"
  ),
  check.names = FALSE
)

#               Etape 1 : Chargement du fichier initial (excel)              ---

wb_figure_fr <- xlsx::loadWorkbook("inst/example/tableaux_creations_entreprises_fr.xlsx")

#                     Etape 2 : Extraction des feuilles                      ---

sheets <- xlsx::getSheets(wb_figure_fr)

#     Etape 3 : Ecriture des dataframes dans les feuilles correspondantes    ---

write_data_to_sheet(sheets[[1]], data_fig_2, start_row = 3, start_col = 2)
write_data_to_sheet(sheets[[2]], data_fig_3, start_row = 4, start_col = 2)
write_data_to_sheet(sheets[[3]], data_fig_4, start_row = 3, start_col = 2)

# Personnalisation d'une cellule spécifique
set_cell_value(sheets[[2]], 2, 3,
               if (lubridate::month(date_revu) == 1) {
                 "Cumul sur le 1er mois de l'année"
               } else if (lubridate::month(date_revu) == 12) {
                 "Cumul sur les 12 mois de l'année"
               } else {
                 paste("Cumul sur les", lubridate::month(date_revu), "premiers mois de l'année")
               }
)

write_data_to_sheet(sheets[[4]], data_fig_7, start_row = 3, start_col = 3)
write_data_to_sheet(sheets[[5]], data_fig_8, start_row = 4, start_col = 3)

#                    Etape 4 : Sauvegarde du fichier                         ---

xlsx::saveWorkbook(wb_figure_fr, "inst/example/data_xlsx.xlsx")

#                      Etape 5 : Duplication en ODS                          ---

convert_xlsx_to_ods("inst/example", "data_xlsx",
                    "inst/example", "data_ods")
