params <- list(
 #set start and end points
  #CLS, Proposal, TLS, FinRec, BoE, Cou, Adopted
  start_phase = "Adopted",
  start_yr = 23,
  end_phase = "CLS",
  end_yr = 24,
  fy = 24,
  # most up-to-date line item and position files for planning year
  # verify with William for most current version
  line.start = "G:/Fiscal Years/Fiscal 2023/Projections Year/1. July 1 Prepwork/Appropriation File/Fiscal 2023 Appropriation File_Change_Tables.xlsx",
  line.end = "G:/Fiscal Years/Fiscal 2024/Planning Year/1. CLS/1. Line Item Reports/line_items_2022-11-2_CLS FINAL AFTER BPFS.xlsx",
  position.start = "G:/Fiscal Years/Fiscal 2023/Projections Year/1. July 1 Prepwork/Positions/Fiscal 2023 Appropriation File_Change_Tables.xlsx",
  position.end = "G:/Fiscal Years/Fiscal 2024/Planning Year/1. CLS/2. Position Reports/PositionsSalariesOpcs_2022-11-2b_CLS FINAL AFTER BPFS.xlsx",
  # leave revenue file blank if not yet available; script will then just pull in last FY's data
  revenue = "G:/BBMR - Revenue Team/1. Fiscal Years/Fiscal 2023/Planning Year/Budget Publication/FY 2023 - Budget Publication - Prelim.xlsx"
)

##dynamic col names======
# generate_cols <- function() {
#   ordered_cols <- list("CLS" = 1, "Proposal" = 2, "TLS" = 3, "FinRec" = 4, "BoE" = 5, "COU" = 6)
#   x = list()
#   #if phases are part of the same fiscal y ear
#   if (params$start_yr == params$end_yr) {
#     start_index = ordered_cols[[params$start_phase]]
#     end_index = ordered_cols[[params$end_phase]]
#     for (i in start_index:end_index) {
#       j = paste0("FY", params$fy, " ", names(ordered_cols)[i])
#       x = append(x, j)
#     }
#   }
#   else if (params$start_yr != params$end_yr) {
#     print("Still working on it.")
#   }
#   return(x)
# }
# 
# cols <- c(paste0("FY", params$start_yr, " ", params$start_phase), paste0("FY", params$end_yr, " ", params$end_phase))
# n = length(cols)


##libraries ==============================
# .libPaths("C:/Users/sara.brumfield2/Anaconda3/envs/bbmr/Lib/R/library")
.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(tidyverse)
library(magrittr)
library(rio)
library(assertthat)
library(httr)
library(jsonlite)
library(openxlsx)
library(readxl)
library(random)
library(stringr)
library(stringi)

devtools::load_all("G:/Analyst Folders/Sara Brumfield/_packages/bbmR")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/change_table.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/export.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/import.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/osos.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/positions.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/scorecard.R")
source("G:/Budget Publications/automation/0_data_prep/bookDataPrep/R/setup.R")

##formatting ===============================
# Getting service list from appropriations file ensures that defunct services are ignored
services <- import(get_last_mod(paste0("G:/Fiscal Years/Fiscal 20", params$fy - 1,
                                       "/Projections Year/1. July 1 Prepwork/Appropriation File"), "^[^~]*Appropriation.*.xlsx")) %>%
  set_colnames(rename_cols(.)) %>%
  # remove services without one-pagers
  distinct(`Agency Name`, `Service ID`) %>%
  extract2("Service ID")
services <- c(services, "833", "619", "168") # add Innovation Fund

# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0;(#,##0)")

