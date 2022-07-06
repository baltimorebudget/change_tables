params <- list(
 #set start and end points
  start_phase = "CLS",
  start_yr = 23,
  end_phase = "COU",
  end_yr = 23,
  fy = 23,
  # most up-to-date line item and position files for planning year
  line.item = "G:/Fiscal Years/Fiscal 2023/Planning Year/7. Council/1. Line Item Reports/line_items_2022-06-24_Final.xlsx",
  position.start = "G:/Fiscal Years/Fiscal 2023/Planning Year/1. CLS/2. Position Reports/PositionsSalariesOpcs_2021-11_03.xlsx",
  position.end = "G:/Fiscal Years/Fiscal 2023/Planning Year/7. Council/3. Position Reports/PositionsSalariesOpcs_2022-06-23.xlsx",
  # leave revenue file blank if not yet available; script will then just pull in last FY's data
  revenue = "G:/BBMR - Revenue Team/1. Fiscal Years/Fiscal 2023/Planning Year/Budget Publication/FY 2023 - Budget Publication - Prelim.xlsx"
)

generate_cols <- function() {
  ordered_cols <- list("CLS" = 1, "Proposal" = 2, "TLS" = 3, "FinRec" = 4, "BoE" = 5, "COU" = 6)
  x = list()
  #if phases are part of the same fiscal y ear
  if (params$start_yr == params$end_yr) {
    start_index = ordered_cols[[params$start_phase]]
    end_index = ordered_cols[[params$end_phase]]
    for (i in start_index:end_index) {
      j = paste0("FY", params$fy, " ", names(ordered_cols)[i])
      x = append(x, j)
    }
  }
  else if (params$start_yr != params$end_yr) {
    print("Still working on it.")
  }
  return(x)
}

cols <- c(paste0("FY", params$start_yr, " ", params$start_phase), paste0("FY", params$end_yr, " ", params$end_phase))
n = length(cols)

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

devtools::load_all("G:/Analyst Folders/Sara Brumfield/bbmR")
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
                                       "/Projections Year/"), "^[^~]*Appropriation.*.xlsx")) %>%
  set_colnames(rename_cols(.)) %>%
  # remove services without one-pagers
  distinct(`Agency Name`, `Service ID`) %>%
  extract2("Service ID")
services <- c(services, "833", "619", "168") # add Innovation Fund

# set number formatting for openxlsx
options("openxlsx.numFmt" = "#,##0")

# check_name_changes("service")

internal <- list(
  output_file_start = paste0("outputs/fy", params$fy,
                             "_", tolower(params$phase), "/"),
  std_oso_salary = c("101", "103", "161", "162", "182"),
  std_oso_opcs = c("201", "231", "233"),
  std_oso = c("000", "101", "103", "161", "162", "182",
              "201", "202", "203", "205", "207", "210", "212", "213", 
              "231", "233", "235", "242", "268", 
              "272", "273", "274", "276", "277", "282", "285", "287", 
              "331", "335", "396", "401", "740")
  std_changes = c("employee compensation and benefits","contractual services expenses",
                  "operating supplies, equipment, software, and computer hardware", 
                  "grants, contributions, and subsidies",
                  "all other"),
  new_services = c("833", "619", "168"), # Innovation Fund
  no_one_pagers = c("129",	# Conditional Purchase Agreement Payments
                    "121",	# Contingent Fund
                    "123",  # General Debt Service
                    "122",	# Miscellaneous General Expenses
                    "124",	# TIF Debt Service
                    "355",	# Employees' Retirement Contribution
                    "126",	# Contribution to Self-Insurance Fund
                    "535",	# Convention Center Hotel
                    "351",  # Retirees' Benefits 
                    "878",  # Disabilities Commision (rolled up with Wage Enforcement)
                    "833",  # Innovation Fund, special agency detail page
                    "352")) # BCPSS, special agency detail page

analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(!is.na(Analyst)) %>%
  mutate(`Agency ID` = as.character(`Agency ID`),
         `File Start` = paste0("outputs/", Analyst,
                               "/", `Agency Name - Cleaned`, " "))
##set paths func ==================
set_paths <- function (path_base = paste0(getwd()), 
                   path_data = paste0(getwd(), "/outputs/", params$start_phase, "_", 
                                      params$end_phase, "_", Sys.Date(), "/")) 
{
  paths <- list(data = path_data)
  dir.create(paths$data)
  paths <- as.list(paste0(list.dirs(paste0(path_base), recursive = FALSE), 
                                           "/")) %>% magrittr::set_names(list.dirs(paste0(path_base), recursive = FALSE, 
                                                                                   full.names = FALSE))

  saveRDS(paths, paste0(paths$data, "paths.Rds"))

  return(paths)
}

paths <- set_paths()



