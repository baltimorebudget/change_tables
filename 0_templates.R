##set to CLS and Prop // will need to change code manually if TLS is incorporated

.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(tidyverse)
library(magrittr)
library(rio)
library(openxlsx)
library(readxl)
library(janitor)

devtools::load_all("G:/Analyst Folders/Sara Brumfield/_packages/bbmR")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/0_import_functions.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/0_latex_functions.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/0_setup_functions.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/1a_agency_summary.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/1b_service_budget_tables.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/1b_service_detail_functions.R")
source("G:/Budget Publications/automation/2_agency_detail/bookAgencyDetail/R/1b_service_one_pagers.R")

params = list(fy = 24)

options("openxlsx.numFmt" = "#,##0")

#data from 0_data_prep ====
analysts <- import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
  filter(!is.na(Analyst) & Operational == TRUE) %>%
  mutate(`Agency ID` = as.character(`Agency ID`)) %>%
  select(-(Operational:Notes))

path <- list(cls = paste0("G:/Budget Publications/automation/0_data_prep/outputs/fy",
                          params$fy, "_cls/"),
             prop = paste0("G:/Budget Publications/automation/0_data_prep/outputs/fy",
                           params$fy, "_prop/"))

cols <- readRDS(paste0(path$prop, "cols.Rds"))
cols$dollars_cls <- paste0("FY", params$fy, " Dollars - CLS")
cols$dollars_prop <- paste0("FY", params$fy, " Dollars - PROP")

line_item <- list(
  cls = readRDS(paste0(path$cls, "expenditure.Rds")),
  prop = readRDS(paste0(path$prop, "expenditure.Rds")))  %>%
  map(group_by_at, vars(starts_with("Agency Name"), starts_with("Service"), starts_with("Fund"), "Objective Name")) %>%
  map(summarize_at, vars(starts_with("FY")), sum, na.rm = TRUE) %>%
  map(ungroup)

line_item$cls <- line_item$cls %>%
  select(-ends_with("Name"), -ends_with("Actual"), -!!paste0("FY", params$fy - 1, " Budget"), `Objective Name`) %>%
  rename(!!cols$dollars_cls := !!paste0("FY", params$fy, " Budget"))

line_item$prop <- line_item$prop %>%
  select(`Agency Name`:`Objective Name`, !!cols$expend$prior, !!cols$expend$projection, 
         !!cols$dollars_prop := !!paste0("FY", params$fy, " Budget")) %>%
  set_colnames(gsub("Budget|Actual", "Dollars", names(.)))

line_item <- line_item$prop %>%
  left_join(line_item$cls) %>%
  relocate(!!cols$dollars_prop, .after = last_col())

positions <- readRDS(paste0(path$prop, "positions.Rds")) 

positions$cls <- readRDS(paste0(path$cls, "positions.Rds")) %>%
  extract2("planning")

positions <- positions %>%
  map(group_by, `Agency Name`, `Service Name`, `Service ID`, `Fund Name`) %>%
  map(count)

positions <- positions$planning %>%
  rename(`FY24 Positions - PROP` = n) %>%
  left_join(positions$cls %>%
              rename(`FY24 Positions - CLS` = n)) %>%
  left_join(positions$projection %>%
              rename(`FY23 Positions` = n)) %>%
  left_join(positions$prior %>%
              rename(`FY22 Positions` = n)) %>%
  rename(`FY22 Actual` = `FY22 Positions`, `FY23 Adopted` = `FY23 Positions`, `FY24 CLS` = `FY24 Positions - CLS`, `FY24 Request` = `FY24 Positions - PROP`) %>%
  mutate(`Objective Name` = "",
         Type = "Positions") %>%
  select(`Agency Name`, `Service Name`, `Service ID`, `Fund Name`, `FY22 Actual`, `FY23 Adopted`, `FY24 CLS`, `FY24 Request`, `Objective Name`, Type) 
  
summary <- line_item %>%
  rename(`FY22 Actual` = `FY22 Dollars`, `FY23 Adopted` = `FY23 Dollars`, `FY24 CLS` = `FY24 Dollars - CLS`, `FY24 Request` = `FY24 Dollars - PROP`) %>%
  select(`Agency Name`, `Service Name`, `Service ID`, `Fund Name`, `FY22 Actual`, `FY23 Adopted`, `FY24 CLS`, `FY24 Request`, `Objective Name`) %>%
  mutate(Type = "Expenditures") %>%
  rbind(positions) %>%
# summary <- line_item %>%
#   left_join(positions) %>%
  # mutate(`Fund` = case_when(`Fund Name` == "General" ~ "General Fund",
  #                           TRUE ~ "All Other Funds")) %>%
  # relocate(Fund, .after = `Fund Name`) %>%
  # select(-`Fund Name`) %>%
  mutate_if(is.numeric, replace_na, 0) %>%
#   select(`Agency Name`:`Fund`, sort(names(.))) %>%
#   select(-`FY24 Prop`) %>%
#   relocate(`FY24 Positions - CLS`, .after = !!cols$dollars_cls) %>%
#   filter(!!sym(cols$dollars_prop) != 0) %>%
  group_by(`Agency Name`, `Service Name`, `Service ID`, Type, `Fund Name`, `Objective Name`) %>%
  summarise_if(is.numeric, sum, na.rm = TRUE)

## to space tables and analysis box
check <- summary %>% group_by(`Agency Name`, `Service ID`) %>% summarise(count = n())

##OSO changes
osos <- readRDS(paste0(path$prop, "expenditure.Rds")) %>%
  select(`Agency Name`:`Service Name`, `Fund ID`:`Fund Name`, `Subobject ID`, `Subobject Name`, `FY24 CLS`:`FY24 Budget`) %>%
  mutate(
    # `Fund` = case_when(`Fund Name` == "General" ~ "General Fund",
    #                         TRUE ~ "All Other Funds"),
         `OSO` = case_when(`Subobject ID` %in% c("101", "161", "162") ~ "Permanent Wages Adjustment",
                           `Subobject ID` %in% c("204", "245", "788", "202", "203", "279") ~ "Pension Adjustment",
                           `Subobject ID` %in% c("207", "85", "285", "368", "766") ~ "Healthcare Adjustment",
                           `Subobject ID` %in% c("357", "335", "401", "331", "341") ~ "Fleet Adjustment",
                           `Subobject ID` %in% c("352", "311", "313", "316", "336", "380") ~ "Building Maintenance and Rental Charge Adjustment",
                           TRUE ~ "Other"),
         Diff = `FY24 Budget` - `FY24 CLS`) %>%
  # relocate(Fund, .after = `Fund Name`) %>%
  # select(-`Fund Name`, -`Fund ID`, -`Subobject ID`, -`Subobject Name`) %>%
  group_by(`Agency Name`, `Service ID`, `Service Name`, `Fund Name`, OSO) %>%
  summarise_if(is.numeric, sum, na.rm = TRUE) %>%
  filter(Diff != 0)

agencies = analysts$`Agency Name`

##agency summary tables from 2_agency_detail ===========

##make and export templates for each service ======

make_template <- function(
    df, oso, svc, tab_name, file_name, type = "new", col_width = 'auto',
    tab_color = NULL, table_name = NULL, show_tab = TRUE, save = TRUE) {
  
  if (missing(df) | missing(tab_name) | missing(file_name)) {
    stop("Missing argument(s)")}
  if (!type %in% c("new", "existing")){
    stop('Please specify if you are exporting a worksheet to a \"new\"
         or \"existing\" Excel file')}
  if (!grepl("\\.xlsx", file_name)){
    stop('Please ensure that the file_name includes the .xlsx extension.')}
  tab_name %<>% as.character
  if (nchar(tab_name) > 31){
    stop("Please shorten the tab name to 31 characters or fewer.")}
  
  Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe") # corrects Rtools in wrong env error
  
  options("openxlsx.orientation" = "landscape",
          "openxlsx.datetimeFormat" = "yyyy-mm-dd",
          "openxlsx.numFmt" = "#,##0;-#,##0;--")
  
  
  agency <- unique(df$`Agency Name`)
  pillar <- unique(levels(factor(df$`Objective Name`)))
  svc_name <- unique(df$`Service Name`)
  svc_id <- unique(df$`Service ID`)
  
  data <- df %>%
    ungroup() %>%
    select(Type, `Fund Name`, `FY22 Actual`:`FY24 Request`) %>%
    mutate(`FY24 TLS` = 0,
           `Change $` = `FY24 Request` - `FY24 CLS`) %>%
    rowwise() %>%
    mutate(`Change %` = ifelse(is.numeric(round(((`FY24 Request` - `FY24 CLS`)/`FY24 CLS`) * 100, 1)), (round(((`FY24 Request` - `FY24 CLS`)/`FY24 CLS`) * 100, 1)), 0)) %>%
    arrange(Type, match(`Fund Name`, c("General", "Internal Service", "Federal", "State", "Special Grant", "Special Revenue", "Parking Management", "Conduit Enterprise", "Stormwater Utility", "Wastewater Utility", "Convention Center Bond", "Loan and Guarantee Enterprise", "Motor Vehicle", NA)))
  
  dollar_data <- data %>%
    filter(Type == "Expenditures") %>%
    select(-Type)

  position_data <- data %>%
    filter(Type == "Positions") %>%
    select(-Type)
  
  tech_adjust <- data.frame(Item = c("Enter item here", "", "", "", "Total"),
                            Object = c(rep("", 5)),
                            Subobject = c(rep("", 5)),
                            Amount = c(rep("", 5)),
                            Decision = c(rep("", 5)))
  
  savings_ideas <- data.frame(Item = c("Enter item here", "", "", "", "Total"),
                              Object = c(rep("", 5)),
                              Subobject = c(rep("", 5)),
                              Amount = c(rep("", 5)),
                              Decision = c(rep("", 5)))
  
  change_table <- data.frame(Adjustments = c("FY2023 Adopted", 
                                                "CLS Adjustments", 
                                                "Permanent Wages Adjustment", 
                                                "Updated Pension costs based on projected rate", 
                                                "Updated healthcare costs based on projected rate", 
                                                "All other benefit updates", 
                                                "Updated Fleet rate", 
                                                "Updated Building/Maintenance/Rental Charges", 
                                                "All other updates",
                                                "Build-ins/Build-outs (please list)", 
                                                "", 
                                                "Request Adjustments", 
                                                "", 
                                                "", 
                                                "TLS Adjustments", 
                                                "", 
                                                "", 
                                                "FinRec Adjustments", 
                                                "", 
                                                "", 
                                                "BoE Adjustments", 
                                                "", 
                                                "", 
                                                "Council Adjustments", 
                                                "", 
                                                "", 
                                                "FY2024 Budget"),
                              Amount = c(sum(data$`FY23 Adopted`[data$Type=="Expenditures" & data$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(data$`FY24 CLS`[data$Type=="Expenditures" & data$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(oso$`Diff`[oso$OSO=="Permanent Wages Adjustment" & oso$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(oso$`Diff`[oso$OSO=="Pension Adjustment" & oso$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(oso$`Diff`[oso$OSO=="Healthcare Adjustment" & oso$`Fund Name` == "General"], na.rm = TRUE), 
                                         "",
                                         sum(oso$`Diff`[oso$OSO=="Fleet Adjustment" & oso$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(oso$`Diff`[oso$OSO=="Building Maintenance and Rental Charge Adjustment" & oso$`Fund Name` == "General"], na.rm = TRUE), 
                                         sum(oso$`Diff`[oso$OSO=="Other" & oso$`Fund Name` == "General"], na.rm = TRUE),
                                         "", 
                                         "",
                                         sum(data$`FY24 Request`[data$Type=="Expenditures" & data$`Fund Name` == "General"], na.rm = TRUE), 
                                         "", 
                                         "", 
                                         sum(data$`FY24 TLS`[data$Type=="Expenditures" & data$`Fund Name` == "General"], na.rm = TRUE), 
                                         "", 
                                         "", 
                                         "Pending", 
                                         "", 
                                         "", 
                                         "Pending", 
                                         "", 
                                         "", 
                                         "Pending", 
                                         "", 
                                         "", 
                                         "Pending"),
                             `Tollgate Decision` = rep("", 27))
  
  excel <- switch(type,
                  "new" = openxlsx::createWorkbook(),
                  "existing" = openxlsx::loadWorkbook(file_name)) %T>%
    openxlsx::addWorksheet(
      tab_name, tabColour = tab_color,
      header = c(gsub('\\..*', '', file_name), "&[Tab]", as.character(Sys.Date())),
      footer = c(NA, "&[Page]", NA), visible = show_tab) %T>%
    ##heading info
    openxlsx::writeData(tab_name, x = "Service Summary", startCol = 1, startRow = 1) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14, textDecoration = "bold", border = "bottom", borderStyle = "thin"), rows = 1, cols = 1) %T>%
    openxlsx::writeData(tab_name, x = "Agency", startCol = 1, startRow = 2) %T>%
    openxlsx::writeData(tab_name, x = "Service Name", startCol = 1, startRow = 3) %T>%
    openxlsx::writeData(tab_name, x = "Service ID", startCol = 1, startRow = 4) %T>%
    openxlsx::writeData(tab_name, x = "Pillar", startCol = 1, startRow = 5) %T>%
    openxlsx::writeData(tab_name, x = agency, startCol = 2, startRow = 2) %T>%
    openxlsx::writeData(tab_name, x = svc_name, startCol = 2, startRow = 3) %T>%
    openxlsx::writeData(tab_name, x = svc_id, startCol = 2, startRow = 4) %T>%
    openxlsx::writeData(tab_name, x = pillar, startCol = 2, startRow = 5) %T>%
    ##budget summary headers
    openxlsx::writeData(tab_name, x = "Budget Summary", startCol = 1, startRow = 7) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12), rows = 7, cols = 1) %T>%
    openxlsx::writeData(tab_name, x = "Budget Amounts", startCol = 3, startRow = 8) %T>%
    openxlsx::mergeCells(tab_name, cols = 3:7, rows = 8) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12, halign = "center", border = c("left","right", "top")), rows = 8, cols = 3) %T>%
    openxlsx::writeData(tab_name, x = "Change Amounts", startCol = 8, startRow = 8) %T>%
    openxlsx::mergeCells(tab_name, cols = 8:9, rows = 8) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12, halign = "center", border = c("left", "right", "top")), rows = 8, cols = 8) %T>%
    ##budget dollars summary data table
    openxlsx::writeData(tab_name, x = "Budget by Fund", startCol = 1, startRow = 10) %T>%
    openxlsx::writeDataTable(
      tab_name, dollar_data, xy = c(2, 9), tableStyle = "TableStyleLight8",
      tableName = paste0("DollarData", svc_id)) %T>%
    ##budget position summary data table
    openxlsx::writeData(tab_name, x = "Positions by Fund", startCol = 1, startRow = dim(dollar_data)[1]+11) %T>%
    openxlsx::writeDataTable(
      tab_name, position_data, xy = c(2, dim(dollar_data)[1]+10), tableStyle = "TableStyleLight8",
      tableName = paste0("PositionData", svc_id)) %T>%
    ##change table
    openxlsx::writeData(tab_name, x = "Change Table (GF Only)", startCol = 11, startRow = 1) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14), rows = 1, cols = 11) %T>%
    openxlsx::writeDataTable(
      tab_name, change_table, xy = c(11, 3),
      tableStyle = "TableStyleLight8", 
      # headerStyle = createStyle(fontSize = 12, textDecoration = "bold", border = "bottom", borderStyle = "thin"),
      tableName = paste0("ChangeTable", svc_id)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 4, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 5, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(indent = 5), rows = c(6,7,8,9,10, 11,12,13), cols = 11) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 15, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 18, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 21, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 24, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#F2F2F2"), rows = 27, cols = c(11:13)) %T>%
    openxlsx::addStyle(tab_name, createStyle(textDecoration = "bold", fgFill = "#D9D9D9"), rows = 30, cols = c(11:13)) %T>%
    ##analysis box
    openxlsx::writeData(tab_name, x = "Analyst Notes", startCol = 1, startRow = 23) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12), rows = 23, cols = 1) %T>%
    openxlsx::addStyle(tab_name, createStyle(border = c("bottom", "top", "left", "right"), borderStyle = "thin"), rows = 24, cols = 1) %T>%
    openxlsx::mergeCells(tab_name, cols = 1:9, rows = 24:29) %T>%
    # openxlsx::writeData(tab_name, x = "Describe major changes, 1 per row", startCol = 1, startRow = 14+dim(data)[1], borders = "surround") %T>%
    ##technical and savings tables
    openxlsx::writeData(tab_name, x = "Tollgate Recommendations", startCol = 15, startRow = 1) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14), rows = 1, cols = 15) %T>%
    openxlsx::writeData(tab_name, x = "Technical Adjustments", startCol = 15, startRow = 3) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12), rows = 3, cols = 15) %T>%
    openxlsx::writeDataTable(
      tab_name, tech_adjust, xy = c(15, 4),
      tableStyle = "TableStyleLight8", 
      # headerStyle = createStyle(fontSize = 12, textDecoration = "bold", border = "bottom", borderStyle = "thin"),
      tableName = paste0("TechnicalAdjustments", svc_id)) %T>%
    openxlsx::writeData(tab_name, x = "Savings Ideas", startCol = 15, startRow = dim(tech_adjust)[1]+6) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 12), rows = dim(tech_adjust)[1]+6, cols = 15) %T>%
    openxlsx::writeDataTable(
      tab_name, savings_ideas, xy = c(15, dim(tech_adjust)[1]+7),
      tableStyle = "TableStyleLight8", 
      # headerStyle =createStyle(fontSize = 12, textDecoration = "bold", border = "bottom", borderStyle = "thin"),
    tableName = paste0("SavingsIdeas", svc_id)) %T>%
    ##sheet settings
    openxlsx::setColWidths(tab_name, 1, widths = 20) %T>%
    openxlsx::setColWidths(tab_name, 2, widths = 20) %T>%
    openxlsx::setColWidths(tab_name, 3:9, widths = 15) %T>%
    openxlsx::setColWidths(tab_name, 10, widths = 5) %T>%
    openxlsx::setColWidths(tab_name, 11, widths = 50) %T>%
    openxlsx::setColWidths(tab_name, 12, widths = 15) %T>%
    openxlsx::setColWidths(tab_name, 13, widths = 20) %T>%
    openxlsx::setColWidths(tab_name, 14, widths = 5) %T>%
    openxlsx::setColWidths(tab_name, 15, widths = 30) %T>%
    openxlsx::pageSetup(tab_name, printTitleRows = 1) # repeat first row when printing
  
  if (save == TRUE) {
    openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)
    base::message(tab_name, ' tab created in the file saved as ', file_name)
  } else {
    base::message(tab_name, ' not saved. Use openxlsx::saveWorkbook().')
    return(excel)
  }
}

make_agency_summary <- function(agency, file_name) {
  
  options("openxlsx.orientation" = "landscape",
          "openxlsx.datetimeFormat" = "yyyy-mm-dd",
          "openxlsx.numFmt" = "#,##0;-#,##0;-")
  
  tab_name = "Agency Summary"
  
  summary_dollars <- summary %>%
    ungroup() %>%
    filter(`Agency Name` == agency & Type == "Expenditures") %>%
    unite("Service", c(`Service ID`, `Service Name`), sep = ": ") %>%
    select(-`Fund Name`, -`Objective Name`, -`Agency Name`, -Type) %>%
    arrange(Service) %>%
    group_by(Service) %>%
    summarise_if(is.numeric, sum, na.rm = TRUE) %>%
    adorn_totals() %>%
    mutate(`Change $` = `FY24 Request`- `FY24 CLS`)
  
  summary_positions <- summary %>%
    ungroup() %>%
    filter(`Agency Name` == agency & Type == "Positions") %>%
    unite("Service", c(`Service ID`, `Service Name`), sep = ": ") %>%
    select(-`Fund Name`, -`Objective Name`, -`Agency Name`, -Type) %>%
    arrange(Service) %>%
    group_by(Service) %>%
    summarise_if(is.numeric, sum, na.rm = TRUE) %>%
    adorn_totals() %>%
    mutate(`Change #` = `FY24 Request`- `FY24 CLS`)
  
  agency_sheet <- openxlsx::loadWorkbook(file_name) %T>%
    openxlsx::addWorksheet(
      tab_name,
      header = c(gsub('\\..*', '', file_name), "&[Tab]", as.character(Sys.Date())),
      footer = c(NA, "&[Page]", NA)) %T>%
    openxlsx::writeData(tab_name, x = agency, startCol = 1, startRow = 1) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14, textDecoration = "bold", border = "bottom", borderStyle = "thin"), rows = 1, cols = 1) %T>%
    ##dollars by service
    openxlsx::writeData(tab_name, x = "Dollars by Service", startCol = 1, startRow = 4) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14,  border = "bottom", borderStyle = "thin"), rows = 4, cols = 1) %T>%
    openxlsx::writeDataTable(
      tab_name, summary_dollars, xy = c(1, 5),
      tableStyle = "TableStyleLight8", 
      # headerStyle = createStyle(fontSize = 12, textDecoration = "bold", border = "bottom", borderStyle = "thin"),
      tableName = paste0("DollarsbyService")) %T>%
    ##positions by service
    openxlsx::writeData(tab_name, x = "Positions by Service", startCol = 1, startRow = dim(summary_dollars)[1]+7) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14,  border = "bottom", borderStyle = "thin"), rows = dim(summary_dollars)[1]+7, cols = 1) %T>%
    openxlsx::writeDataTable(
      tab_name, summary_positions, xy = c(1, dim(summary_dollars)[1]+8),
      tableStyle = "TableStyleLight8", 
      headerStyle = createStyle(fontSize = 12),
      tableName = paste0("PositionsbyService")) %T>%
    ##notes
    openxlsx::writeData(tab_name, x = "Agency Budget Highlights", startCol = 8, startRow = 4) %T>%
    openxlsx::addStyle(tab_name, createStyle(fontSize = 14), rows = 4, cols = 8) %T>%
    # openxlsx::writeData(tab_name, x = "Add text here", startCol = dim(summary_dollars)[2]+2, startRow = 2) %T>%
    openxlsx::mergeCells(tab_name, cols = 8:14, rows = 5:20)%T>%
    openxlsx::setColWidths(tab_name, 8, widths = 45) %T>%
    openxlsx::setColWidths(tab_name, 1, widths = 45) %T>%
    openxlsx::setColWidths(tab_name, 2:6, 16)

  openxlsx::saveWorkbook(agency_sheet, file_name, overwrite = TRUE)
  base::message(tab_name, ' tab created in the file saved as ', file_name)
}

export_template <- function(agency, data) {
  
  file_name <- paste0("outputs/FY", params$fy, " Change Table ", 
                      analysts$`Agency Name - Cleaned`[analysts$`Agency Name` == agency], 
                       ".xlsx")
  
  agency <- gsub("&", "and", agency)
  
  output <- data %>% 
    filter(`Agency Name` == agency) %>% 
    arrange(`Service ID`)
  
  svcs <- output %>%
    group_by(`Service ID`) %>%
    ungroup() %>%
    extract2("Service ID") %>%
    unique() %>%
    sort()
  
  #Excel
  if (nrow(output) > 0) {
    n = 1
    for (i in svcs) {
      
      df <- output %>% 
        filter(`Service ID` == i)
      
      oso <- osos %>%
        filter(`Service ID` == i)

        excel <- suppressMessages(
          make_template(df = df, oso = oso, svc = i, tab_name = i, file_name = file_name,
                       save = FALSE,
                       type = ifelse(i == svcs[1], "new", "existing")))
        
        ##style elements
        # mergeCells(excel, n, cols = 5, rows = 2:12)
        # style <- createStyle(wrapText = TRUE) %T>%
        # addStyle(wb, sheet = n, style, cols = 5:6, rows = 2:12, gridExpand = TRUE)
        # setColWidths(wb, i, 1, widths = 45, hidden = FALSE) 
        # setColWidths(wb, i, 2, widths = 30, hidden = FALSE) 
        # setColWidths(wb, i, 3:6, widths = 15, hidden = FALSE)
        # writeFormula(excel, n, x = "CHAR(10)", startCol = 5, startRow = 2)
        #freeze notes cells next to budget line

      openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)    
      n = n + 1
      # cat(".") # progress bar of sorts
      
      cat("Change table sheet saved in", file_name, "\n")}
    
    make_agency_summary(agency, file_name)
  }}

##testing
export_template("Sheriff", summary)

map(agencies, export_template, summary)