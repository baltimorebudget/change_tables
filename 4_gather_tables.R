.libPaths("C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/Documents/r_library")
library(janitor)
library(purrr)
library(readxl)
library(stringr)
library(dplyr)
library(tidyverse)
library(magrittr)

devtools::load_all("G:/Analyst Folders/Sara Brumfield/_packages/bbmR")

#get data from SharePonit files
files <- list.files(path = "C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/",
                   pattern = paste0("^FY24 Change Table*"),
                   full.names = TRUE, recursive = TRUE)
extract_table <- function(df) {
  table <- df[which(df$`Tollgate Recommendations`=="Technical Adjustments"):which(df$`Tollgate Recommendations`=="Total")[1],]
  return(table)
}

import_tech_tables <- function(files) {
  for (file in files) {
    agency = str_extract(file, "((?<=C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/FY24 Change Table ).+(?=.xlsx))")
    sheets = na.omit(str_extract(excel_sheets(file), "\\d{3}"))
    df = data.frame()
    for (s in sheets) {
      z = read_excel(file, s) %>% 
        select(`Service Summary`, `...2`, `Tollgate Recommendations`:`...18`) %>%
        extract_table() %>%
        mutate(ID = s,
               Agency = agency)
      # x = read_excel(file, s) %>%
      #   mutate(ID = s,
      #          Agency = agency)
      print(paste0(s, " added from ", file))
      df = rbind(df, z)
    }
    return(df)
  } 
}

import_change_tables <- function(files) {
  for (file in files) {
    agency = str_extract(file, "((?<=C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/FY24 Change Table ).+(?=.xlsx))")
    sheets = na.omit(str_extract(excel_sheets(file), "\\d{3}"))
    df = data.frame()
    for (s in sheets) {
      x = read_excel(file, s) %>%
        mutate(ID = s,
               Agency = agency)
      print(paste0(s, " added from ", file))
      df = rbind(df, x)
    }
    return(df)
  } 
}

##technical adjustments =====

data <- map(files, import_tech_tables) %>%
  bind_rows()

##clean up file
df = data %>%
  select(-`Service Summary`, -`...2`) %>%
  relocate(Agency, .before = `Tollgate Recommendations`) %>%
  relocate(ID, .after = Agency) %>%
  rename(`Service ID` = ID, `Object` = `...16`, `Subobject` = `...17`, `Amount` = `...18`, `BPFS Adjustment` = `Tollgate Recommendations`) %>%
  filter()


#check for all services
length(unique(df$ID))

#clean up the file
# output <- x %>%
#   # filter(!(`Tollgate Recommendations` %in% c("None", "N/A", "Technical Adjustments", "Item", "Enter item here", "Total", "Savings Ideas", "Tollgate Notes")) &
#   #          !is.na(`Tollgate Recommendations`)) %>%
#   # filter((!is.na(`...4`) & `...4` != "Object") & (!is.na(`...5`) & `...5` != "Subobject") & 
#   #          (!is.na(`...6`) & `...6` != "Amount")) %>%
#   rename(`Service ID` = ID, `Object` = `...16`, `Subobject` = `...17`, `Amount` = `...18`, `BPFS Adjustment` = `Tollgate Recommendations`) %>%
#   filter(`BPFS Adjustment` != "Item")

#export
export_excel(df, "outputs/FY24 Propoposal Technical Adjustments for BPFS.xlsx")

##change tables ======

data <- map(files, import_change_tables) 

df = data %>%
  map(select, c(Agency, ID, `Service Summary`:`...2`, `Change Table (GF Only)`, `...12`, `...13`)) %>%
  # map(~ mutate(., Pillar = ...2[`Service Summary`=="Pillar"][2])) %>%
  # map(~ mutate(., Service = ...2[`Service Summary`=="Service Name"][2])) %>%
  map(filter, !is.na(`Change Table (GF Only)`)) %>%
  map(select, -`Service Summary`, -`...2`) %>%
  map(rename, `Amount` = `...12`, `Tollgate Decision` = `...13`) %>%
  bind_rows()

pillar_map <- readRDS("G:/Budget Publications/automation/0_data_prep/outputs/fy24_tls/expenditure.Rds") %>% select( `Service ID`, `Service Name`, `Objective Name`) %>% distinct()

pillar_join <- df %>% left_join(pillar_map, by = c("ID" = "Service ID")) %>%
  # select(-Pillar) %>%
  rename(Pillar = `Objective Name`, Service = `Service Name`)

pillars <- pillar_join %>% 
  group_by(`Pillar`) %>%
  ungroup() %>%
  extract2("Pillar") %>%
  unique() %>%
  sort()

consolidate_tables <- function(df = x, tab_name = a, type, file_name = file_path, save= TRUE) {
  
  num_style <- createStyle(numFmt = "#,##0;(#,##0)")
  
  header_style <- createStyle(fgFill = "white", border = "TopBottomLeftRight",
                              borderColour = "black", textDecoration = "bold", fontColour = "darkblue",
                              wrapText = TRUE)
  
  style <- createStyle(fgFill = "darkblue", border = "TopBottomLeftRight",
                       borderColour = "black", textDecoration = "bold", fontColour = "white",
                       wrapText = TRUE)
  
  style2 <- createStyle(textDecoration = c("bold", "italic"),
                        wrapText = TRUE)
  
  style_rows <- which(x$`Change Table (GF Only)`=="Adjustments") + 1
  
  style_rows2 <- which(x$`Change Table (GF Only)` %in% c("FY2023 Adopted", "CLS Adjustments", "Request Adjustments")) + 1
  
  num_rows <- which(x$`Amount`!="Amount") + 1
  
  excel <- switch(type,
                  "new" = openxlsx::createWorkbook(),
                  "existing" = openxlsx::loadWorkbook(file_name)) %T>%
    openxlsx::addWorksheet(tab_name) %T>%
    openxlsx::writeDataTable(tab_name, x = df) %T>%
    addStyle(tab_name, style = style, rows = style_rows, cols = 1:4, gridExpand = TRUE, stack = FALSE) %T>%
    addStyle(tab_name, style = header_style, rows = 1, cols = 1:4, gridExpand = TRUE, stack = FALSE) %T>%
    addStyle(tab_name, style = style2, rows = style_rows2, cols = 1:4, gridExpand = TRUE, stack = FALSE) %T>%
    addStyle(tab_name, style = num_style, rows = num_rows, cols = 3, gridExpand = TRUE, stack = FALSE) %T>%
    setColWidths(tab_name, cols = 1, widths = 10) %T>%
    setColWidths(tab_name, cols = 2, widths = 45) %T>%
    setColWidths(tab_name, cols = 3:4, widths = 17)
  
  saveWorkbook(excel, file_name, overwrite = TRUE)

  message(a, " change tables exported.")
  # 
  # if (save == TRUE) {
  #   openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)
  #   base::message(tab_name, ' tab created in the file saved as ', file_name)
  # } else {
  #   base::message(tab_name, ' not saved. Use openxlsx::saveWorkbook().')
  #   return(excel)
  # }
}


for (p in pillars) {
  
  file_path <- paste0("outputs/", p," Pillar Change Tables.xlsx")
  
  pillar_data <- pillar_join %>% filter(Pillar == p & Amount != "Pending" & `Change Table (GF Only)` != "TLS Adjustments")
  
  agencies <- pillar_data %>%
    group_by(`Agency`) %>%
    ungroup() %>%
    extract2("Agency") %>%
    unique() %>%
    sort()
    
  for(a in agencies){
    
    x <- pillar_data %>% filter(Agency == a) %>%
      select(-Pillar, -Service, -Agency) %>%
      mutate(Amount = as.numeric(Amount))
  
    excel <- suppressMessages(
      consolidate_tables(df = x, tab_name = substr(a, 1,30), file_name = file_path,
                    save = FALSE,
                    type = ifelse(a == agencies[[1]], "new", "existing")))
    
    # openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)    
    # 
    cat(a, "change table saved in", p, "\n")
  }
  
  # openxlsx::saveWorkbook(excel, file_path, overwrite = TRUE)   
  # message(p, " change tables exported.")
  
}

##performance data to accompany change tables
pms_join <- pms %>% left_join(pillar_map, by = "Service ID")

for (p in pillars) {
  x = pms_join %>% filter(`Objective Name` == p) %>%
    select(-`Extra PM Table`) %>%
    relocate(`Agency Name - Cleaned`, .before = `Service ID`)
  
  export_excel(x, tab_name = substr(p, 1, 30), file_name = paste0("outputs/", p, " Performance Measures.xlsx"))
}