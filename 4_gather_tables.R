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

import_tables <- function(files) {
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

data <- map(files, import_tables) 

df = data %>%
  map(select, c(Agency, ID, `Tollgate Recommendations`:`...18`)) %>%
  bind_rows() 

#check for all services
length(unique(df$ID))

#clean up the file
output <- df %>%
  filter(!(`Tollgate Recommendations` %in% c("None", "N/A", "Technical Adjustments", "Item", "Enter item here", "Total", "Savings Ideas", "Tollgate Notes")) &
           !is.na(`Tollgate Recommendations`)) %>%
  # filter((!is.na(`...4`) & `...4` != "Object") & (!is.na(`...5`) & `...5` != "Subobject") & 
  #          (!is.na(`...6`) & `...6` != "Amount")) %>%
  rename(`Service ID` = ID, `Object` = `...4`, `Subobject` = `...5`, `Amount` = `...6`, `BPFS Adjustment` = `Tollgate Recommendations`)

#export
write.csv(output, "outputs/FY24 Propoposal Technical Adjustments for BPFS.csv")

##change tables ======

data <- map(files, import_tables) 

df = data %>%
  map(select, c(Agency, ID, `Service Summary`:`...2`, `Change Table (GF Only)`, `...12`, `...13`)) %>%
  map(~ mutate(., Pillar = ...2[`Service Summary`=="Pillar"][2])) %>%
  map(~ mutate(., Service = ...2[`Service Summary`=="Service Name"][2])) %>%
  map(filter, !is.na(`Change Table (GF Only)`)) %>%
  map(select, -`Service Summary`, -`...2`) %>%
  map(rename, `Amount` = `...12`, `Tollgate Decision` = `...13`) %>%
  bind_rows()

agencies <- df %>% select(Agency) %>% distinct()

for(a in agencies){
  
  file_path <- paste0("outputs/", a," Change Tables.xlsx")
  
  x <- df %>% filter(Agency == a) %>%
    select(-Pillar, -Service, -Agency)
  
  header_style <- createStyle(fgFill = "white", border = "TopBottomLeftRight",
                              borderColour = "black", textDecoration = "bold", fontColour = "darkblue",
                              wrapText = TRUE)
  
  style <- createStyle(fgFill = "darkblue", border = "TopBottomLeftRight",
                       borderColour = "black", textDecoration = "bold", fontColour = "white",
                       wrapText = TRUE)
  
  style_rows <- which(x$`Amount`=="Amount") + 1
  
  wb <- createWorkbook()
  addWorksheet(wb, a)
  writeDataTable(wb, 1, x = x)
  
  addStyle(wb, 1, style = style, rows = style_rows, cols = 1:5, gridExpand = TRUE, stack = FALSE)
  addStyle(wb, 1, style = header_style, rows = 1, cols = 1:5, gridExpand = TRUE, stack = FALSE)
  setColWidths(wb, 1, cols = 1, widths = 35)
  setColWidths(wb, 1, cols = 2, widths = 10)
  setColWidths(wb, 1, cols = 3, widths = 45)
  setColWidths(wb, 1, cols = 4:5, widths = 17)
  
  saveWorkbook(wb, file_path, overwrite = TRUE)
  
  message(a, " change tables exported.")
}
