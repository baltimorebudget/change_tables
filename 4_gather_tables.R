library(janitor)
library(purrr)
library(readxl)

files <- list.files(path = "C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/",
                   pattern = paste0("^FY24 Change Table*"),
                   full.names = TRUE, recursive = TRUE)

import_tech_tables <- function(files) {
 for (file in files) {
    sheets = na.omit(str_extract(excel_sheets(file), "\\d{3}"))
    for (s in sheets) {
      df = read_excel(file, s)
      return(df)
    }
  } 
}


data <- map_df(files, import_tech_tables)

df <- data %>%
  select(`...2`, `Tollgate Recommendations`:`...19`) 

x <- remove_empty(df, which = "rows")
