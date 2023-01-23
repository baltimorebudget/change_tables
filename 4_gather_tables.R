library(janitor)
library(purrr)
library(readxl)

files <- list.files(path = "C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/",
                   pattern = paste0("^FY24 Change Table*"),
                   full.names = TRUE, recursive = TRUE)

import_tech_tables <- function(files) {
  for (file in files) {
    sheets = na.omit(str_extract(excel_sheets(file), "\\d{3}"))
    df = data.frame()
    for (s in sheets) {
      x = read_excel(file, s) %>%
        mutate(ID = s)
      print(paste0(s, " added from ", file))
      df = rbind(df, x)
    }
    return(df)
  } 
}


data <- map_df(files, import_tech_tables)

df <- data %>%
  select(ID, `Tollgate Recommendations`:`...19`) %>%
  filter((!is.na(`...16`) & `...16` != "Object") & (!is.na(`...17`) & `...17` != "Subobject") & 
           (!is.na(`...18`) & `...18` != "Amount") & (!is.na(`...19`) & `...19` != "Decision")) 

