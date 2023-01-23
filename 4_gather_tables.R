library(janitor)
library(purrr)
library(readxl)

#get data from SharePonit files
files <- list.files(path = "C:/Users/sara.brumfield2/OneDrive - City Of Baltimore/FY2024 Planning/03-TLS-BBMR Review/Agency Analysis Tools/",
                   pattern = paste0("^FY24 Change Table*"),
                   full.names = TRUE, recursive = TRUE)

import_tech_tables <- function(files) {
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


data <- map(files, import_tech_tables) 

df = data %>%
  map(select, c(Agency, ID, `Tollgate Recommendations`:`...18`)) %>%
  bind_rows() 

#check for all services
length(unique(df$ID))

#clean up the file
output <- df %>%
  filter((!is.na(`...3`) & `...3` != "Object") & (!is.na(`...4`) & `...4` != "Subobject") & 
           (!is.na(`...5`) & `...5` != "Amount")) %>%
  rename(`Service ID` = ID, `Object` = `...3`, `Subobject` = `...4`, `Amount` = `...5`, `BPFS Adjustment` = `Tollgate Recommendations`)



