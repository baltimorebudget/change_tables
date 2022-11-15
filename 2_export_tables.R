#make change table==================================
#make rows for all object id's (0-9)
obj_ids <- change_tables$gen_fund %>%
  select(`Object ID`, `Object Name`) %>%
  unique() 

svc_ids <- change_tables$gen_fund %>%
  select(`Service ID`, `Service Name`) %>%
  unique() 

obj_base <- crossing(obj_ids, svc_ids) %>%
  filter(`Object ID` != 2)

#read in analyst notes ===================
# notes <- list()
# 
# #get all files
# notes$files <- list.files(".", pattern = "*2022-09-02.xlsx", full.names = TRUE, recursive = TRUE)
# 
# #get metadata for files
# notes$df = data.frame(Agency = character(), Services = factor(), File = character())
# for (f in notes$files) {
#   xtrct = str_extract_all(f, '([^/]+)(?:[^/])')
#   notes$df <- notes$df %>% add_row(Agency = xtrct[[1]][3], Services = c(excel_sheets(f)), File = f)
# }
# 
# ##check
# notes$sheets <- map(notes$files, excel_sheets)
# notes$services = vector(mode = "list", length = 0)
# for (lst in notes$sheets) {
#   for (i in lst) {
#     notes$services <- c(notes$services, i)
#   }
# }
# 
# if (dim(notes$df)[1] == length(notes$services)) {
#   print("All services accounted for.")} else {
#     print("Services missing.")
#   }
# 
# gf_svcs <- unique(change_tables$gen_fund$`Service ID`)
# `%!in%` <- Negate(`%in%`)
# for (i in notes$services) {
#   if (i %!in% gf_svcs) {
#     print(paste(i, " is not in General Fund."))
#   }
# }
# 
# 
# #pull in data from each previous sheet in each file 
# df_notes = data.frame()
# for (f in notes$df$File) {
#   svcs <- notes$df$Services[notes$df$File == f]
#   for (s in svcs) {
#     mini_df <- read_excel(f, s)
#     df_notes <- rbind(df_notes, mini_df)
#   }
# }
# 
# ##check
# for (s in unique(df_notes$`Service ID`)) {
#   if (s %!in% notes$services) {
#     print(paste(s, " is missing."))
#   }
# }

#create change table parent file ===================================
change_table <- left_join(change_tables$obj$summary, position_summary, by = c("Service ID", "Object ID")) %>%
  # left_join(flat_adjustments, by = c("Service ID", "Object ID")) %>%
  mutate(`Position Note` = case_when(`Object ID` == 1 & `Position Note` != "" ~ paste0(`Position Note`, " at a cost of ", `Total Cost Diff`),
                                     TRUE ~ ""),
         `Changes or adjustments` = case_when(
           Diff > 0 ~ paste("Increase in", `Object Name`),
           Diff < 0 ~ paste("Decrease in", `Object Name`),
           Diff == 0 ~ paste("No change in", `Object Name`),
           TRUE ~ `Object Name`)) %>%
  select(`Service ID`, `Object ID`, `Changes or adjustments`, `Diff`, `Position Note`) %>%
  group_by(`Service ID`, `Object ID`, `Changes or adjustments`, `Position Note`) %>%
  summarise(`Difference` = sum(`Diff`, na.rm = TRUE), .groups = "drop") %>%
  mutate(`Analyst Notes` = "") %>%
  select(`Service ID`, `Object ID`, `Changes or adjustments`, `Difference`, `Position Note`, `Analyst Notes`) %>% 
  arrange(desc(`Service ID`))

change_table_ids <- full_join(obj_base, change_table, by = c("Service ID", "Object ID")) %>%
  select(`Service ID`, `Object ID`, `Changes or adjustments`, `Difference`, `Position Note`, `Analyst Notes`)

#budget data
budget <- change_tables$gen_fund %>%
  pivot_longer(cols = c(!!sym(cols[1]), !!sym(cols[2])),
               names_to = "Changes or adjustments",
               values_to = "Amount") %>%
  mutate(`Changes or adjustments` = case_when(
    `Changes or adjustments` == cols[1] ~
      paste0(cols[1], " Budget"),
    `Changes or adjustments` == cols[2] ~
      paste0(cols[2], " Budget"),
    TRUE ~ `Changes or adjustments`)) %>%
  group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`, `Changes or adjustments`) %>%
  summarize(Amount = sum(Amount, na.rm = TRUE))

svcs <- change_tables$gen_fund %>% select(`Service ID`) %>% unique()

#combine data
change_tables$df <- budget %>% filter(`Changes or adjustments` == paste0(cols[1], " Budget")) %>%
  bind_rows(change_table_ids) %>%
  bind_rows(budget %>% filter(`Changes or adjustments` == paste0(cols[2], " Budget"))) %>%
  mutate(`Amount` = case_when(is.na(`Amount`) ~ `Difference`,
                              TRUE ~ `Amount`),
         `Notes` = "",
         `Changes or adjustments` = case_when(is.na(`Changes or adjustments`) & `Object ID` == 0 ~ "No Transfers",
                                              is.na(`Changes or adjustments`) & `Object ID` == 1 ~ "No Employee Costs",
                                              is.na(`Changes or adjustments`) & `Object ID` == 3 ~ "No Contractual Services",
                                              is.na(`Changes or adjustments`) & `Object ID` == 4 ~ "No Materials and Supplies",
                                              is.na(`Changes or adjustments`) & `Object ID` == 5 ~ "No Equipment < $5K",
                                              is.na(`Changes or adjustments`) & `Object ID` == 6 ~ "No Equipment > $5K",
                                              is.na(`Changes or adjustments`) & `Object ID` == 7 ~ "No Grants, Subsidies or Contributions",
                                              is.na(`Changes or adjustments`) & `Object ID` == 8 ~ "No Debt Service",
                                              is.na(`Changes or adjustments`) & `Object ID` == 9 ~ "No Capital Improvements",
                                              TRUE ~ `Changes or adjustments`)) %>%
  ungroup() %>%
  rename(`Automated Notes` = `Position Note`) %>%
  select(`Service ID`, `Object ID`, `Changes or adjustments`, `Amount`, `Automated Notes`, `Analyst Notes`)


#add agency back in for filtering
ag_svc <- change_tables$gen_fund %>% select(`Agency Name`, `Service ID`) %>% unique()
change_tables$agency <- change_tables$df %>% left_join(ag_svc, by = "Service ID") 

change_tables$agency$`Object ID`[change_tables$agency$`Object ID` == 1] <- "1, 2"

##testing only!! add random strings to Analyst Notes to test carry forward process
# change_tables$agency$`Analyst Notes` <- random::randomStrings(n = dim(change_tables$agency)[1], len = 10, digits = FALSE)

##add analyst notes from previous file!!
# test <- left_join(change_tables$agency, df_notes, by = c("Service ID", "Object ID", "Agency Name")) %>%
#   rename("Analyst Notes" = "Analyst Notes.y", "Changes or adjustments" = "Changes or adjustments.x", "Amount" = "Amount.x", "Automated Notes" = "Automated Notes.x") %>%
#   select(`Agency Name`, `Service ID`, `Object ID`, `Changes or adjustments`, `Amount`, `Automated Notes`, `Analyst Notes`)

##remove duplicates! from blank Object ID joins -OR- tell analysts not to enter notes there.

##make folders and file paths=================

file_info <- list(
  analysts = import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
    filter(!is.na(Analyst) & `Operational` == "TRUE") %>%
    mutate(#`Agency ID` = as.character(`Agency ID`),
      `File Start` = paste0("outputs/", `Agency Name - Cleaned`)),
  agencies = import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
    filter(`Operational` == "TRUE") %>%
    select(`Agency Name`),
  short = import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
    filter(`Operational` == "TRUE") %>%
    select(`Agency Name`, `Agency Short`)
)

# create_analyst_dirs <- function(path, additional_folders = NULL) {
#   
#   if (missing(path)) {
#     path <- getwd()
#   }
#   
#   # analysts <- unique(file_info$analysts$Analyst)
#   file_path <- file_info$analysts$`File Start`
#   
#   if (stringr::str_trunc(path, 1, side = "left", ellipsis = "") == "/") {
#     dirs <- paste0(path, file_path)
#   } else {
#     dirs <- paste0(path, "/", file_path)
#   }
#   
#   dirs %>%
#     sapply(dir.create, recursive = TRUE)
# }
# 
# create_analyst_dirs()

agencies <- file_info$agencies$`Agency Name` %>% unique()

##export services by agency and analyst=========
export_change_table_file <- function(agency, change_table_df) {
  
  file_name <- paste0("outputs/Change Table_", 
                      file_info$short$`Agency Short`[file_info$short$`Agency Name` == agency], 
                      "_FY", 
                      params$start_yr, params$start_phase, 
                      "-FY", params$end_yr, params$end_phase, "_", Sys.Date(), ".xlsx")
  
  agency <- gsub("&", "and", agency)
  
  output <- change_table_df %>% filter(`Agency Name` == agency) %>% arrange(`Service ID`)
  
  svcs <- output %>%
    group_by(`Service ID`) %>%
    # remove any services without a budget in current AND target FY
    filter(any(Amount != 0)) %>%
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
      
      excel <- suppressMessages(
        export_excel(df = df, tab_name = i, file_name = file_name,
                     save = FALSE,
                     type = ifelse(i == svcs[1], "new", "existing")))
      
      ##style elements
      # mergeCells(excel, n, cols = 5, rows = 2:12)
      style <- createStyle(wrapText = TRUE)
      addStyle(excel, sheet = n, style, cols = 5:6, rows = 2:12, gridExpand = TRUE)
      setColWidths(excel, n, 5, widths = 45, hidden = FALSE)
      # writeFormula(excel, n, x = "CHAR(10)", startCol = 5, startRow = 2)
      #freeze notes cells next to budget lines
      
      openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)
      n = n + 1
      # cat(".") # progress bar of sorts
      
      cat("Change table file saved as", file_name, "\n")}
  }}

map(agencies, export_change_table_file, change_tables$agency)


##export full data ===========================================
export_excel(change_tables$agency, "By Agency", paste0(getwd(),"/outputs/FY", params$start_yr, " ", params$start_phase, " - FY", params$end_yr, " ", params$end_phase, " OSO Changes.xlsx"), type = "new")
