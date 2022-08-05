#make change table==================================

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
  select(`Service ID`, `Changes or adjustments`, `Difference`, `Position Note`, `Analyst Notes`) %>% 
  arrange(desc(`Service ID`))

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


#combine data
change_tables$df <- budget %>% filter(`Changes or adjustments` == paste0(cols[1], " Budget")) %>%
  bind_rows(change_table) %>%
  bind_rows(budget %>% filter(`Changes or adjustments` == paste0(cols[2], " Budget"))) %>%
  mutate(`Amount` = case_when(is.na(`Amount`) ~ `Difference`,
                              TRUE ~ `Amount`),
         `Notes` = "") %>%
  ungroup()%>%
  select(`Service ID`, `Changes or adjustments`, `Amount`, `Position Note`, `Analyst Notes`)


#add agency back in for filtering
ag_svc <- change_tables$gen_fund %>% select(`Agency Name`, `Service ID`) %>% unique()
change_tables$agency <- change_tables$df %>% left_join(ag_svc, by = "Service ID")


##make folders and file paths=================

file_info <- list(
  analysts = import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
    filter(!is.na(Analyst) & `Operational` == "TRUE") %>%
    mutate(#`Agency ID` = as.character(`Agency ID`),
      `File Start` = paste0("outputs/", Analyst,
                            "/", `Agency Name - Cleaned`)),
  agencies = import("G:/Analyst Folders/Sara Brumfield/_ref/Analyst Assignments.xlsx") %>%
    filter(`Operational` == "TRUE") %>%
    select(`Agency Name`)
)

create_analyst_dirs <- function(path, additional_folders = NULL) {
  
  if (missing(path)) {
    path <- getwd()
  }
  
  # analysts <- unique(file_info$analysts$Analyst)
  file_path <- file_info$analysts$`File Start`
  
  if (stringr::str_trunc(path, 1, side = "left", ellipsis = "") == "/") {
    dirs <- paste0(path, file_path)
  } else {
    dirs <- paste0(path, "/", file_path)
  }
  
  dirs %>%
    sapply(dir.create, recursive = TRUE)
}

create_analyst_dirs()

agencies <- file_info$agencies$`Agency Name` %>% unique()

##export services by agency and analyst=========
export_change_table_file <- function(agency, change_table_df) {

  file_name <- paste0(file_info$analysts$`File Start`[file_info$agencies == agency], "/Change Table", "_FY", params$start_yr, params$start_phase, "-FY", params$end_yr, params$end_phase, ".xlsx")
  
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
    
    for (i in svcs) {
      
      df <- output %>% 
        filter(`Service ID` == i)
      
      excel <- suppressMessages(
        export_excel(df = df, tab_name = i, file_name = file_name,
                     save = FALSE,
                     type = ifelse(i == svcs[1], "new", "existing")))
      
      openxlsx::saveWorkbook(excel, file_name, overwrite = TRUE)

     # cat(".") # progress bar of sorts

  cat("Change table file saved as", file_name, "\n")}
}}

map(agencies, export_change_table_file, change_tables$agency)


##export full data ===========================================
export_excel(change_tables$oso$detail, "OSO Detail", paste0(getwd(),"/outputs/FY23 CLS - COU OSO Changes.xlsx"), type = "new")
export_excel(change_tables$oso$summary, "OSO Summary", paste0(getwd(), "/outputs/FY23 CLS - COU OSO Changes.xlsx"), type = "existing")