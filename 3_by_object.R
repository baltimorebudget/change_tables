##combines R scripts 0-2

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

##read in data =====================
cols <- c(paste0("FY", params$start_yr, " ", params$start_phase), 
          paste0("FY", params$end_yr, " ", params$end_phase), 
          paste0("FY", params$start_yr, " COU"),
          paste0("FY", params$start_yr, " All Positions File"))

line_item_start <- readxl::read_excel(path = params$line.start, sheet = "Details") %>%
  rename(`Service ID` = `Program ID`,
         `Service Name` = `Program Name`)

line_item_end <- readxl::read_excel(path = params$line.end, sheet = "Details") %>%
  rename(`Service ID` = `Program ID`,
         `Service Name` = `Program Name`)

change_tables <- list(
  gen_fund = line_item_end %>%
    filter(`Fund ID` == 1001 & !is.na(`Agency ID`)) %>%
    select(`Agency ID`:`Subobject Name`,
           !!sym(cols[1]), !!sym(cols[2]),
           Justification),
  position = list(),
  obj = list(),
  tags = categorize_tags()
)

#objects=====================
change_tables$obj$detail <- change_tables$gen_fund %>%
  group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`,
           `Object ID`, `Object Name`, Justification) %>%
  summarize(!!sym(cols[1]) := sum(!!sym(cols[1]), na.rm = TRUE),
            !!sym(cols[2]) := sum(!!sym(cols[2]), na.rm = TRUE), .groups = "drop") %>%
  mutate(Diff = !!sym(cols[2]) - !!sym(cols[1])) %>%
  ungroup() %>%
  mutate(`Percent Change` = Diff / !!sym(cols[1]),
         # `Signif Diff` = ifelse((Diff > 5000 | Diff < -5000) & 
         #                          (`Percent Change` >= 0.2 | `Percent Change` <= -0.2), "Yes", "No"),
         `Percent Change` = scales::percent(`Percent Change`, accuracy = 0.1))

change_tables$obj$summary <-  change_tables$obj$detail %>%
  mutate(`Object ID` = case_when(`Object ID` %in% c(1,2) ~ 1,
                                 TRUE ~ `Object ID`),
         `Object Name` = case_when(`Object Name` %in% c("Salaries", "Other Personnel Costs") ~ "All Employee Costs",
                                   TRUE ~ `Object Name`)) %>%
  group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`, `Object ID`, `Object Name`) %>%
  summarize(!!sym(cols[1]) :=
              sum(!!sym(cols[1]), na.rm = TRUE),
            !!sym(cols[2]) :=
              sum(!!sym(cols[2]), na.rm = TRUE),
            Diff= sum(Diff, na.rm = TRUE), .groups = "drop") %>%
  ungroup()

#positions=================
##Year start requires some manual changes.
position_start <- readxl::read_excel(path = params$position.start, sheet = cols[4])
position_end <- readxl::read_excel(path = params$position.end, sheet = "PositionsSalariesOPCs")

pos_cols <- list( # dynamic col names based on FY
  base = list(
    salary = "Salary FY",
    opcs = "OPCs FY",
    total_cost = "Total Cost FY",
    service_id = "Service ID FY",
    service_name = "Service Name FY",
    class_id = "Classification ID FY",
    class_name = "Classification Name FY",
    fund_id = "Fund ID FY",
    fund_name = "Fund Name FY"))

pos_cols$projection <- pos_cols$base %>%
  map(function(x) paste0(x, params$start_yr, " ", params$start_phase)) %>%
  map(sym)

pos_cols$planning <- pos_cols$base %>%
  map(function(x) paste0(x, params$end_yr, " ", params$end_phase)) %>%
  map(sym)


clean_pos_files <- function(df) {
  df <- df %>%
    select(-starts_with("OSO")) %>%
    mutate_if(is.numeric , replace_na, replace = 0) %>%
    mutate(OPCs = `TOTAL COST` - Salary) 
  return(df)
}

position_start <- clean_pos_files(position_start)
position_end <- clean_pos_files(position_end)

df <- list(
  create = anti_join(position_end, position_start, by = "JOB NUMBER") %>%
    # XX is a placeholder for the number of positions;
    # this will be use and replaced by calc_change_table()
    mutate(`Change Single` = paste("create XX", `Classification Name`, "position"),
           !!pos_cols$projection$salary := NA,
           !!pos_cols$projection$opcs := NA,
           !!pos_cols$projection$total_cost := NA,
           `Change Net` = 1) %>%
    rename(!!pos_cols$planning$salary := Salary,
           !!pos_cols$planning$opcs := OPCs,
           !!pos_cols$planning$total_cost := `TOTAL COST`) %>%
    filter(`FUND NAME` == "General"))

df$eliminate <- anti_join(position_start, position_end, by = "JOB NUMBER") %>%
  mutate(`Change Single` =  paste("eliminate XX", `Classification Name`, "position"),
         !!pos_cols$planning$salary := NA,
         !!pos_cols$planning$opcs := NA,
         !!pos_cols$planning$total_cost := NA,
         `Change Net` = -1) %>%
  rename(!!pos_cols$projection$salary:= Salary,
         !!pos_cols$projection$opcs := OPCs,
         !!pos_cols$projection$total_cost := `TOTAL COST`)

df$service_from <- position_end %>%
  left_join(position_start %>%
              select(`JOB NUMBER`, `Service ID`, `Service Name`, Salary, OPCs, `TOTAL COST`),
            by = "JOB NUMBER",
            suffix = c(paste0(" FY", params$start_yr, " ", params$start_phase), paste0(" FY", params$end_yr, " ", params$end_phase))) %>%
  filter(!!pos_cols$projection$service_id != !!pos_cols$planning$service_id) %>%
  mutate(`Change Single` = paste0("transfer XX ", `Classification Name`, " position from Service ",
                                  !!pos_cols$projection$service_id,
                                  ": ", !!pos_cols$projection$service_name),
         `Change Multiple` = paste0("transfer from Service ", !!pos_cols$projection$service_id,
                                    ": ", !!pos_cols$projection$service_name),
         `Change Net` = 1) %>%
  select(-!!pos_cols$projection$service_id) %>%
  rename(`Service ID` := !!pos_cols$planning$service_id)

df$service_to <- position_start %>%
  left_join(position_end %>%
              select(`JOB NUMBER`, `Service ID`, `Service Name`, Salary, OPCs, `TOTAL COST`),
            by = "JOB NUMBER",
            suffix = c(paste0(" FY", params$start_yr, " ", params$start_phase), paste0(" FY", params$end_yr, " ", params$end_phase)))  %>%
  filter(!!pos_cols$projection$service_id != !!pos_cols$planning$service_id) %>%
  mutate(`Change Single` =
           paste0("transfer XX ", `Classification Name`, " position to Service ",
                  !!pos_cols$planning$service_id, ": ", !!pos_cols$planning$service_name),
         `Change Multiple` =
           paste0("transfer to Service ",
                  !!pos_cols$planning$service_id, ": ", !!pos_cols$planning$service_name),
         `Change Net` = -1) %>%
  select(-!!pos_cols$planning$service_id) %>%
  rename(`Service ID` := !!pos_cols$projection$service_id)

position_start <- position_start %>% rename(`Classification ID` = `CLASSIFICATION ID`)
position_end <- position_end %>% rename(`Classification ID` = `CLASSIFICATION ID`)

df$reclass_to <- position_start %>%
  left_join(position_end %>%
              select(`JOB NUMBER`, `Service ID`, `Classification ID`, `Classification Name`, Salary, OPCs, `TOTAL COST`),
            by = "JOB NUMBER",
            suffix = c(paste0(" FY", params$start_yr, " ", params$start_phase), paste0(" FY", params$end_yr, " ", params$end_phase)))  %>%
  filter(!!pos_cols$projection$class_id !=
           !!pos_cols$planning$class_id) %>%
  mutate(`Change Single` = paste("reclassify XX",
                                 !!pos_cols$projection$class_name,
                                 "position to",
                                 !!pos_cols$planning$class_name),
         `Change Multiple` = paste("reclassify to", !!pos_cols$planning$class_name),
         `Change Net` = 0) %>%
  select(-!!pos_cols$projection$service_id,
         -!!pos_cols$projection$class_id,
         -!!pos_cols$projection$class_name) %>%
  rename(`Service ID` := !!pos_cols$planning$service_id,
         `Classification ID` := !!pos_cols$planning$class_id,
         `Classification Name` := !!pos_cols$planning$class_name)

position_start <- position_start %>% rename(`Fund ID` = `FUND ID`, `Fund Name` = `FUND NAME`)
position_end <- position_end %>% rename(`Fund ID` = `FUND ID`, `Fund Name` = `FUND NAME`)

df$fund <- position_start %>%
  left_join(position_end %>%
              select(`JOB NUMBER`, `Service ID`, `Fund ID`, `Fund Name`, Salary, OPCs, `TOTAL COST`),
            by = "JOB NUMBER",
            suffix = c(paste0(" FY", params$start_yr, " ", params$start_phase), paste0(" FY", params$end_yr, " ", params$end_phase)))  %>%
  filter(!!pos_cols$projection$fund_id != !!pos_cols$planning$fund_id) %>%
  mutate(`Change Single` = 
           case_when(
             !!pos_cols$projection$fund_id == "1001" ~
               paste("transfer XX", `Classification Name`, "position to",
                     !!pos_cols$planning$fund_name, "Fund from General Fund"),
             !!pos_cols$planning$fund_id == "1001" ~
               paste("transfer XX", `Classification Name`, "position from",
                     !!pos_cols$projection$fund_name, "Fund to General Fund"),
             TRUE ~ paste("transfer XX", `Classification Name`, "position from",
                          !!pos_cols$projection$fund_name, "Fund to", !!pos_cols$planning$fund_name, "Fund")),
         `Change Multiple` =
           case_when(
             !!pos_cols$projection$fund_id == "1001" ~
               paste("transfer to", !!pos_cols$planning$fund_name, "Fund from General Fund"),
             !!pos_cols$planning$fund_id == "1001" ~
               paste("transfer from", !!pos_cols$projection$fund_name, "Fund to General Fund"),
             TRUE ~ paste(" transfer from", !!pos_cols$projection$fund_name,
                          "Fund to", !!pos_cols$planning$fund_name, "Fund"))) %>%
  select(-!!pos_cols$projection$service_id,
         -!!pos_cols$projection$fund_id,
         -!!pos_cols$projection$fund_name)  %>%
  rename(`Service ID` := !!pos_cols$planning$service_id,
         `Fund ID` := !!pos_cols$planning$fund_id,
         `Fund Name` := !!pos_cols$planning$fund_name)

position_start <- position_start %>% rename(`Total Cost` = `TOTAL COST`)
position_end <- position_end %>% rename(`Total Cost` = `TOTAL COST`)

df$fund_to <- df$fund %>%
  filter(grepl("transfer to", `Change Multiple`, fixed = TRUE)) %>%
  mutate(`Change Net` = -1)

df$fund_from <- df$fund %>%
  filter(grepl("transfer from", `Change Multiple`, fixed = TRUE) & `Fund Name` == "General") %>%
  mutate(`Change Net` = 1)

df$fund_non_gf <- df$fund %>%
  filter(!grepl("General Fund", `Change Multiple`, fixed = TRUE))

df$fund <- NULL

df <- df %>%
  map(select, starts_with("Agency"), starts_with("Service"), starts_with("Fund"),
      `JOB NUMBER`, starts_with("Classification"), `SI ID`, starts_with("Change"),
      starts_with("salary"), starts_with("OPCs"), starts_with("tOTAL COST")) %>%
  map(mutate_at, vars(matches("ID$|JOB NUMBER")), as.character)

fund_non_gf <- df$fund_non_gf
df$fund_non_gf <- NULL

df <- df %>%
  bind_rows(.id = "Change Type") %>%
  arrange(`Service ID`) %>%
  mutate_if(is.numeric, replace_na, 0) %>%
  mutate(
    `Service Name` = case_when(
      is.na(`Service Name`) & `Change Type` == "service_to" ~
        !!pos_cols$projection$service_name,
      is.na(`Service Name`) & `Change Type` == "service_from" ~
        !!pos_cols$planning$service_name,
      TRUE ~ `Service Name`),
    # Diff for transferred positions should reflect the fact that the entire position
    # moved bw services/fund, not the diff bw the new salary/total cost and old one
    
    # The planning year budget should be used for for position changes b/w services/funds
    # for presentation purposes; adjust salary adjustment line to make everything add up
    !!pos_cols$projection$salary := 
      ifelse(`Change Type` %in% c("service_from", "fund_from"), 0, !!pos_cols$projection$salary),
    !!pos_cols$planning$salary := 
      ifelse(`Change Type` %in% c("service_to", "fund_to"), 0, !!pos_cols$planning$salary),
    !!pos_cols$projection$opcs := 
      ifelse(`Change Type` %in% c("service_from", "fund_from"), 0, !!pos_cols$projection$opcs),
    !!pos_cols$planning$opcs := 
      ifelse(`Change Type` %in% c("service_to", "fund_to"), 0, !!pos_cols$planning$opcs),
    !!pos_cols$projection$total_cost := 
      ifelse(`Change Type` %in% c("service_from", "fund_from"), 0, !!pos_cols$projection$total_cost),
    !!pos_cols$planning$total_cost := 
      ifelse(`Change Type` %in% c("service_to", "fund_to"), 0, !!pos_cols$planning$total_cost),
    `Salary Diff` = !!pos_cols$planning$salary - !!pos_cols$projection$salary,
    `OPCs Diff` = !!pos_cols$planning$opcs - !!pos_cols$projection$opcs,
    `Total Cost Diff` = !!pos_cols$planning$total_cost - !!pos_cols$projection$total_cost,
    `Change Type` = factor(
      `Change Type`, 
      c("create", "service_to", "service_from", "fund_to", "fund_from", "eliminate", "reclass_to", "fund_non_gf"))) %>%
  select(`AGENCY ID`:`SI ID`, starts_with("Change"),
         !!pos_cols$projection$salary, !!pos_cols$planning$salary, `Salary Diff`,
         !!pos_cols$projection$opcs, !!pos_cols$planning$opcs, `OPCs Diff`,
         !!pos_cols$projection$total_cost, !!pos_cols$planning$total_cost, `Total Cost Diff`) %>%
  arrange(`AGENCY NAME`, `Service ID`, `FUND NAME`, `Classification Name`, `Change Type`)

df <- df %>%
  group_by(`AGENCY ID`, `Service ID`, `FUND ID`,`JOB NUMBER`,
           `CLASSIFICATION ID`, `SI ID`) %>%
  mutate(
    `Change Num` = n(),
    `Change Order` = row_number(),
    Change  = ifelse(
      `Change Num` > 1 & `Change Order` > 1, `Change Multiple`, `Change Single`),
    Change = ifelse(
      `Change Num` == 2,
      paste0(Change, collapse = " and "),
      paste0(Change, collapse = ", ")),
    # replace the last comma with ", and"
    Change = stringi::stri_replace_last(Change, fixed = ",", ", and"),
    # capitalize first character in sentence
    Change = paste(toupper(substr(Change, 1, 1)),
                   substr(Change, 2, nchar(Change)), sep = ""),
    # do not repeat "transfer" twice in order to shorten sentences
    # stri_replace_last() will replace the only occurrence of the word and capitalizing
    # stops it from doing this
    Change = ifelse(`Change Num` > 1,
                    stringi::stri_replace_last(Change, fixed = "transfer", ""),
                    Change)) %>%
  select(-`Change Multiple`, -`Change Single`)

change_single <- df %>%
  filter(`Change Num` == 1)

change_multiple <- df %>%
  filter(`Change Num` > 1) %>%
  # arranging by desc because "Transfer" will appear first,
  # thus keeping it when we use distinct()
  arrange(`Service ID`, `JOB NUMBER`, `Change Order`) %>%
  distinct(`Service ID`, `JOB NUMBER`, Change, .keep_all = TRUE)

change_tables$position$detail <- change_single %>%
  bind_rows(change_multiple)  %>%
  ungroup() %>%
  select(-`Change Order`) %>%
  mutate(
    `Change Type` = factor(
      `Change Type`, 
      c("create", "service_to", "service_from", "fund_to", "fund_from", "eliminate", "reclass_to", "fund_non_gf")),
    `Service ID` = as.numeric(`Service ID`)) %>%
  arrange(`AGENCY NAME`, `Service ID`, `Change Type`) %>%
  bind_rows(fund_non_gf %>%
              mutate(`Service ID` = as.numeric(`Service ID`))%>%
              select(-`Change Single`) %>%
              rename(Change = `Change Multiple`) %>%
              mutate(`Change Type` = "fund_non_gf",
                     Change = gsub("transfer", "Transfer", Change)))

change_tables$position$summary <- df %>%
  group_by(`AGENCY ID`, `AGENCY NAME`, `Service ID`, `Service Name`,
           `FUND NAME`, `CLASSIFICATION ID`, `Classification Name`,
           `SI ID`, Change, `Change Num`) %>%
  summarize_if(is.numeric, sum, na.rm = TRUE) %>%
  left_join(df %>%
              group_by(`Service ID`, `FUND NAME`,`CLASSIFICATION ID`,
                       `SI ID`, Change) %>%
              count()) %>%
  rename(`Num of Positions` = n) %>%
  filter(!is.na(`AGENCY NAME`)) %>% # get rid of manually totaled rows
  ungroup() %>%
  select(-(!!sym(paste0("Salary FY", params$start_yr, " ", params$start_phase)):`Total Cost Diff`),
         !!sym(paste0("Salary FY", params$end_yr, " ", params$end_phase)):`Total Cost Diff`) %>%
  mutate(
    Change = ifelse(
      `Num of Positions` == 1,
      str_replace_all(Change, " XX", ""),
      str_replace_all(Change, " XX ", paste0(" ", `Num of Positions`, " "))),
    Change = ifelse(
      `Num of Positions` == 1, Change,
      str_replace_all(Change, " position", " positions")),
    `Service ID` = as.numeric(`Service ID`))

position_summary <- change_tables$position$summary %>%
  mutate(`Object ID` = 1) %>%
  filter(`FUND NAME` == "General") %>%
  group_by(`Service ID`, `Object ID`) %>%
  #force new lines between notes for readability
  mutate(`Position Note` = paste(`Change`, collapse = " \n ")) %>%
  group_by(`Service ID`, `Object ID`, `Position Note`) %>%
  summarise(`Total Cost Diff` = sum(`Total Cost Diff`, na.rm = TRUE), .groups = "drop") %>%
  mutate(`Service ID` = as.numeric(`Service ID`)) %>%
  ungroup()

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
