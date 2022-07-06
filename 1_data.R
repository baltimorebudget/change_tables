##read in data =====================
line_item <- readxl::read_excel(path = params$line.item, sheet = "Details") %>%
  rename(`Service ID` = `Program ID`,
         `Service Name` = `Program Name`)

change_tables <- list(
  gen_fund = line_item %>%
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
  group_by(`Agency ID`, `Agency Name`, `Service ID`, `Service Name`, `Object ID`, `Object Name`) %>%
  summarize(!!sym(cols[1]) :=
              sum(!!sym(cols[1]), na.rm = TRUE),
            !!sym(cols[2]) :=
              sum(!!sym(cols[2]), na.rm = TRUE),
            Diff= sum(Diff, na.rm = TRUE), .groups = "drop") %>%
  ungroup()

#positions=================
position_start <- readxl::read_excel(path = params$position.start, sheet = cols[1])
position_end <- readxl::read_excel(path = params$position.end, sheet = cols[2])

pos_cols <- list( # dynamic col names based on FY
  planning = list(
    salary = "Salary FY",
    opcs = "OPCs FY",
    total_cost = "Total Cost FY",
    service_id = "Service ID FY",
    service_name = "Service Name FY",
    class_id = "Classification ID FY",
    class_name = "Classification Name FY",
    fund_id = "Fund ID FY",
    fund_name = "Fund Name FY"))

pos_cols$projection <- pos_cols$planning %>%
  map(function(x) paste0(x, params$fy, " ", params$start_phase)) %>%
  map(sym)

pos_cols$planning <- pos_cols$planning %>%
  map(function(x) paste0(x, params$fy, " ", params$end_phase)) %>%
  map(sym)


clean_pos_files <- function(df) {
  df <- df %>%
  select(-starts_with("OSO")) %>%
    rename(Salary = SALARY, `Service ID` = `PROGRAM ID`, `Service Name` = `PROGRAM NAME`) %>%
    mutate_if(is.numeric , replace_na, replace = 0) %>%
    mutate(OPCs = `TOTAL COST` - Salary,
           # fix the few classification names that are all caps
           `Classification Name` = ifelse(
             toupper(`CLASSIFICATION NAME`) == `CLASSIFICATION NAME`,
             tools::toTitleCase(tolower(`CLASSIFICATION NAME`)),
             `CLASSIFICATION NAME`))
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
            suffix = c(paste0(" FY", params$fy, " ", params$start_phase), paste0(" FY", params$fy, " ", params$end_phase))) %>%
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
            suffix = c(paste0(" FY", params$fy, " ", params$start_phase), paste0(" FY", params$fy, " ", params$end_phase)))  %>%
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
            suffix = c(paste0(" FY", params$fy, " ", params$start_phase), paste0(" FY", params$fy, " ", params$end_phase)))  %>%
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
            suffix = c(paste0(" FY", params$fy, " ", params$start_phase), paste0(" FY", params$fy, " ", params$end_phase)))  %>%
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
      c("create", "service_to", "service_from", "fund_to", "fund_from", "eliminate", "reclass_to", "fund_non_gf"))) %>%
  arrange(`AGENCY NAME`, `Service ID`, `Change Type`) %>%
  bind_rows(fund_non_gf %>%
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
  select(-(!!sym(paste0("Salary FY", params$fy, " ", params$start_phase)):`Total Cost Diff`),
         !!sym(paste0("Salary FY", params$fy, " ", params$end_phase)):`Total Cost Diff`) %>%
  mutate(
    Change = ifelse(
      `Num of Positions` == 1,
      str_replace_all(Change, " XX", ""),
      str_replace_all(Change, " XX ", paste0(" ", `Num of Positions`, " "))),
    Change = ifelse(
      `Num of Positions` == 1, Change,
      str_replace_all(Change, " position", " positions")))

#salaries==================

#nonglobal adjustments
non_global <- import("G:/Budget Publications/automation/0_data_prep/inputs/Change Table Adjustments FY23.xlsx")

non_global_adj <- non_global %>%
  mutate_at(vars(ends_with("ID")), as.character) %>%
  select(-starts_with("Agency"), -`DetailedFund ID`, -`Fund ID`, -`Activity ID`, Tag = `Object ID`,
         -Type , -`One-Time or Recurring?`,
         `Non-Global Adj Amount` = Amount, `Changes or adjustments` = `Change Table Note`) %>%
  rename(`Service ID` = `Program ID`) %>%
  group_by(`Service ID`, Tag) %>%
  mutate(`Non-Global Adj Count` = n(),
         `Non-Global Adj ID` = paste(`Service ID`, `Tag`, sep = "-")) %>%
  ungroup()

flat_adjustments <- non_global %>%
  filter(`Include in Change Table?` == "Yes") %>%
  group_by(`Program ID`, `Change Table Note`) %>%
  summarise(`Amount` = sum(`Amount`, na.rm = TRUE), .groups = "drop") %>%
  ungroup() %>%
  rename(`Service ID` = `Program ID`, `Changes or adjustments` = `Change Table Note`, `Difference` = `Amount`)

#make change table==================================
change_table <- change_tables$obj$summary %>%
  mutate(`Changes or adjustments` = case_when(
    Diff > 0 ~ paste("Increase in", `Object Name`),
    Diff < 0 ~ paste("Decrease in", `Object Name`),
    Diff == 0 ~ paste("No change in", `Object Name`),
    TRUE ~ `Object Name`)#,
    # `Changes or adjustments` = factor(`Changes or adjustments`, c(
    #   # "Salary adjustment",
    #   # "Adjustment for other positional costs",
    #   "Change in active employee health benefit costs",
    #   "Change in pension contributions",
    #   "Adjustment for City fleet rental, repair, and fuel charges",
    #   "Adjustment for City building rental charges",
    #   "Change in allocation for workers' compensation expense",
    #   "Change in cost transfers to capital budget, to/from other funds, and reimbursed expenses",
    #   paste(c("Increase in", "Decrease in", "No change in"), "employee compensation and benefits"),
    #   paste(c("Increase in", "Decrease in", "No change in"), "contractual services expenses"),
    #   paste(c("Increase in", "Decrease in", "No change in"), "operating supplies, equipment, software, and computer hardware"),
    #   paste(c("Increase in", "Decrease in", "No change in"), "grants, contributions, and subsidies"),
    #   paste(c("Increase in", "Decrease in", "No change in"), "all other")), ordered = TRUE)
    ) %>%
  select(`Service ID`, `Object ID`, `Changes or adjustments`, `Diff`) %>%
  group_by(`Service ID`, `Object ID`, `Changes or adjustments`) %>%
  summarise(`Difference` = sum(`Diff`, na.rm = TRUE), .groups = "drop") %>%
  select(-`Object ID`)

#join with flat adjustments
test <- change_table %>% bind_rows(flat_adjustments) %>%
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

for (i in c(642)) {
    budg <- budget %>% filter(`Service ID` == i)
    table <- change_table %>% filter(`Service ID` == i)
  
    
    change_tables$df <- budg %>% filter(`Changes or adjustments` == paste0(cols[1], " Budget")) %>%
      bind_rows(table) %>%
      bind_rows(budg %>% filter(`Changes or adjustments` == paste0(cols[2], " Budget"))) %>%
      mutate(`Amount` = case_when(is.na(`Amount`) ~ `Difference`,
                                  TRUE ~ `Amount`),
             `Notes` = "") %>%
      ungroup()%>%
      select(`Service ID`, `Changes or adjustments`, `Amount`, `Notes`) %>%
      filter(Amount != 0)
    
    diff = filter(change_tables$df, !grepl("Budget", `Changes or adjustments`)) %>% select(Amount) %>% sum()
    minuend = change_tables$df$Amount[change_tables$df$`Changes or adjustments`==paste(cols[2], "Budget")]
    subtrahend = change_tables$df$Amount[change_tables$df$`Changes or adjustments`==paste(cols[1], "Budget")]
    
    if (diff == (minuend - subtrahend)) {print("All good.")}
    else if (diff != (minuend - subtrahend)) {print("Numbers don't add up.")}

return(change_tables$df)
}

#export data ===========================================
export_excel(change_tables$oso$detail, "OSO Detail", paste0(getwd(),"/outputs/FY23 CLS - COU OSO Changes.xlsx"), type = "new")
export_excel(change_tables$oso$summary, "OSO Summary", paste0(getwd(), "/outputs/FY23 CLS - COU OSO Changes.xlsx"), type = "existing")