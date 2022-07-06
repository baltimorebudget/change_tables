#use for function development -------------------------
change_tables <- list(
  gen_fund = line_item %>%
    filter(`Fund ID` == 1001) %>%
    select(`Agency ID`:`Subobject Name`,
           !!sym(cols[[n-(n-1)]]),
           !!sym(cols[[n-(n-2)]]),
           !!sym(cols[[n-(n-3)]]),
           !!sym(cols[[n-(n-4)]]),
           !!sym(cols[[n-(n-5)]]),
           !!sym(cols[[n]]),
           Justification),
  position = list(),
  oso = list(),
  tags = categorize_tags()
)



##connect to db, all schemas, tables available ----------------------------------
db_params <- list(hist_exp_schema = "PLANNINGYEAR",
                  hist_exp_table = "BUDGETS_N_ACTUALS_CLEAN_DF",
                  curr_exp_schema = "planningyear",
                  curr_exp_table = "CURRENT_YEAR_EXPENDITURE",
                  line_item_schema = "planningyear",
                  line_item_table = "LINE_ITEM_REPORT",
                  line_sum_schema = "planningyear",
                  line_sum_table = paste0("LINE_ITEM_SUMMARY_", params$end_phase))

con <- DBI::dbConnect(odbc::odbc(), Driver = "SQL Server", Server = Sys.getenv("DB_BPFS_SERVER"),
                      Database = "Finance_BPFS", UID = Sys.getenv("DB_BPFS_USER"),
                      PWD = Sys.getenv("DB_BPFS_PW"))


##planning year -----------------------------------------------------------------
##planning year line item report details tab
##remember to add cls, proposal, tls, finrec, boe and cou as we proceed through phases!!
##1.Line Item Reports > line_items_yyyy-mm.xlsx
line_detail_query <- DBI::dbGetQuery(conn = con,
                                     statement = "SELECT agency_id AS 'Agency ID', agency_name AS 'Agency Name', program_id AS 'Program ID',
                                     program_name AS 'Program Name', activity_id AS 'Activity ID', activity_name AS 'Activity Name',
                                     subactivity_id AS 'Subactivity ID', subactivity_name AS 'Subactivity Name', fund_id AS 'Fund ID',
                                     fund_name AS 'Fund Name', detailed_fund_id AS 'Detail Fund ID', detailed_fund_name 'Detail Fund Name',
                                     object_id AS 'Object ID', object_name AS 'Object Name', subobject_id AS 'Subobject ID',
                                     subobject_name AS 'Subobject Name', objective_id AS 'Objective ID', objective_name AS 'Objective Name',
                                     current_adopted AS 'FY22 Adopted', cls_request AS 'FY23 CLS', proposal1 AS 'FY23 Proposal',
                                     tls_request AS 'FY23 TLS', finance_recommendation AS 'FY23 FinRec',
                                     change_vs_adopted AS '$ - Change vs Adopted',
                                     FORMAT(perc_change_vs_adopted*100, DECIMAL) AS '% - Change vs Adopted',
                                     justification
                                     FROM [Finance_BPFS].[planningyear23].LINE_ITEM_REPORT
                                     ORDER BY agency_id, program_id, activity_id, subactivity_id, fund_id, detailed_fund_id, object_id, subobject_id, objective_id")

##planning year line item report summary tab
##1.Line Item Reports > line_items_yyyy-mm.xlsx
line_summary_query <- DBI::dbGetQuery(conn = con,
                                      statement = "SELECT AGENCY_ID AS 'Agency ID', AGENCY_NAME as 'Agency Name',
                                      PROGRAM_ID AS 'Program ID', PROGRAM_NAME AS 'Program Name',
                                      FUND_1001 AS 'General Fund', FUND_2000 AS 'Internal Service',
                                      FUND_2024 AS 'Conduit Enterprise', FUND_2070 AS 'Wastewater Utility',
                                      FUND_2071 AS 'Water Utility', FUND_2072 AS 'Stormwater Utility',
                                      FUND_2075 AS 'Parking Enterprise', FUND_2076 AS 'Parking Management',
                                      FUND_4000 AS 'Federal', FUND_4001 AS 'ARPA', FUND_5000 AS 'State',
                                      FUND_6000 AS 'Special Revenue', FUND_7000 AS 'Special Grant',
                                      FUND_1001 + FUND_2000 + FUND_2024 + FUND_2070 + FUND_2071 + FUND_2072 + FUND_2075 + FUND_2076 + FUND_4000 + FUND_4001 + FUND_5000 + FUND_6000 + FUND_7000  AS 'Total'
                                      FROM Finance_BPFS.planningyear23.LINE_ITEM_SUMMARY_TLS
                                      WHERE FUND_1001 + FUND_2000 + FUND_2024 + FUND_2070 + FUND_2071 + FUND_2072 + FUND_2075 + FUND_2076 + FUND_4000 + FUND_4001 + FUND_5000 + FUND_6000 + FUND_7000 > 0
                                      ORDER BY AGENCY_ID, PROGRAM_ID")

##planning year OPC report
##2.Position Reports > PositionsSalariesOpcs_yyy-mm-dd-xlsx
opc_query <- DBI::dbGetQuery(conn = con,
                             statement = "SELECT job_number AS 'JOB NUMBER', class_id AS 'CLASSIFICATION ID', union_code AS 'UNION ID',
                             agency AS 'AGENCY ID', program AS 'PROGRAM ID', activity AS 'ACTIVITY ID', fund AS 'FUND ID',
                             detailed_fund AS 'DETAILED FUND ID', special_indicator AS 'SI ID', adopted_position AS 'ADOPTED',
                             total_cost AS 'SALARY', opc_201 AS 'OSO 201', opc_202 AS 'OSO 202', opc_203 AS 'OSO 203',
                             opc_205 AS 'OSO 205', opc_207 AS 'OSO 207', opc_210 AS 'OSO 210', opc_212 AS 'OSO 212',
                             opc_213 AS 'OSO 213', opc_231 AS 'OSO 231', opc_233 AS 'OSO 233', opc_235 AS 'OSO 235', total_cost AS 'TOTAL COST'
                             FROM [Finance_BPFS].[planningyear23].[all_opcs]")

##planning year OPC report summary tab needed
##2.Position Reports > PositionsSalariesOpcs_yyy-mm-dd-xlsx


##projections year -----------------------------------------------------------------------
##projection year expenditure monthly report detail
##2.Monthly Expenditures > Expenditure yyyy-mm.xlsx
exp_detail_query <- DBI::dbGetQuery(conn = con,
                                    statement = "SELECT a.agency_id AS 'Agency ID', a.agency_name AS 'Agency',
                                    a.program_id AS 'Program ID', a.program_name AS 'Program Name',
                                    a.activity_id AS 'Activity ID', a.activity_name AS 'Activity Name',
                                    a.fund_id AS 'Fund ID', a.fund_name AS 'Fund Name',
                                    a.object_id AS 'Object ID',  a.object_name AS 'Object Name',
                                    a.subobject_id AS 'Subobject ID', a.subobject_name AS 'Subobject Name',
                                    (prior_adopted) AS 'FY21 Adopted',
                                    (prior_actual) AS 'FY21 Actual', (prior_ytd_actual) AS 'FY21 YTD Actual',
                                    (curr_adopted) AS 'FY22 Adopted', (carry_fwd_a) AS 'Carry Forward - A',
                                    (carry_fwd_b) AS 'Carry Forward - B', (carry_fwd_total) AS 'Carry Forward - Total',
                                    (app_adj) AS 'Appropriations Adjustments', (tot_budget) AS 'Total Budget',
                                    (jul) AS 'July', (aug) AS 'August', (sep) AS 'September',
                                    (oct) AS 'October', (nov) AS 'November', (dec) AS 'December',
                                    (jan) AS 'January', (feb) AS 'February', (mar) AS 'March',
                                    (apr) AS 'April', (may) AS 'May', (jun) AS 'June',
                                    (ytd_exp) AS 'BAPS YTD Exp', (tot_enc) AS 'Total Encumbrance', (cur_acc) as 'Current Accrual',
                                    c.TEXT AS 'Carry Forward Purpose',
                                    j.justification AS 'Justification'
                                    FROM [Finance_BPFS].[planningyear23].[CURRENT_YEAR_EXPENDITURE] a
                                    LEFT JOIN (SELECT * FROM [Finance_BPFS].[planningyear23].LINE_ITEM_REPORT) j
                                    ON (j.agency_id = a.agency_id AND j.program_id = a.program_id AND j.activity_id = a.activity_id AND j.fund_id = a.fund_id AND j.subobject_id = a.subobject_id)
                                    LEFT JOIN (SELECT * FROM [Finance_BPFS].[planningyear22].[CARRY_FWD_PURPOSE]) c
                                    ON (a.agency_id = c.AGENCY_ID AND a.program_id = C.PROGRAM_ID AND a.activity_id = C.ACTIVITY_ID AND a.fund_id = C.FUND_ID AND a.object_id = C.OBJECT_ID
                                    AND a.subobject_id = C.SUBOBJECT_ID)
                                    ORDER BY a.agency_id, a.agency_name, a.program_id, a.program_name, a.activity_id, a.activity_name,
                                    a.fund_id, a.fund_name, a.object_id, a.object_name, a.subobject_id, a.subobject_name
                                    ")
## needs fixing
##projection year monthly expenditure report agency pivot table (new, Maggie's request)
##2.Monthly Expenditures > Expenditure yyyy-mm.xlsx
exp_agency_query <- DBI::dbGetQuery(conn = con,
                                    statement = "SELECT agency_id AS 'Agency ID', agency_name AS 'Agency Name',
                                    fund_id AS 'Fund ID', fund_name AS 'Fund Name',
                                    SUM(current_adopted) AS 'Total FY22 Adopted', SUM(cls_request) AS 'Total FY23 CLS',
                                    SUM(proposal1) AS 'Total FY23 Proposal', SUM(tls_request) AS 'Total FY23 TLS'
                                    FROM [Finance_BPFS].[planningyear23].[CURRENT_YEAR_EXPENDITURE]
                                    GROUP BY agency_id, agency_name, fund_id, fund_name
                                    ORDER BY agency_id, fund_id")

##PROJECTION YEAR MONTHLY EXPENDITURE REPORT SUMMARY BY AGENCY TAB
##2.Monthly Expenditures > Expenditure yyyy-mm.xlsx
exp_summ_query <- DBI::dbGetQuery(conn = con,
                                     statement = "SELECT agency_name AS 'Agency', SUM(prior_adopted) AS 'FY21 Adopted',
                                  SUM(prior_actual) AS 'FY21 Actual', SUM(prior_ytd_actual) AS 'FY21 YTD Actual',
                                  SUM(curr_adopted) AS 'FY22 Adopted', SUM(carry_fwd_a) AS 'Carry Forward - A',
                                  SUM(carry_fwd_b) AS 'Carry Forward - B', SUM(carry_fwd_total) AS 'Carry Forward - Total',
                                  SUM(APP_ADJ) AS 'AAO', SUM(TOT_BUDGET) AS 'Total Budget', SUM(YTD_EXP) AS 'YTD Exp',
                                  SUM(tot_enc) AS 'Encumbrances', SUM(total_spent) AS 'Total Spent',
                                  SUM(surp_def) AS 'Surplus/Deficit', SUM(TOT_BUDGET - YTD_EXP) AS 'Net BBMR Sur/Def'
                                  FROM [Finance_BPFS].[planningyear23].[CYE_SUMMARY]
                                  GROUP BY agency_name
                                  ORDER BY agency_name")

