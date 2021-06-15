library(redcapAPI)
library(xlsx)

ReadData <- function(api.url, api.token) {
  #browser()
  rcon <- redcapConnection(api.url, api.token)
  data <- exportRecords(rcon, factors = F, labels = F)
}

ExportDataAllHealthFacilities <- function(redcap.api.url, redcap.tokens) {
  # Export data from each of the ICARIA Health Facility REDCap projects and 
  # append all data sets in a unique data frame for analisys.
  #
  # Args:
  #   redcap.api.url: String representing the URL to access the REDCap API.
  #   redcap.tokens:  List of tokens (Strings) to access each of the ICARIA 
  #                   REDCap projects.
  # 
  # Returns:
  #   Data frame with all the data together of different ICARIA Health 
  #   Facilities.
  
  data <- data.frame()
  for (hf in names(redcap.tokens)) {
    #browser()
    if (hf != "profile") {
      print(paste("Extracting data from", hf))
      hf.data <- ReadData(redcap.api.url, redcap.tokens[[hf]])
      hf.data <- cbind(hf = hf, hf.data)
      data <- rbind(data, hf.data)
    }
  }
  
  return(data)
}

ExportDataTrialProfile <- function(redcap.api.url, redcap.tokens) {
  # Export data from the Trial Profile REDCap project containing the daily 
  # aggregated reports of the Screening Log in the ICARIA Health Facilities.
  #
  # Args:
  #   redcap.api.url: String representing the URL to access the REDCap API.
  #   redcap.tokens:  List of tokens (Strings) to access each of the ICARIA 
  #                   REDCap projects, among them the Trial Profile project.
  # 
  # Returns:
  #   Data frame with the Trial Profile data.
  
  print("Extracting data from profile")
  profile <- ReadData(redcap.api.url, redcap.tokens[['profile']])
  
  return(profile)
}

SummarizeData <- function(hf.list, profile, data) {
  # Compute the data frame to produce the general progress table.
  #
  # Args:
  #   hf.list: List of ICARIA health facilities IDs (integers) to be summarized.
  #   profile: Data frame containing the trial profile data extracted from the 
  #            ICARIA Trial Profile REDCap project.
  #   data:    Data frame containing the CRF data extracted from the ICARIA 
  #            REDCap projects.
  # 
  # Returns:
  #   Data frame with all the indicators by health facility to produce the 
  #   general progress table of the report.
  
  profile.sum <- SummarizeProfileData(hf.list, profile)
  crf.sum <- SummarizeCRFData(hf.list, data)
  
  # Merge both data frames and convert all columns to the same type: numeric
  summary <- data.frame(profile.sum, crf.sum[, -1])
  summary$hf.list <- as.character(summary$hf.list)  # Remove factors
  summary <- as.data.frame(lapply(summary, as.numeric))
  
  return(summary)
}

SummarizeProfileData <- function(hf.list, profile) {
  # Compute and returns the sum of the trial profile variables by health 
  # facility. These variables are:
  #  (1) n_penta1:      Number of children vaccinated with Penta1
  #  (2) n_approached:  Number of children vaccinated with Penta1 and approached 
  #                     by the ICARIA nurses
  #  (3) n_underweight: Number of pre-srceened children not selected for 
  #                     screening due to underweight
  #  (4) n_over_age:    Number of pre-screened children not selected for
  #                     screening due to the age
  #  (5) n_refusal:     Number of children going through the informed consent
  #                     process in which their caretakers do not accept to sign 
  #                     the informed content form.
  #  (6) n_consent:     Number of children going through the informed consent
  #                     process in which their caretakers do sign the informed 
  #                     content form.
  #
  # Args:
  #   hf.list: List of ICARIA health facilities IDs (integers) to be summarized.
  #   profile: Data frame containing the trial profile data extracted from the 
  #            ICARIA Trial Profile REDCap project.
  # 
  # Returns:
  #   Data frame with one row per health facility and one column per variable to
  #   be summarized.
  
  # Trial profile variables to be summarized
  vars <- c('n_penta1', 'n_approached', 'n_underweight', 'n_over_age', 
            'n_consent')
  
  # Ordered variables to be visualized in the progress report
  ordered.vars <- c('hf.list', 'n_penta1', 'n_approached', 'n_underweight', 
                    'n_over_age', 'n_refusal', 'n_consent')
  
  # Collapse HF IDs in the hf column no matter in which district the HF is
  profile$hf <- rowSums(
    x     = profile[, c("hf_bombali", "hf_tonkolili", "hf_port_loko")], 
    na.rm = T
  )
  
  # Summarize all trial profiles variables by health facility
  summary <- data.frame(hf.list)
  for (var in vars) {
    col <- c()
    for (hf in hf.list) {
      col[hf] <- sum(profile[which(profile$hf == hf), var], na.rm = T)
    }
    summary[var] <- col
  }
  
  # Compute the number of refusals based on the captured variables
  summary$n_refusal <- summary$n_approached - summary$n_underweight - 
    summary$n_over_age - summary$n_consent
  
  # Reorder columns to rescpect the progress report table design
  summary <- summary[, ordered.vars]
  
  return(summary)
}

CountNumberOfResponses <- function(data, var, val, event = NULL) {
  # Count number of concrete responses in a concrete variable either generally
  # or by REDCap event and by ICARIA Health Facility.
  #
  # Args:
  #   data:  Data frame containing the CRF data extracted from the ICARIA REDCap 
  #          projects.
  #   var:   Variable name to count.
  #   val:   String represeting the reponse value to count.
  #   event: String representing the REDCap event to filter by when counting 
  #          responses.
  # 
  # Returns:
  #   List of occurences by ICARIA Health Facility.
  if (is.null(event)){
    condition <- which(data[var] == val)
  } else {
    condition <- which(data['redcap_event_name'] == event & data[var] == val)
  }
  col <- table(data[condition, c('hf', var)])
  if (length(col) == 0)
    return(0)  
  
  return(col)
}

SummarizeCRFData <- function(hf.list, data) {
  # Compute and returns the sum of the recruitment progress variables by health 
  # facility. These variables are:
  #   (1) n_consent:     Number of children going through the informed consent
  #                      process in which their caretakers do sign the informed 
  #                      content form.
  #   (2)  n_failures:   Number of total screening failures 
  #   (3)  n_ic_10w:     Number of screening failures due to more than 10 weeks
  #                      inclusion criterion.
  #   (4)  n_ic_penta1:  Number of screening failures due to non-eligibility to
  #                      receive Penta1 inclusion criterion.
  #   (5)  n_ic_weight:  Number of screening failures due to underweight 
  #                      inclusion criterion.
  #   (6)  n_ic_res:     Number of screening failures due to residency inclusion 
  #                      criterion.
  #   (7)  n_ec_study:   Number of screening failures due to other study
  #                      exclusion criterion.
  #   (8)  n_ec_allergy: Number of screening failures due to allergies exclusion 
  #                      criterion.
  #   (9)  n_ec_disease: Number of screening failures due to serious disease
  #                      exclusion criterion.
  #   (10) n_ec_illness: Number of screening failures due to acute illness
  #                      exclusion criterion.
  #   (11) n_random:     Number of randomized participants
  #   (12) n_azi1:       Number of 1st AZi/Pbo doses already administered
  #   (13) n_azi2:       Number of 2nd AZi/Pbo doses already administered
  #   (14) n_azi3:       Number of 3rd AZi/Pbo doses already administered
  #   (15) n_nc:         Number of non-compliant participants
  #   (16) n_wdw:        Number of total study withdrawals
  #   (17) n_wdw_parent: Number of withdrawals by parent request
  #   (18) n_wdw_inv:    Number of withdrawals by investigator request
  #   (19) n_wdw_mig:    Number of withdrawals due to migration
  #   (20) n_wdw_other:  Number of withdrawals due to other reason
  #   (21) n_deaths:     Number of participants who die
  #
  # Args:
  #   hf.list: List of ICARIA health facilities IDs (integers) to be summarized.
  #   data:    Data frame containing ALL Health Facility data sets extracted 
  #            from the ICARIA REDCap projects.
  # 
  # Returns:
  #   Data frame with one row per health facility and one column per variable to
  #   be summarized.
  
  # CRF events
  crf.azi.events <- c('epipenta1_v0_recru_arm_1', 'epimvr1_v4_iptisp4_arm_1',
                      'epimvr2_v6_iptisp6_arm_1')
  
  # CRF variables (equal to 0) to be summarized
  vars.to.0 <- c('ic_age_10w', 'ic_age', 'ic_weight', 'ic_residency_now', 
                 'eligible')
  
  # CRF variables (equal to 1) to be summarized
  vars.to.1 <- c('screening_consent', 'ec_other_study', 'ec_allergies', 
                 'ec_serious_disease', 'ec_acute_illness', 'eligible')
  
  # CRF variables (equal to 1) by AZi event to be summarized
  vars.to.1.event <- c('int_azi')
  
  # Withdrawal reasons
  wdw.reasons <- c(
    1,  # Parent/guardian's request 
    2,  # Investigator's decision 
    3,  # Migration
    88  # Other
  )
  
  # Ordered variables to be visualized in the progress report
  ordered.vars <- c('hf.list', 'screening_consent_1', 'eligible_0', 
                    'ic_age_10w_0', 'ic_age_0', 'ic_weight_0', 
                    'ic_residency_now_0', 'ec_other_study_1', 'ec_allergies_1', 
                    'ec_serious_disease_1', 'ec_acute_illness_1', 'eligible_1',
                    'epipenta1_v0_recru_arm_1_int_azi_1',
                    'epimvr1_v4_iptisp4_arm_1_int_azi_1',
                    'epimvr2_v6_iptisp6_arm_1_int_azi_1', 'nc', 'n_wdw', 
                    'wdrawal_reason_1', 'wdrawal_reason_2', 'wdrawal_reason_3',
                    'wdrawal_reason_88', 'death_complete')
  
  # Variable names
  var.names <- c('hf.list', 'n_consent', 'n_failures', 'n_ic_10w', 
                 'n_ic_penta1', 'n_ic_weight', 'n_ic_res', 'n_ec_study', 
                 'n_ec_allergy', 'n_ec_disease', 'n_ec_illness', 'n_random', 
                 'n_azi1', 'n_azi2', 'n_azi3', 'n_nc', 'n_wdw', 'n_wdw_parent', 
                 'n_wdw_inv', 'n_wdw_mig', 'n_wdw_other', 'n_deaths')
  
  # Summarize all CRF variables by health facility
  summary <- data.frame(hf.list)

  # Summarize CRF variables in which the value of interest is 0
  for (var in vars.to.0) {
    column.name <- paste(var, "0", sep = "_")
    summary[column.name] <- CountNumberOfResponses(data, var, 0)
  }
  
  # Summarize CRF variables in which the value of interest is 1
  for (var in vars.to.1) {
    column.name <- paste(var, "1", sep = "_")
    summary[column.name] <- CountNumberOfResponses(data, var, 1)
  }
  
  # Summarize CRF variables by event in which the value of interest is 1
  for (var in vars.to.1.event) {
    for (event in crf.azi.events) {
      column.name <- paste(event, var, "1", sep = "_")
      summary[column.name] <- CountNumberOfResponses(data, var, 1, event)  
    }
  }
  
  # Summarize non-compliant participants
  nc.column <- 'nc'
  data[nc.column] <- startsWith(data$child_fu_status, 'NC@')
  summary[nc.column] <- CountNumberOfResponses(data, nc.column, T)
  
  # Summarize withdrawal reasons
  for (reason in wdw.reasons) {
    var <- 'wdrawal_reason'
    column.name <- paste(var, reason, sep = "_")
    summary[column.name] <- CountNumberOfResponses(data, var, reason)   
  }
  
  # Summarize deaths
  var <- 'death_complete'
  summary[var] <- CountNumberOfResponses(data, var, 2)
  
  # Aggregated columns: n_wdw
  summary$n_wdw <- summary$wdrawal_reason_1 + summary$wdrawal_reason_2 + 
    summary$wdrawal_reason_3 + summary$wdrawal_reason_88

  # Reorder and rename columns to rescpect the progress report table design
  summary <- summary[, ordered.vars]
  colnames(summary) <- var.names
  
  return(summary)
}

PreVisualizationProcess <- function(df, columns.remove.if.zero) {
  # Process data frame before visualizing it.
  #
  # Args:
  #   columns.remove.if.zero: List of data frame columns to be removed if all
  #                           the values are zero.
  # 
  # Returns:
  #   Data frame processed and ready for viasualization
  
  for (column in columns.remove.if.zero) {
    if (sum(df[column]) == 0) {
      df[column] <- NULL  
    }
  }
  
  return(df)
}

CreateExcelReport <- function(filename, report.date, general.progress, 
                              health.facilities) {
  
  # Sizes
  kNarrowColumn <- 9
  kNormalColumn <- 11
  kWideColumn   <- 36
  kSmallFont    <- 10
  kTinyFont     <- 9
  
  # Colors
  kHEXDarkBlue      <- "#44546A"
  kHEXBlueGrayLight <- "#E6EEFA"
  kHEXBlueGrayDark  <- "#D6E2F6"
  kHEXGray          <- "#E7E6E6"
  kINDWhite         <- 9
  
  # Set districts and health facility names
  general.progress$hf.list <- NULL
  rownames(general.progress) <- paste(health.facilities$district, 
                                      health.facilities$code, 
                                      health.facilities$name)
  
  # Set column names
  colnames(general.progress) <- c("Penta1", "Approached", "Underweight", 
                                  "Over Age", "Refusals", "ICF Signed (LOG)", 
                                  "ICF Signed (CRF)", "Screening Failures", 
                                  "More than 10w", "No Penta1", "Less than 4kg",
                                  "Catchment Area", "Other Study", "Allergy", 
                                  "Disease", "Illness", "Randomized", 
                                  "AZi/Pbo1", "AZi/Pbo2", "AZi/Pbo3", 
                                  "Non-Compliants", "Withdrawals", 
                                  "Parent Request", "Investigator", "Migration",
                                  "Other", "Deaths")
  
  # Create the totals row
  general.progress <- rbind(general.progress, colSums(general.progress))
  last.row <- nrow(general.progress)
  rownames(general.progress)[last.row] <- "Total"
  
  # Process data frame for visualization: (1) Remove failure and withdrawal
  # reason columns if zero
  columns.can.be.hidden <- c("More than 10w", "No Penta1", "Less than 4kg", 
                             "Catchment Area", "Other Study", "Allergy", 
                             "Disease", "Illness", "Parent Request",
                             "Investigator", "Migration", "Other")
  general.progress <- PreVisualizationProcess(general.progress, 
                                              columns.can.be.hidden)
  
  # Create Excel Work Book
  wb <- createWorkbook(type = "xlsx")
  
  # Workbook styles
  right.align <- Alignment(
    h        = "ALIGN_RIGHT", 
    wrapText = T
  )
  left.align <- Alignment(
    h        = "ALIGN_LEFT"
  )
  header.background <- Fill(
    backgroundColor = kHEXDarkBlue, 
    foregroundColor = kHEXDarkBlue,
    pattern         = "SOLID_FOREGROUND"
  )
  header.font <- Font(
    wb             = wb, 
    color          = kINDWhite, 
    isBold         = T,
    heightInPoints = kSmallFont
  )
  table.header <- CellStyle(
    wb        = wb,
    alignment = right.align,
    fill      = header.background,
    font      = header.font
  )
  
  subcolumn.background.main <- Fill(
    backgroundColor = kHEXBlueGrayDark, 
    foregroundColor = kHEXBlueGrayDark,
    pattern         = "SOLID_FOREGROUND"
  )
  subcolumn.background <- Fill(
    backgroundColor = kHEXBlueGrayLight, 
    foregroundColor = kHEXBlueGrayLight,
    pattern         = "SOLID_FOREGROUND"
  )
  subcolumn.font <- Font(
    wb             = wb, 
    heightInPoints = kTinyFont,
    isItalic       = T
  )
  table.subcolumn <- CellStyle(
    wb        = wb,
    fill      = subcolumn.background.main
  )
  
  totals.background <- Fill(
    backgroundColor = kHEXGray, 
    foregroundColor = kHEXGray,
    pattern         = "SOLID_FOREGROUND"
  )
  totals.font <- Font(
    wb     = wb, 
    isBold = T
  )
  table.totals <- CellStyle(
    wb   = wb,
    fill = totals.background,
    font = totals.font
  )
  
  # Create first Excel sheet called Overview containing the ICARIA TRIAL and
  # COHORT general progress by Health Facility
  overview.sheet <- createSheet(wb, "Overview")
  
  # Set columns widths
  setColumnWidth(overview.sheet, 1, kWideColumn)      # District + HF column
  setColumnWidth(overview.sheet, 2:21, kNormalColumn) # Indicators
  
  # Add report date
  row <- createRow(overview.sheet, rowIndex = 1)
  cell <- createCell(row, colIndex = 1)[[1]]
  date <- format(report.date, format = "%Y%m%d %H:%M")
  setCellValue(cell, paste("Report Date:", date))
  
  # Add ICARIA TRIAL general progress table
  addDataFrame(
    x             = general.progress[-last.row, ], 
    sheet         = overview.sheet, 
    startColumn   = 1,
    startRow      = 2,
    colnamesStyle = table.header,
    rownamesStyle = table.header + left.align,
    colStyle      = list(
      '8'  = table.subcolumn,
      '9'  = table.subcolumn + subcolumn.background + subcolumn.font, 
      '10' = table.subcolumn + subcolumn.background + subcolumn.font,
      '11' = table.subcolumn + subcolumn.background + subcolumn.font,
      '17' = table.subcolumn,
      '18' = table.subcolumn + subcolumn.background + subcolumn.font)
  )
  
  # Add Totals row at the bottom of the table as an independent data frame to
  # control de style
  col.style <- rep(list(table.totals), ncol(general.progress))
  names(col.style) <- seq(1, ncol(general.progress), by = 1)
  
  addDataFrame(
    x = general.progress[last.row, ],
    sheet = overview.sheet,
    startColumn = 1,
    startRow = nrow(general.progress) + 2,
    col.names = F,
    rownamesStyle = table.header,
    colStyle = col.style
  )
  
  # Save Excel Work Book
  saveWorkbook(wb, filename)
}