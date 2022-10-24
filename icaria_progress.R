library(redcapAPI)
library(xlsx)
library(lubridate)
library(stringr)
library(english)
library(echarts4r)

kCRFAZiEvents <- c(
  'epipenta1_v0_recru_arm_1',  # EPI-Penta1 V0 Recruit AZi/Pbo1
  'epimvr1_v4_iptisp4_arm_1',  # EPI-MVR1 V4 IPTi-SP4 AZi/Pbo2
  'epimvr2_v6_iptisp6_arm_1'   # EPI-MVR2 V6 IPTi-SP6 AZi/Pbo3
)
kCRFHHEvent <- c(
  'hhafter_1st_dose_o_arm_1',  # HH-After 1st dose of AZi/Pbo
  'hhafter_2nd_dose_o_arm_1',  # HH-After 2nd dose of AZi/Pbo
  'hhafter_3rd_dose_o_arm_1',  # HH-After 3rd dose of AZi/Pbo
  'hhat_18th_month_of_arm_1'   # HH-At 18th month of age
)
kCRFNonAZiEvents <- c(
  'epipenta2_v1_iptis_arm_1',  # 3  EPI-Penta2 V1 IPTi-SP1
  'epipenta3_v2_iptis_arm_1',  # 4  EPI-Penta3 V2 IPTi-SP2
  'epivita_v3_iptisp3_arm_1',  # 5  EPI-VitA V3 IPTi-SP3
  'epivita_v5_iptisp5_arm_1'   # 8  EPI-VitA V5 IPTi-SP5
)
kCRFEvents <- c(
  kCRFAZiEvents[1],            # EPI-Penta1 V0 Recruit AZi/Pbo1
  kCRFHHEvent[1],              # HH-After 1st dose of AZi/Pbo
  kCRFNonAZiEvents[1],         # EPI-Penta2 V1 IPTi-SP1
  kCRFNonAZiEvents[2],         # EPI-Penta3 V2 IPTi-SP2
  kCRFNonAZiEvents[3],         # EPI-VitA V3 IPTi-SP3
  kCRFAZiEvents[2],            # EPI-MVR1 V4 IPTi-SP4 AZi/Pbo2
  kCRFHHEvent[2],              # HH-After 2nd dose of AZi/Pbo
  kCRFNonAZiEvents[4],         # EPI-VitA V5 IPTi-SP5
  kCRFAZiEvents[3],            # EPI-MVR2 V6 IPTi-SP6 AZi/Pbo3
  kCRFHHEvent[3],              # HH-After 3rd dose of AZi/Pbo
  kCRFHHEvent[4]               # HH-At 18th month of age
)

kEventsDateVars <- c(
  'screening_date',            # EPI-Penta1 V0 Recruit AZi/Pbo1
  'hh_date',                   # HH-After 1st dose of AZi/Pbo
  'int_date',                  # EPI-Penta2 V1 IPTi-SP1
  'int_date',                  # EPI-Penta3 V2 IPTi-SP2
  'int_date',                  # EPI-VitA V3 IPTi-SP3
  'int_date',                  # EPI-MVR1 V4 IPTi-SP4 AZi/Pbo2
  'hh_date',                   # HH-After 2nd dose of AZi/Pbo
  'int_date',                  # EPI-VitA V5 IPTi-SP5
  'int_date',                  # EPI-MVR2 V6 IPTi-SP6 AZi/Pbo3
  'hh_date',                   # HH-After 3rd dose of AZi/Pbo
  'hh_date'                    # HH-At 18th month of age
)

kCOHORTIPTiEvents <- c(
  'ipti_1__10_weeks_r_arm_1',  # IPTi 1 - 10 weeks Recruit
  'ipti_2__14_weeks_arm_1',    # IPTi 2 - 14 weeks
  'ipti_3__9_months_arm_1'     # IPTi 3 - 9 months
)
kCOHORTNonIPTiEvents <- c(
  'mrv_2__15_months_arm_1',    # MRV 2 - 15 months
  'after_mrv_2_arm_1'          # After MRV 2
)

ReadData <- function(api.url, api.token, variables = NULL) {
  #browser()
  rcon <- redcapConnection(api.url, api.token)
  data <- exportRecords(rcon, factors = F, labels = F, fields = variables)
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
    if (hf != "profile" & hf != "cohort") {
      print(paste("Extracting data from", hf))
      
      # TODO: The set of variables to be stracted from REDCap projects should be
      #       predifined in order to improve efficiency
      hf.data <- ReadData(redcap.api.url, redcap.tokens[[hf]])
      if (hf == 'HF02.02') {
        hf_print <- 'HF02'
      } 
      else if (hf =='HF01.01') {
        hf_print <- 'HF01'
      }
      else if (hf =='HF16.01') {
        hf_print <- 'HF16'
      }
      else {
        hf_print <- hf
        
      }
      hf.data <- cbind(hf = hf_print, hf.data)
      data <- rbind(data, hf.data)
    }
  }
  
  # In order to count data by HF (table), we need to encode HF column as factor
  data$hf <- as.factor(data$hf)
  
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

ExportDataCohort <- function(redcap.api.url, redcap.tokens) {
  # Export data from the COHORT ancillary study REDCap project.
  #
  # Args:
  #   redcap.api.url: String representing the URL to access the REDCap API.
  #   redcap.tokens:  List of tokens (Strings) to access each of the ICARIA 
  #                   REDCap projects, among them the COHORT project.
  # 
  # Returns:
  #   Data frame with the Trial Profile data.
  
  print("Extracting data from cohort")
  cohort <- ReadData(redcap.api.url, redcap.tokens[['cohort']])
  
  # Health Facility IDs are scattered in three variables: hf_bombali, 
  # hf_port_loko and hf_tonkolili
  hf.columns <- c("hf_bombali", "hf_port_loko", "hf_tonkolili")
  record.in.hf <- cohort[, c("record_id", hf.columns)]
  record.in.hf$hf <- rowSums(record.in.hf[, hf.columns], na.rm = T)
  record.in.hf <- record.in.hf[which(record.in.hf$hf != 0), ]
  cohort$hf <- lapply(cohort$record_id, function(id) { 
    record.in.hf$hf[which(record.in.hf$record_id == id)] })
  cohort$hf <- as.factor(as.numeric(cohort$hf))
  
  return(cohort)
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

CountNumberOfResponses <- function(data, var, val, event = NULL, by.hf = T) {
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
  #   by.hf: True/False if result must be disaggregated by ICARIA Health 
  #          Facility or not.
  # 
  # Returns:
  #   List of occurences by ICARIA Health Facility.
  
  if (nrow(data) == 0)
    return(0)
  
  if (is.null(event)){
    condition <- which(data[var] == val)
  } else {
    condition <- which(data['redcap_event_name'] == event & data[var] == val)
  }
  
  if (by.hf) {
    col <- table(data[condition, c('hf', var)])
  } else {
    col <- table(data[condition, var])
  }
  
  if (length(col) == 0)
    return(0)  
  
  return(col)
}

GetMigrations <- function(data) {
  # Compute and returns a data frame with one row per participant migration
  # event. A migration can be an OUT migration, when the participant is leaving
  # a health facility catchment area or an IN migration, when the participant
  # is coming to a different catchment area in which s/he was recruited. If a
  # participant is moving form one ICARIA health facility to another, in this
  # case s/he will have two rows here, one OUT migration and one IN migration.
  # This data frame is composed by the following columns:
  #   (1) hf: String representing the ICARIA HF code.
  #   (2) record_id: Integer representing the REDCap record id of the migrated
  #                  participant.
  #   (3) mig_date:  Date of the migration.
  #   (4) origin:      Integer representing the ID of the origin HF.
  #   (5) destination: Integer representing the ID of the destination HF.
  #   (6) in_mig:      Boolean indicating whether or not this is an IN 
  #                    migration.
  #   (7) azi2_date:   Date of the second (middle) AZi/Pbo dose.
  #
  # Args:
  #   data:    Data frame containing ALL Health Facility data sets extracted 
  #            from the ICARIA REDCap projects.
  # 
  # Returns:
  #   Data frame with one row per migration.
  
  # Prepare list of variables to be extracted from the project data frame
  origin.prefix <- "mig_origin_hf_"
  destination.prefix <- "mig_destination_hf_"
  districts <- c("bombali", "port_loko", "tonkolili")
  origin.columns <- paste0(origin.prefix, districts)
  destination.columns <- paste0(destination.prefix, districts)
  mig.columns <- c("hf", "record_id", "mig_reported_date", origin.columns, 
                   destination.columns)
  
  # Extract migrations variables from the project data frame
  migrations <- data[which(data$migration_complete == 2), mig.columns]
  
  # Collapse the origin and destination of the migration in just two columns
  # independently of the district
  migrations$origin <- rowSums(migrations[, origin.columns], na.rm = T)
  migrations$destination <- rowSums(migrations[, destination.columns], 
                                    na.rm = T)
  
  # Compute if each migration is an IN or OUT migration 
  migrations$in_mig <- 
    as.integer(substring(migrations$hf, 3)) == migrations$destination
  
  # Include the dates all the events after recruitment toknow when the 
  # migration occurred (between which two events). 
  # TODO: There's no 2nd doses yet. Needs to be tested! (20210721)
  merge.columns <- c("hf", "record_id")
  for (i in 1:length(kCRFEvents)) {
    event.date <- data[
      which(data$redcap_event_name == kCRFEvents[i]),
      c(merge.columns, kEventsDateVars[i])
    ]
    
    date.column <- kCRFEvents[i]
    colnames(event.date)[3] <- date.column
    
    event.date <- event.date[which(!is.na(event.date[date.column])), ]
    
    migrations <- merge(
      x     = migrations,
      y     = event.date,
      by    = merge.columns,
      all.x = T
    )
  }
  
  # Filter non-relevant columns
  relevant.cols <- c("hf", "record_id", "mig_reported_date", "origin", 
                     "destination", "in_mig", kCRFEvents)
  migrations <- migrations[, relevant.cols]
  
  return(migrations)
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
    for (event in kCRFAZiEvents) {
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
 
  # Apply migrations to summary
  migrations <- GetMigrations(data)

  # We have always to substract the number of IN Migrations to the following 
  # summary data frame columns: screening_consent_1 (ICF Signed) and eligible_1 
  # (Randomized)
  summary$in_mig <- as.vector(table(migrations$hf[migrations$in_mig]))
  summary$screening_consent_1 <- summary$screening_consent_1 - summary$in_mig
  summary$eligible_1 <- summary$eligible_1 - summary$in_mig
  
  # Check when the IN migrations occurred and substract in the AZi/Pbo events 
  # where the data collection was done in the previous health facility. I.e. 
  # substract when migration date is after event date
  for (azi.event in kCRFAZiEvents) {
    # Compute the number of IN migrations per HF that occurred after the event 
    # date
    filter <- migrations$in_mig & 
      migrations$mig_reported_date > migrations[[azi.event]]
    mig.after.event <- migrations[which(filter), ]
    summary$in_mig <- 
      as.vector(table(mig.after.event$hf[mig.after.event$in_mig]))
    
    # Substract this number as this activity occurred in the previous HF
    azi.col.name <- paste0(azi.event, "_int_azi_1")
    summary[azi.col.name] <- summary[[azi.col.name]] - summary$in_mig
  }
  
  # Reorder and rename columns to rescpect the progress report table design
  summary <- summary[, ordered.vars]
  colnames(summary) <- var.names
  
  return(summary)
}

SummarizeCohortData <- function(hf.list, data) {
  # Compute and returns the sum of the recruitment progress variables by health 
  # facility in the COHORT ancillary study. These variables are:
  #   (1) n_recruited:  Number of children going through the informed consent
  #                     process in which their caretakers do sign the informed 
  #                     content form.
  #   (2) n_ipti1:      Number of children who took the 1st dose of IPTi.
  #   (3) n_ipti2:      Number of children who took the 2nd dose of IPTi.
  #   (4) n_ipti3:      Number of children who took the 3rd dose of IPTi.
  #   (5) n_mrv2:       Number of children who took the 2nd dose of MRV.
  #   (6) n_after_mrv2: Number of children who was visited for the end of follow
  #                     up visit.
  #
  # Args:
  #   hf.list: List of ICARIA health facilities IDs (integers) to be summarized.
  #   data:    Data frame containing COHORT study data.
  # 
  # Returns:
  #   Data frame with one row per health facility and one column per variable to
  #   be summarized.
  
  # Variable names
  var.names <- c('hf.list', 'n_recruited', 'n_ipti1', 'n_ipti2', 'n_ipti3', 
                 'n_mrv2', 'n_after_mrv2')
  
  # Summarize all COHORT variables by health facility
  summary <- data.frame(hf.list)
  
  # Get number of recruited participants. In this case, unlike the TRIAL 
  # projects, we don't have the eligible variable. So we have to check how many
  # study number do we have
  summary['n_recruited'] <- table(data$hf[which(!is.na(data$study_number))])
  
  # Summarize IPTi events
  for (event in kCOHORTIPTiEvents) {
    # The Vaccination Status DCI is included in all EPI visits in which IPTi is
    # administered according to the Sierra Leona Immunization Program, although 
    # it is not administered in all visits (I don't know why - missing U5 card?)
    var <- 'intervention_complete'
    summary[event] <- CountNumberOfResponses(data, var, 2, event)  
  }
  # Summarize NON IPTi events
  for (event in kCOHORTNonIPTiEvents) {
    # The Clinical History is included in all EPI visits in which IPTi is not
    # administered.
    var <- 'clinical_history_complete'
    summary[event] <- CountNumberOfResponses(data, var, 2, event)  
  }
  
  # Convert all columns to the same type: numeric
  summary$hf.list <- as.character(summary$hf.list)  # Remove factors
  summary <- as.data.frame(lapply(summary, as.numeric))
  
  # Rename columns to respect the progress report table design
  colnames(summary) <- var.names
  
  return(summary)
}

NextWeekDay <- function(date, week.day) {
  # Compute the date of the requested next week day since the provided date.
  #
  # Args:
  #   date:     Date from which the date of the next day of the week will be 
  #             calculated.
  #   week.day: Integer representing the next day of the week (1 = Sunday, 
  #             2 = Monday, 3 = Tuesday, 4 = Wednesday, 5 = Thursday, 6 = Friday 
  #             and 7 = Saturday)
  # 
  # Returns:
  #   Date of the requested next week day since the provided date.
  
  date <- as.Date(date)
  diff <- week.day - wday(date)
  
  if( diff < 0 ) {
    diff <- diff + 7
  }
  return(date + diff)
}

GetHealthFacilityTimeSeries <- function(hf.id, hf.data, report.date, 
                                        precision = "w", profile = NULL) {
  # Compute and returns the status of the recruitment progress variables in a 
  # time series fashion for concrete ICARIA Health Facility. These variables 
  # are:
  #   (1)  n_random:     Number of randomized participants
  #   (2)  n_in_mig:     Number of in-migrated participants
  #   (3)  n_out_mig:    Number of out-migrated participants
  #   (4)  n_azi1:       Number of 1st AZi/Pbo doses already administered
  #   (5)  n_hh1:        Number of household visits already performed for AZi1
  #   (6)  n_penta2:     Number of participants who came for Penta2
  #   (7)  n_penta3:     Number of participants who came for Penta3
  #   (8)  n_vit_a1:     Number of participants who came for 1st Vitamin A dose
  #   (9)  n_azi2:       Number of 2nd AZi/Pbo doses already administered
  #   (10) n_hh2:        Number of household visits already performed for AZi2
  #   (11) n_vit_a2:     Number of participants who came for 2nd Vitamin A dose
  #   (12) n_azi3:       Number of 3rd AZi/Pbo doses already administered
  #   (13) n_hh3:        Number of household visits already performed for AZi3
  #   (14) n_end_fu:     Number of participants who completed follow up
  #   (15) n_wdw:        Number of total study withdrawals
  #   (16) n_deaths:     Number of participants who die
  #
  # Args:
  #   hf.id:       Integer representing the ID of the ICARIA Health Facility.
  #   hf.data:     Data frame containing the Health Facility data set extracted 
  #                from the corresponding ICARIA REDCap project.
  #   report.date: Date until the time series has to be reported.
  #   precision:   Character indicating the precision of the time series 
  #                (d = daily, w = weekly, m = monthy, y = yearly)
  # 
  # Returns:
  #   Data frame with one row time point (depending on precision) and one column 
  #   per variable.
  
  # TODO: Implement precision feature. Right now, only weekly prescion is
  #       coded.
  
  time.series <- data.frame()
  
  # Ordered variables to be visualized in the progress report
  profile.vars <- c('n_penta1', 'n_underweight', 'n_over_age', 'n_consent')
  ordered.vars <- c(
    c('date'), 
    profile.vars, 
    c('n_random', 'n_in_mig', 'n_out_mig'), 
    kCRFEvents, 
    c('n_wdw', 'n_deaths')
  )
  
  # Variable names
  var.names <- c('date', 'n_penta1', 'n_underweight', 'n_over_age', 'n_consent', 
                 'n_random', 'n_in_mig', 'n_out_mig', 'n_azi1', 'n_hh1',
                 'n_penta2', 'n_penta3', 'n_vit_a1', 'n_azi2', 'n_hh2', 
                 'n_vit_a2', 'n_azi3', 'n_hh3', 'n_end_fu', 'n_wdw', 'n_deaths')
  
  kWeekDay <- 2 # Monday 00:00 after starting
  start.date <- min(hf.data$screening_date, na.rm = T)
  time.point <- NextWeekDay(start.date, kWeekDay) 
  
  # Collapse HF IDs in the hf column no matter in which district the HF is and
  # keep records of the current HF
  if (!is.null(profile)) {
    profile$hf <- rowSums(
      x     = profile[, c("hf_bombali", "hf_tonkolili", "hf_port_loko")], 
      na.rm = T
    )
    
    profile <- profile[which(profile$hf == hf.id), ]
  }
  
  while (time.point <= report.date) {
    point <- list()
    
    point['date'] <- as.character.Date(time.point)
    
    # If we have profile data, compute pre-screening & screening indicators
    if (!is.null(profile)) {
      for (var in profile.vars)
      point[var] = sum(
        profile[which(profile$screening_date < time.point), c(var)], 
        na.rm = T
      )
    }
    
    # Randomized participants
    point['n_random'] <- CountNumberOfResponses(
      data  =  hf.data[which(hf.data$screening_date < time.point), ], 
      var   = "eligible", 
      val   = 1, 
      by.hf = F
    )
    
    # AZi/Pbo doses
    for (azi.event in kCRFAZiEvents) {
      point[azi.event] <- CountNumberOfResponses(
        data  =  hf.data[which(hf.data$int_date < time.point), ], 
        var   = "int_azi", 
        val   = 1, 
        event = azi.event, 
        by.hf = F
      )
    }
    
    # Household post-AZi/Pbo supervision visits
    for (hh.event in kCRFHHEvent) {
      point[hh.event] <- CountNumberOfResponses(
        data  =  hf.data[which(hf.data$hh_date < time.point), ],
        var   = "hh_child_seen",
        val   = 1,
        event = hh.event,
        by.hf = F
      )
    }
    
    # Non-AZi/Pbo EPI visits
    for (epi.visit in kCRFNonAZiEvents) {
      point[epi.visit] <- CountNumberOfResponses(
        data = hf.data[which(hf.data$int_date < time.point), ],
        var  = "intervention_complete",
        val  = 2,
        event = epi.visit,
        by.hf = F
      )
    }
    
    # Withdrawals
    point['n_wdw'] <- CountNumberOfResponses(
      data = hf.data[which(hf.data$wdrawal_date < time.point), ],
      var  = "withdrawal_complete",
      val  = 2,
      by.hf = F
    )
    
    # Deaths
    point['n_deaths'] <- CountNumberOfResponses(
      data = hf.data[which(hf.data$death_date < time.point), ],
      var  = "death_complete",
      val  = 2,
      by.hf = F
    )
    
    # Out Migrations - Migrations in which the origin is this health facility
    mig_origin_cols <- c('mig_origin_hf_bombali', 
                         'mig_origin_hf_port_loko', 
                         'mig_origin_hf_tonkolili')
    hf.data$mig_origin_hf <- rowSums(hf.data[, mig_origin_cols], na.rm = T)
    point['n_out_mig'] <- CountNumberOfResponses(
      data = hf.data[which(hf.data$mig_reported_date < time.point & 
                             hf.data$mig_origin_hf == hf.id), ],
      var  = "migration_complete",
      val  = 2,
      by.hf = F
    )
    
    # In Migrations - Migrations in which the destination is this health 
    # facility
    mig_destination_cols <- c('mig_destination_hf_bombali', 
                              'mig_destination_hf_port_loko', 
                              'mig_destination_hf_tonkolili')
    hf.data$mig_destination_hf <- rowSums(hf.data[, mig_destination_cols], 
                                          na.rm = T)
    point['n_in_mig'] <- CountNumberOfResponses(
      data = hf.data[which(hf.data$mig_reported_date < time.point & 
                             hf.data$mig_destination_hf == hf.id), ],
      var  = "migration_complete",
      val  = 2,
      by.hf = F
    )
    
    # Apply migrations to the series point
    migrations <- GetMigrations(hf.data)
    if (nrow(migrations) > 0) {
      # We have always to substract the number of IN Migrations to the indicator
      # n_random (Randomized)
      point$n_random <- point$n_random - point$n_in_mig
      
      # Check when the IN migration occurred and substract in the events where
      # the data collection was done in the previous health facility. I.e. 
      # substract when migration date is after event date
      for (event in kCRFEvents) {
        # Compute the number of IN migrations that occurred after the event date
        filter <- migrations$mig_reported_date < time.point & 
          migrations$in_mig & migrations$mig_reported_date > migrations[[event]]
        migrations.number <- nrow(migrations[which(filter), ])
        
        # Substract this number as this activity occurred in the previous HF
        point[event] <- point[[event]] - migrations.number
      }
    }
    
    time.series <- rbind(time.series, point, stringsAsFactors = F)
    time.point <- NextWeekDay(time.point + 1, kWeekDay)
  }
  
  # Reorder and rename columns to rescpect the progress report table design
  time.series <- time.series[, ordered.vars]
  colnames(time.series) <- var.names
  
  return(time.series)
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

CreateExcelReport <- function(metadata, general.progress, cohort.progress, 
                              time.series, health.facilities) {
  
  # Positions in spreadsheet
  kStartColumn <- 1
  kStartRow    <- 3
  
  # Positions in data frame
  kICFLogColumn          <- 6
  kICFeCRFColumn         <- 7
  kNumberOfLogIndicators <- 6
  
  # Sizes
  kNarrowColumn <- 9
  kNormalColumn <- 11
  kWideColumn   <- 36
  kLargeColumn  <- 70
  
  kWideRow      <- 30
  
  kSmallFont    <- 10
  kTinyFont     <- 9
  
  # Colors
  kHEXDarkBlue       <- "#44546A"
  kHEXBlueGrayLight  <- "#E6EEFA"
  kHEXBlueGrayLight2 <- "#ACB9CA"
  kHEXBlueGrayDark   <- "#D6E2F6"
  kHEXBlueGrayDark2  <- "#8497B0"
  kHEXGray           <- "#E7E6E6"
  kHEXYellowLight    <- "#FFF2CC"
  kHEXYellowDark     <- "#BF8F00"
  kHEXRed            <- "#FC0000"
  kINDWhite          <- 9
  
  # Set districts and health facility names
  general.progress$hf.list <- NULL
  rownames(general.progress) <- paste(
    health.facilities$district[health.facilities$trial], 
    health.facilities$code[health.facilities$trial], 
    health.facilities$name[health.facilities$trial]
  )
  
  cohort.progress$hf.list <- NULL
  rownames(cohort.progress) <- paste(
    health.facilities$district[health.facilities$cohort], 
    health.facilities$code[health.facilities$cohort], 
    health.facilities$name[health.facilities$cohort]
  )
  
  # Super header names
  kSuperHeaders <- c("Source: Screening Log", "Source: eCRF")
  
  # Set column names
  colnames(general.progress) <- c("Penta1", "Approached", "Underweight", 
                                  "Over Age", "Refusals", "ICF Signed", 
                                  "ICF Signed", "Screening Failures", 
                                  "More than 10w", "No Penta1", "Less than 4kg",
                                  "Catchment Area", "Other Study", "Allergy", 
                                  "Disease", "Illness", "Randomized", 
                                  "AZi/Pbo1", "AZi/Pbo2", "AZi/Pbo3", 
                                  "Non-Compliants", "Withdrawals", 
                                  "Parent Request", "Investigator", "Migration",
                                  "Other", "Deaths")
  
  colnames(cohort.progress) <- c("Recruited", "IPTi1", "IPTi2", "IPTi3", "MRV2",
                                 "After MRV2")
  
  # Create the totals rows
  total.string <- "Total"
  general.progress <- rbind(general.progress, colSums(general.progress))
  general.last.row <- nrow(general.progress)
  rownames(general.progress)[general.last.row] <- total.string
  
  cohort.progress <- rbind(cohort.progress, colSums(cohort.progress))
  cohort.last.row <- nrow(cohort.progress)
  rownames(cohort.progress)[cohort.last.row] <- total.string
  
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
  center.align <- Alignment(
    h        = "ALIGN_CENTER"
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
  
  super.header.background.1 <- Fill(
    backgroundColor = kHEXBlueGrayDark2, 
    foregroundColor = kHEXBlueGrayDark2,
    pattern         = "SOLID_FOREGROUND"
  )
  super.header.background.2 <- Fill(
    backgroundColor = kHEXBlueGrayLight2, 
    foregroundColor = kHEXBlueGrayLight2,
    pattern         = "SOLID_FOREGROUND"
  )
  super.header.font <- Font(
    wb             = wb, 
    color          = kINDWhite, 
    isBold         = T,
    isItalic       = T, 
    heightInPoints = kSmallFont
  )
  super.header.1 <- CellStyle(
    wb        = wb,
    alignment = center.align,
    fill      = super.header.background.1,
    font      = super.header.font
  )
  super.header.2 <- CellStyle(
    wb        = wb,
    alignment = center.align,
    fill      = super.header.background.2,
    font      = super.header.font
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
  
  comments.font <- Font(
    wb             = wb, 
    heightInPoints = kTinyFont
  )
  table.comments <- CellStyle(
    wb        = wb,
    font      = comments.font
  )
  
  highlight.background <- Fill(
    backgroundColor = kHEXYellowLight, 
    foregroundColor = kHEXYellowLight,
    pattern         = "SOLID_FOREGROUND"
  )
  highlight.header.background <- Fill(
    backgroundColor = kHEXYellowDark, 
    foregroundColor = kHEXYellowDark,
    pattern         = "SOLID_FOREGROUND"
  )
  highlight.font <- Font(
    wb       = wb,
    isItalic = T
  )
  highlight <- CellStyle(
    wb   = wb,
    font = highlight.font
  )
  
  time.points <- CellStyle(
    wb   = wb,
    fill = totals.background
  )
  
  coherence.font <- Font(
    wb    = wb,
    color = kHEXRed
  )
  coherence <- CellStyle(
    wb   = wb,
    font = coherence.font
  )
  
  # Create first Excel sheet called META containing the report meta-data
  meta.sheet <- createSheet(wb, "META")
  setColumnWidth(meta.sheet, 1, kWideColumn)     # Meta-data key
  setColumnWidth(meta.sheet, 2, kLargeColumn)    # Meta-data value
  
  meta.df <- data.frame(row.names = str_to_title(names(metadata)))
  meta.df$value <- metadata
  
  addDataFrame(
    x             = meta.df,
    sheet         = meta.sheet,
    startColumn   = 1,
    startRow      = 1,
    rownamesStyle = table.header + left.align,
    col.names     = F,
    colStyle = list('1' = table.comments)
  )
  
  row <- getRows(meta.sheet, 7)
  cells <- getCells(row, 2)
  addHyperlink(cells[[1]], meta.df$value[[7]], "URL", table.comments)
  
  # Add report date
  #row <- createRow(trial.sheet, rowIndex = 1)
  #cell <- createCell(row, colIndex = 1)[[1]]
  #date <- format(report.date, format = "%Y%m%d %H:%M")
  #setCellValue(cell, paste("Report Date:", date))
  
  # Create second Excel sheet called TRIAL containing the ICARIA TRIAL general 
  #progress by Health Facility
  trial.sheet <- createSheet(wb, "TRIAL")
  
  # Set columns widths
  num.cols <- ncol(general.progress) + 1
  setColumnWidth(trial.sheet, 1, kWideColumn)              # District + HF
  setColumnWidth(trial.sheet, 2:num.cols, kNormalColumn)   # Indicators
  setColumnWidth(trial.sheet, num.cols + 1, kLargeColumn)  # Comments
  
  # Combine cells for the two table super-headers which indicate the source
  # where each indicator comes from and create them
  row <- createRow(trial.sheet, kStartRow - 1)
  log.super.header <- createCell(
    row      = row, 
    colIndex = kStartColumn + 1
  )
  crf.super.header <- createCell(
    row      = row, 
    colIndex = kStartColumn + kNumberOfLogIndicators + 1
  )
  setCellValue(log.super.header[[1]], kSuperHeaders[1])
  setCellValue(crf.super.header[[1]], kSuperHeaders[2])
  setCellStyle(log.super.header[[1]], super.header.1)
  setCellStyle(crf.super.header[[1]], super.header.2)
  
  addMergedRegion(  # Screening log indicators
    sheet       = trial.sheet, 
    startRow    = kStartRow - 1, 
    endRow      = kStartRow - 1,
    startColumn = kStartColumn + 1,
    endColumn   = kStartColumn + kNumberOfLogIndicators
  )
  addMergedRegion(  # eCRF indicators
    sheet       = trial.sheet, 
    startRow    = kStartRow - 1, 
    endRow      = kStartRow - 1,
    startColumn = kStartColumn + kNumberOfLogIndicators + 1,
    endColumn   = kStartColumn + ncol(general.progress)
  )
  
  # Add ICARIA TRIAL general progress table
  addDataFrame(
    x             = general.progress[-general.last.row, ], 
    sheet         = trial.sheet, 
    startColumn   = kStartColumn,
    startRow      = kStartRow,
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
  
  # Set table header row height
  table.header.row <- getRows(trial.sheet, kStartRow)
  setRowHeight(table.header.row, kWideRow)
  
  # Add Totals row at the bottom of the table as an independent data frame to
  # control de style
  col.style <- rep(list(table.totals), ncol(general.progress))
  names(col.style) <- seq(1, ncol(general.progress), by = 1)
  
  addDataFrame(
    x = general.progress[general.last.row, ],
    sheet = trial.sheet,
    startColumn = kStartColumn,
    startRow = kStartRow + nrow(general.progress),
    col.names = F,
    rownamesStyle = table.header,
    colStyle = col.style
  )
  
  # Add the Comments column at the end of the table as an independed data frame
  comments <- list(
    "Comments" = health.facilities$comments_trial[health.facilities$trial])
  addDataFrame(
    x = comments,
    sheet = trial.sheet,
    startColumn = ncol(general.progress) + 2,
    startRow = kStartRow,
    row.names = F,
    colnamesStyle = table.header + left.align,
    colStyle = list('1' = table.comments)
  )
  
  # Highlight ICF cells when screening log and eCRF does not match
  for (i in 1:nrow(general.progress[-general.last.row, ])) {
    if (general.progress[i, kICFLogColumn] != 
        general.progress[i, kICFeCRFColumn]) {
      row <- getRows(trial.sheet, i + kStartRow)
      
      # Values
      cells <- getCells(row, kStartColumn + c(kICFLogColumn, kICFeCRFColumn))
      for (cell in cells) {
        setCellStyle(cell, coherence) 
      }
      
    }
  }
  
  # Create third Excel sheet called COHORT containing the COHORT general 
  # progress by Health Facility
  cohort.sheet <- createSheet(wb, "COHORT")
  
  # Set columns widths
  num.cols <- ncol(cohort.progress) + 1
  setColumnWidth(cohort.sheet, 1, kWideColumn)             # District + HF
  setColumnWidth(cohort.sheet, 2:num.cols, kNormalColumn)  # Indicators
  setColumnWidth(cohort.sheet, num.cols + 1, kLargeColumn) # Comments
  
  # Add COHORT general progress table
  addDataFrame(
    x             = cohort.progress[-cohort.last.row, ], 
    sheet         = cohort.sheet, 
    startColumn   = kStartColumn,
    startRow      = kStartRow,
    colnamesStyle = table.header,
    rownamesStyle = table.header + left.align
  )
  
  # Set table header row height
  table.header.row <- getRows(cohort.sheet, kStartRow)
  setRowHeight(table.header.row, kWideRow)
  
  # Add Totals row at the bottom of the table as an independent data frame to
  # control de style
  col.style <- rep(list(table.totals), ncol(cohort.progress))
  names(col.style) <- seq(1, ncol(cohort.progress), by = 1)
  
  addDataFrame(
    x = cohort.progress[cohort.last.row, ],
    sheet = cohort.sheet,
    startColumn = kStartColumn,
    startRow = kStartRow + nrow(cohort.progress),
    col.names = F,
    rownamesStyle = table.header,
    colStyle = col.style
  )
  
  # Add the Comments column at the end of the table as an independed data frame
  comments <- list(
    "Comments" = health.facilities$comments_cohort[health.facilities$cohort])
  addDataFrame(
    x = comments,
    sheet = cohort.sheet,
    startColumn = ncol(cohort.progress) + 2,
    startRow = kStartRow,
    row.names = F,
    colnamesStyle = table.header + left.align,
    colStyle = list('1' = table.comments)
  )
  
  # Highlight rows
  cohort.health.facilities <- health.facilities[health.facilities$cohort, ]
  for (i in 1:nrow(cohort.health.facilities)) {
    if (cohort.health.facilities$cohort_highlighted[i]) {
      row <- getRows(cohort.sheet, i + kStartRow)
      
      # Header
      cells <- getCells(row, kStartColumn)
      setCellStyle(cells[[1]], highlight + highlight.header.background + 
                     header.font)
      
      # Values
      cells <- getCells(row, (kStartColumn + 1):(ncol(cohort.progress) + 1))
      for (cell in cells) {
        setCellStyle(cell, highlight + highlight.background) 
      }
      
      # Comments
      cells <- getCells(row, ncol(cohort.progress) + 2)
      setCellStyle(cells[[1]], highlight + highlight.background + comments.font)
    }
  }
  
  # Create one Excel sheet per Health Facility containing the progress in time
  for (hf.id in names(time.series)) {
    # Set row names as date points
    series <- time.series[[hf.id]]
    row.names(series) <- paste("Week", 1:nrow(series), "-", "Monday", series$date, "00:00:00")
    series$date <- NULL
    
    # Set column names
    colnames(series) <- c("Penta1", "Underweight", "Over Age", "ICF Signed", 
                          "Randomized", "IN Migrated", "OUT Migrated", 
                          "Penta1-AZi/Pbo", "HH-Visit1", "Penta2", "Penta3", 
                          "VitA", "MVR1-AZi/Pbo", "HH-Visit2", "VitA", 
                          "MVR2-AZi/Pbo", "HH-Visit3", "HH-Visit END F/U", 
                          "Withdrawals", "Deaths")
    
    # Create new sheet for the Health Facility
    sheet.name <- as.character(health.facilities[hf.id, 'code'])
    time.series.sheet <- createSheet(wb, sheet.name)
    
    # Set columns widths
    setColumnWidth(time.series.sheet, 1, kWideColumn)      # Time point
    setColumnWidth(time.series.sheet, 2:17, kNormalColumn) # Indicators
    
    # Add Helath Facility details
    row <- createRow(time.series.sheet, rowIndex = 1)
    cell <- createCell(row, colIndex = 1)[[1]]
    hf.details <- paste(health.facilities[hf.id, 'district'], 
                        health.facilities[hf.id, 'code'], 
                        health.facilities[hf.id, 'name'])
    setCellValue(cell, hf.details)
    
    # Add ICARIA Health Facility time progress table
    addDataFrame(
      x             = series, 
      sheet         = time.series.sheet, 
      startColumn   = kStartColumn,
      startRow      = kStartRow,
      colnamesStyle = table.header,
      rownamesStyle = time.points
    )
    
    # Set table header row height
    table.header.row <- getRows(time.series.sheet, kStartRow)
    setRowHeight(table.header.row, kWideRow)
  }
  
  # Save Excel Work Book
  saveWorkbook(wb, metadata$filename)
}

expectedAtDate <- function(date, weekly.rate, start) {
  #browser()
  kWeekDay <- 2 # Monday 00:00 after starting
  expect <- 0
  until <- as.Date(start)
  while (until < date) {
    expect <- expect + weekly.rate
    until <- NextWeekDay(until + 1, kWeekDay)
  }
  
  expect
}

formatReportDate <- function(date) {
  
  date.string <- str_to_title(paste0(
    format(date, '%b'), '. ', 
    as.integer(format(date, '%d')), 
    str_sub(ordinal(as.integer(format(date, '%d'))), -2), ' ',
    format(date, '%Y')
  ))
  
  date.string
}

buildHFProgressBar <- function(series) {
  
  kSize = "435px"
  series |>
    e_chart(date, width = kSize, height = kSize) |>
    e_grid(left = 60, right = 40, top = 10, bottom = 30) |>
    e_color(color = c(kGreen, kOrange)) |>
    e_line(expected, endLabel = list(show = T, color = kGreen)) |>
    e_line(enroled, endLabel = list(show = T, color = kOrange)) |>
    e_tooltip(trigger = 'axis') |>
    e_legend(show = F) -> hf.line.plot
  
  hf.line.plot
}
