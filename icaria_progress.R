library(redcapAPI)

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
  profile.sum <- SummarizeProfileData(hf.list, profile)
  
}

SummarizeProfileData <- function(hf.list, profile) {
  # Compute and returns the sum of the trial profile variables by health 
  # facility. These variables are:
  #    (1) n_penta1:      Number of children vaccinated with Penta1
  #    (2) n_approached:  Number of children vaccinated with Penta1 and 
  #                       approached by the ICARIA nurses
  #    (3) n_underweight: Number of pre-srceened children not selected for 
  #                       screening due to underweight
  #    (4) n_over_age:    Number of pre-screened children not selected for
  #                       screening due to the age
  #    (5) n_refusal:     Number of children going through the informed consent
  #                       process in which their caretakers do not accept to
  #                       sign the informed content form.
  #    (6) n_consent:     Number of children going through the informed consent
  #                       process in which their caretakers do sign the informed 
  #                       content form.
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
  ordered.vars <- c('n_penta1', 'n_approached', 'n_underweight', 'n_over_age', 
                    'n_refusal', 'n_consent')
  
  # Collapse HF IDs in the hf column no matter in which district the HF is
  profile$hf <- rowSums(
    x     = profile[, c("hf_bombali", "hf_tonkolili", "hf_port_loko")], 
    na.rm = T
  )
  
  # Summarized all trial profiles variables by health facility
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