#####################################################################
####            Customized Tables for Feed the Future            ####  
####                  Created by: Greg Haugan                    ####
#####################################################################

###Load packages
packages <- c("dplyr" , "tidyr" , "srvyr", "broom" , "survey" , "openxlsx")
lapply(packages, require, character.only = TRUE)

rm(list=ls())

ftftable <- function(design, outcome, disaggregates , col_var = NULL, decimals=1) {
  
  # Initialize the results data frame with correct column names
  results <- tibble(
    Characteristic = character(),
    Percent = numeric(),
    Sig = character(),
    N = integer()
  )
  
  # Calculate overall mean
  overall_mean <- design %>%
    summarise(mean_outcome = survey_mean(!!sym(outcome), 
                na.rm = TRUE)) %>%
    pull(mean_outcome)
  overall_n <- design %>% summarise(n = sum(!is.na(!!sym(outcome)))) %>% 
    pull(n)
  results <- bind_rows(results, tibble(
    Characteristic = "All",
    Percent = round(overall_mean, decimals),
    Sig = "",
    N = overall_n
  ))
  
  # Loop through each disaggregate variable
  for (disaggregate in disaggregates) {
    
    # Run ANOVA with svyglm and regTermTest - regular ANOVA won't account for survey design
    model <- survey::svyglm(as.formula(paste(outcome, "~", disaggregate)), 
                            design = design)
    anova_result <- regTermTest(model, disaggregate)
    p_value <- anova_result$p
    
    # Determine significance stars
    significance <- case_when(
      p_value < 0.01  ~ "***",
      p_value < 0.05  ~ "**",
      p_value < 0.1   ~ "*",
      TRUE            ~ "N/S" # not significant
    )    
    
    # Add a row with the disaggregate name
    results <- bind_rows(results, tibble(
      Characteristic = disaggregate,
      Percent = NA,
      Sig = significance,
      N = NA
    ))
    
    # Calculate means and Ns for each level of the disaggregate variable
    summary_stats <- design %>%
      group_by(!!sym(disaggregate)) %>%
      summarise(
        Percent = round(survey_mean(!!sym(outcome), na.rm = TRUE),decimals),
        N = sum(!is.na(!!sym(outcome))),
        .groups = "drop"
      ) %>%
      select(-Percent_se)
    
    # Add rows to results
    summary_stats <- summary_stats %>% 
      mutate(
        Characteristic = !!sym(disaggregate),
        Sig = ""
      )  
    summary_stats <- summary_stats[, !(names(summary_stats) %in% disaggregates)]
    
    results <- bind_rows(
      results %>% mutate(Percent = as.character(Percent)),
      summary_stats %>% mutate(Percent = as.character(Percent))
    ) %>%
      mutate( # if N<30, don't report results      
        Percent = ifelse(N < 30, "^", Percent),
      ) 
  }
  
  # If col_var is provided, loop through its levels
  if (!is.null(col_var)) {
    col_levels <- design %>% pull(!!sym(col_var)) %>% unique()
    
    for (col_level in col_levels) {
      design_subset <- design %>% filter(!!sym(col_var) == col_level)
      
      col_results <- ftftable(design_subset, outcome, disaggregates, decimals=decimals)
      col_results <- col_results %>% rename_with(~ paste0(., "_", col_level), 
                                                 -Characteristic)
      
      results <- left_join(results, col_results, by = "Characteristic")
    }
  }
  
  return(results)
}

#Write function to export results into Excel
export_ftftable <- function(table , workbook_name = "Table.xlsx", 
                            sheet_name = "Table") {
  
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add a worksheet
  addWorksheet(wb, paste(sheet_name))
  
  # Write the data to the sheet
  writeData(wb, paste(sheet_name), table)
  
  # Add styling
  header_style <- createStyle(
    fontName = "Gill Sans MT" ,
    fontSize = 12, 
    fontColour = "white", 
    fgFill = "#4F81BD", 
    halign = "CENTER", 
    textDecoration = "Bold"
  )
  
  body_style <- createStyle(
    fontName = "Gill Sans MT" ,
    fontSize = 11,
    halign = "LEFT"
  )
  
  # Apply styles: header for the first row, body for the rest
  addStyle(wb, paste(sheet_name), header_style, rows = 1, cols = 1:ncol(table), 
           gridExpand = TRUE)
  addStyle(wb, paste(sheet_name), body_style, rows = 2:(nrow(table) + 1), 
           cols = 1:ncol(table), gridExpand = TRUE)
  
  # Auto-size the columns for better readability
  setColWidths(wb, paste(sheet_name), cols = 1:ncol(table), widths = "auto")
  
  # Save the Excel file
  saveWorkbook(wb, paste(workbook_name), overwrite = TRUE)
}




# Example data frame
set.seed(123)
df <- data.frame(
  Age = sample(c('15-19', '20-24', '25-29', '30-34', '35-39', '40-44', '45-49'), 
               2000, replace = TRUE),
  Education = factor(sample(c('No Education', 'Primary 1-3', 'Primary 4-8', 
        'Secondary', 'Higher'), 2000, replace = TRUE, prob=c(.2,.45, .2, .1 , .05)),
        levels = c('No Education', 'Primary 1-3', 'Primary 4-8', 'Secondary', 
        'Higher')),
  PregnancyStatus = factor(sample(c('Not Pregnant', 'Pregnant'), 2000, replace = TRUE,
        prob=c(.90,.10)), levels = c('Not Pregnant', 'Pregnant')),
  GenderedHouseholdType = factor(sample(c('Female adults only', 'Male adults only', 
        'Male and Female Adults', 'Children only'), 
        2000, replace = TRUE , prob=c(.25,.10, .64 , .01)),
        levels = c('Female adults only', 'Male adults only', 
        'Male and Female Adults', 'Children only')),
  PovertyStatus = factor(sample(c('Non-Poor', 'Poor'), 2000, replace = TRUE, 
        prob=c(.3,.70)), levels = c('Non-Poor', 'Poor')),
  WealthQuintile = factor(sample(c('Lowest', 'Second', 'Middle', 'Fourth', 
        'Highest'), 2000, replace = TRUE), levels = c('Lowest', 'Second', 
        'Middle', 'Fourth','Highest')) ,
  Region = sample(c('NAL', 'HR1/SA2'), 2000, replace = TRUE),
  weight = runif(2000, 1, 3)   
) %>%
  mutate(
    HHEA = case_when(
      Region == 'NAL' ~ sample(1:80, n(), replace = TRUE),
      Region == 'HR1/SA2' ~ sample(81:160, n(), replace = TRUE)
    ),
    Strata = Region ,
    outcome = case_when( #we'll make outcome dependent on poverty
      PovertyStatus == 'Poor' ~ rbinom(n=2000, size=1, prob=0.25),  
      PovertyStatus == 'Non-Poor' ~ rbinom(n=2000, size=1, prob=0.5),  
    )
  ) 

# Define the survey design
design <- df %>% as_survey_design(
  ids = HHEA,       # Primary sampling unit (PSU)
  strata = Strata,   # Strata variable
  weights = weight,  # Sampling weights
  nest = TRUE        # For nested designs
)

#Run and export results
table1 <- ftftable(design, "outcome", c("Age", "Education" , 
              "PregnancyStatus" , "GenderedHouseholdType" , "PovertyStatus" , 
              "WealthQuintile"), "Region", decimals=3)

export_ftftable(table1 , workbook_name = "Table1.xlsx", sheet_name = "Table1")