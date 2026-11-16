## === World Drug Index ===
##   Computes a global drug-harm index score (0-1, higher = "better" outcome)
##   for each country, both overall and per substance category (cannabis,
##   hallucinogens, amphetamines, cocaine, opioids), split by all/male/female.

## Inputs have been hardcoded here, change to your directory
##   1. Main data workbook, called 'database_input' - one sheet per country with raw indicator values,
##      plus a "Global Min/Max" sheet used for benchmarking.
##   2. Weights file - Called 'weights' one row per indicator code, giving its composite/
##      male/female weight, group_id (for weight redistribution), and
##      substance category.

## Indicator coding convention:
##   Each measured concept has up to three separate indicator codes, one per
##   sex breakdown (e.g. 110 = all, 111 = male, 112 = female). These are
##   independent codes with their own weights and their own global benchmarks

## Output:
##   WDI_Global_Results.xlsx, written next to the main data file, containing:
##     - one tab per category (Total + 5 substances) with country scores/tiers
##     - a detailed_indicators tab (per-indicator benchmark diagnostics)
##     - one detail tab per country (indicator-level score breakdown)

## Features:
## 1. GLOBAL ONLY: No regional partitions
## 2. DATA AVAILABILITY TIERS: Tier-1/Tier-2/Tier-3 are strictly based on DATA AVAILABILITY.
## 3. DETAILED INDICATORS: Single collapsed columns for N, Avg Score, and Color
## 4. WINSORIZATION: P5-P95 applied universally to min and max normalization parameters
## 5. TIER RULE: One standard p33/p66 rule applied uniformly to every category
##    (no substance-specific exceptions). For a category/sex combination whose
##    indicator-count distribution clusters tightly enough that p33 and p66
##    land on the same or adjacent integers, Tier-2 can come out empty for
##    that combination - this is reported as-is rather than hidden by a
##    category-specific rule.


## cleans R environment
rm(list = ls())

# Ensure required packages are installed and loaded
required_packages <- c("readxl", "openxlsx", "dplyr", "stringr")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)
suppressPackageStartupMessages({
  library(readxl)
  library(openxlsx)
  library(dplyr)
  library(stringr)
})


## these are indicators that relate to the prices of substances across regions that 
## were originally imported into the database but later it was decided to not include 
## in the index calculation given that retail/market price (USD per gram) is not 
## conceptually comparable to the prevalence- and rate-based indicators that make up 
## the rest of the index - price reflects market conditions (supply, enforcement, 
## purity, local cost of living) rather than drug-use harm or prevalence
# --- 1. CONFIGURATION ---
cat('--- STEP 1: INITIALIZATION ---\n')
excluded_indicators <- c("13110", "13210", "13310", "13410", "13510", "14110", "14210",
                         "14310", "14410", "15110", "15210", "16110", "16210", "16310",
                         "18110", "18210")
excluded_indicators_numeric <- as.numeric(excluded_indicators)

## these are the sub-index codes.. note that there are indicators that do not belong to any substance class
## the logic is this: an overall composite index is built REGARDLESS of sub-codes
## then sub-indices are aggregated based on sub-codes e.g., a global cannabis index
## the non-class indicators are included in the overall indices (e.g., total, male, female)
## but are excluded from specific substance class index calculation
sub_codes <- c(1, 2, 3, 5, 6)
sub_names <- c('Canna', 'Hallu', 'Amph', 'Cocaine', 'Opioid')
run_categories <- c('Total', sub_names)

# --- 2. FILE PATHS ---
## Input/output folder and file paths. Edit INPUT_DIR (and the two filenames
## below, if yours differ) to point at your own copies before running - this
## replaces the interactive file.choose() dialogs so the script can run
## non-interactively (e.g. via Rscript, or in an automated pipeline) and so
## that anyone cloning this repo only has to edit these few lines rather
## than click through dialogs.
cat('--- STEP 2: LOCATING INPUT FILES ---\n')
INPUT_DIR <- "C:/Users/ethan/Documents/WDI/Results_AND_data/formatted_input_2_script_data+RESULTS/validation"

## Main data workbook: one sheet per country with raw indicator values, plus
## the "Global Min/Max" benchmark sheet.
filepath <- file.path(INPUT_DIR, "database_input.xlsx")
if (!file.exists(filepath)) stop(sprintf('Main data file not found: %s', filepath))
filepath_noext <- sub("\\.[^.]*$", "", filepath)
cat(sprintf('Using main data file: %s\n', filepath))

## Results workbook is written back into the same folder as the inputs.
file_master <- file.path(INPUT_DIR, 'WDI_Global_Results.xlsx')

## Weights file ("Indicator list"): one row per indicator code with its
## composite/male/female weight, group_id, and substance category.
w_filepath <- file.path(INPUT_DIR, "weights.xlsx")
if (!file.exists(w_filepath)) stop(sprintf('Weights file not found: %s', w_filepath))
cat(sprintf('Using weights file: %s\n', w_filepath))

w_data <- read_excel(w_filepath, sheet = 1, range = "A2:F1000", col_names = FALSE)
colnames(w_data) <- c('indicator', 'w_all', 'w_male', 'w_female', 'group_id', 'substance')

## some data cleaning, forcing text to numeric, also protects against floating-point artifacts in the input data
## any cell that isn't valid becomes NA which downstream is assigned a 0 score and 0 weight
## meaning NAs contribute nothing to the denominator or numerator in the downstream index calculations
w_data$indicator <- round(as.numeric(w_data$indicator))
## drop the ~893 blank padding rows from the A2:F1000 range read
w_data$w_all[is.na(as.numeric(w_data$w_all))] <- 0
w_data$w_male[is.na(as.numeric(w_data$w_male))] <- 0
w_data$w_female[is.na(as.numeric(w_data$w_female))] <- 0

w_data$w_all <- as.numeric(w_data$w_all)
w_data$w_male <- as.numeric(w_data$w_male)
w_data$w_female <- as.numeric(w_data$w_female)
w_data$substance <- as.numeric(w_data$substance)
w_data$group_id <- str_trim(as.character(w_data$group_id))

## Assign a unique "Singleton_<row>" group_id to any indicator with a blank/missing
## group_id, so every indicator belongs to some group (even if just itself) and
## can flow through the same group-based weight-redistribution logic later
missing_mask <- is.na(w_data$group_id) | w_data$group_id == ""
if (any(missing_mask)) {
  idx_m <- which(missing_mask)
  for (k in seq_along(idx_m)) {
    w_data$group_id[idx_m[k]] <- paste0("Singleton_", idx_m[k])
  }
}

# --- 3. LOAD BASE METADATA ---
## loads metadata and checks for the Min and Max sheet
cat('Loading Metadata...\n')
sheets <- excel_sheets(filepath)
global_sheet_idx <- which(str_detect(sheets, "Global") & str_detect(sheets, "Min"))[1]
if (is.na(global_sheet_idx)) stop('Could not find Global MinMax sheet.')

nt <- read_excel(filepath, sheet = sheets[global_sheet_idx], range = "A1:I1000")
colnames(nt)[1:9] <- c('Indicator', 'Category', 'Min_Value', 'Min_Year', 'Min_Sheet', 'Max_Value', 'Max_Year', 'Max_Sheet', 'Good')

## Build a clean lookup table of indicator -> "good direction" flag.
## If an indicator appears more than once, keep the LAST occurrence
## and drop any rows with no valid indicator code.
BaseMeta <- data.frame(
  indicator = round(as.numeric(nt$Indicator)),
  good = as.numeric(nt$Good)
)
BaseMeta <- BaseMeta[!duplicated(BaseMeta$indicator, fromLast = TRUE), ]
BaseMeta <- BaseMeta[!is.na(BaseMeta$indicator), ]

# --- 4. PASS 1: AGGREGATE DATA (WITH COLLAPSED TRACKING) ---

## AllRawData: global pool of every raw value seen per indicator, across all
## countries/sexes - used later to compute P5-P95 benchmarks.
## CountryList: one cleaned data table per country sheet.
## Skip any sheet with "min" in its name - those are benchmark sheets e.g., min_max global
cat('\n--- PASS 1: SCANNING DISTRIBUTION ---\n')
AllRawData <- new.env(hash = TRUE)
CountryList <- list()

for (i in seq_along(sheets)) {
  sn <- sheets[i]
  if (str_detect(tolower(sn), "min")) next
  
  ##reads the sheet and finds columns by fuzzy header matching because country sheets may
  ## not be in consistent column order/naming {should be but just in case they aren't}
  raw_t <- read_excel(filepath, sheet = sn, range = "A1:Z1000", .name_repair = "minimal")
  if (ncol(raw_t) < 3) next
  
  headers <- str_trim(tolower(colnames(raw_t)))
  idx_ind <- which(str_detect(headers, 'ind'))[1]
  idx_yr  <- which(str_detect(headers, 'year') | str_detect(headers, 'yr'))[1]
  idx_all <- which(headers %in% c('all', 'total', 'value', 'val'))[1]
  idx_m   <- which(headers %in% c('male', 'males', 'men'))[1]
  idx_f   <- which(headers %in% c('female', 'females', 'women'))[1]
  
  ## skip sheet if no identifiable indicator or all value column
  if (is.na(idx_ind) || is.na(idx_all)) next
  
  ##build the cleaned per-sheet table into a standard shape: all, year, male, female
  ## if year column not found, default to 2025
  ## fill with NA if male/female columns not found
  data <- data.frame(
    indicator = round(as.numeric(raw_t[[idx_ind]])),
    all = as.numeric(raw_t[[idx_all]])
  )
  
  if (!is.na(idx_yr)) data$year <- as.numeric(raw_t[[idx_yr]]) else data$year <- rep(2025, nrow(data))
  if (!is.na(idx_m)) data$male <- as.numeric(raw_t[[idx_m]]) else data$male <- rep(NA_real_, nrow(data))
  if (!is.na(idx_f)) data$female <- as.numeric(raw_t[[idx_f]]) else data$female <- rep(NA_real_, nrow(data))
  
  ## drop rows with no valid indicator codes or the excluded indicators mentioned at start of script
  ## IMPORTANT - if indicator has various years of data, take the most recent value
  ## keep country's cleaned table
  data <- data[!is.na(data$indicator) & !(data$indicator %in% excluded_indicators_numeric), ]
  if (nrow(data) == 0) next
  
  data <- data[order(data$year, decreasing = TRUE), ]
  data <- data[!duplicated(data$indicator), ]
  
  entry <- list(Name = sn, Data = data, Year = max(data$year, na.rm = TRUE))
  CountryList[[length(CountryList) + 1]] <- entry
  
  ## Pool raw values into a global, per-indicator-code distribution. Because
  ## all/male/female breakdowns of the same underlying measure use DIFFERENT
  ## indicator codes (e.g. 110=all, 111=male, 112=female), each code gets its
  ## own independent pool achieved by the data having separate codes, not
  ## by any sex-splitting logic in this script.
  
  
  ## Process one row (one indicator code) at a time. The first time a given
  ## code is seen anywhere, create an empty pool for it (a values vector and a matching countries vector).
  ## Fetch that indicator's current pool so it can be appended to below.
  for (j in 1:nrow(data)) {
    ind <- as.character(data$indicator[j])
    
    if (!exists(ind, envir = AllRawData)) {
      assign(ind, list(pooled_vals = numeric(), pooled_countries = character()), envir = AllRawData)
    }
    s <- get(ind, envir = AllRawData)
    
    ## If this indicator (row) reported an "all" value, append it to this code's global pool,
    ## tagging it with this country's sheet name for later min/max lookups
    ##- "male"/"female"-coded rows (111/112) don't carry an "all" value.
    if (!is.na(data$all[j])) {
      s$pooled_vals <- c(s$pooled_vals, data$all[j])
      s$pooled_countries <- c(s$pooled_countries, sn)
    }
    
    ## same as above but pools only for indicators that report something in the male column
    ## e.g., indicator codes that end in 1.. does not mix values across codes (e.g., across sex)
    if (!is.na(data$male[j])) {
      s$pooled_vals <- c(s$pooled_vals, data$male[j])
      s$pooled_countries <- c(s$pooled_countries, sn)
    }
    
    ## same as above, pools only for indicators (rows) whose code is the female-variant (e.g., ends in 2)
    if (!is.na(data$female[j])) {
      s$pooled_vals <- c(s$pooled_vals, data$female[j])
      s$pooled_countries <- c(s$pooled_countries, sn)
    }
    
    ## write update pool back to environment
    assign(ind, s, envir = AllRawData)
  }
}

# --- 5. CALCULATE BENCHMARKS & DETAILED INDICATORS ---
cat('Calculating Benchmarks & Detailed Indicators...\n')

## traffic-light classified for normalised scores; bottom third - red
## middle - orange, top third - green.. NA (No data) - "None"
get_color <- function(val) {
  if (is.na(val)) return("None")
  if (val <= 0.33333) return("Red")
  if (val >= 0.6667) return("Green")
  return("Orange")
}

## creates two empty containers, filled with upcoming loop for each indicator code
## environment is set up here instead of data.frame which would be slower when looped
## across all countries
BenchmarkStruct <- new.env(hash = TRUE)
AnalysisReport <- data.frame(
  Indicator = numeric(), Good_Flag = numeric(), Raw_Min = numeric(), Min_Country = character(),
  Winsorized_P5 = numeric(), Raw_Max = numeric(), Max_Country = character(), Winsorized_P95 = numeric(),
  N_Countries = numeric(), Avg_Normalized_Score = numeric(), Score_Color = character(),
  stringsAsFactors = FALSE
)

## this next part finds the floor and ceilings for each indicator
for (i in 1:nrow(BaseMeta)) {
  
  ## process one indicator code per loop
  ind <- BaseMeta$indicator[i]
  
  ##build look up key for Benchmark Struc - starts with an 'I' because keys need to
  ## not look like a numeric (e.g., 110 becomes 'I110')
  key <- paste0("I", ind)
  
  ## flags the 'good direction' for each indicator - will decide whether or not the
  ## indicator normalised score will get flipped later on
  gd <- BaseMeta$good[i]
  
  ## starts with worse case scenario - defaults everything to 'no data' before
  ## computing anything so that every indicator in BaseMeta gets a row entry even if no data
  p_low <- NA; p_high <- NA; raw_min <- NA; raw_max <- NA
  min_country <- "N/A"; max_country <- "N/A"
  n_countries <- 0; avg_norm <- NA; col_label <- "None"
  
  ## converts indicator codes to strings, to match AllRawData in section 4
  ## note this is a diff environment to BenchmarkStruct above
  ind_char <- as.character(ind)
  
  ## check if indicator showed up in any country's data
  if (exists(ind_char, envir = AllRawData)) {
    
    ## if existed, get the indicator pool, 's' is a copy of list fected in Section 4
    s <- get(ind_char, envir = AllRawData)
    
    ## second check, confirms there is actually one real number in pool
    if (length(s$pooled_vals) > 0) {
      
      ## compute in plain numbers the 5th and 95th percentile of every pooled value
      ## for thsi indicator
      p_low <- quantile(s$pooled_vals, 0.05, names = FALSE, na.rm = TRUE, type = 5)
      p_high <- quantile(s$pooled_vals, 0.95, names = FALSE, na.rm = TRUE, type = 5)
      
      ## finds position of smallest value in pool
      min_idx <- which.min(s$pooled_vals)
      
      ## uses position to pull min raw value, is different from p_low - only used to find
      ## min country
      raw_min <- s$pooled_vals[min_idx]
      
      ## identifies the country attached to raw_min
      min_country <- s$pooled_countries[min_idx]
      
      ##same as above, except for maximum
      max_idx <- which.max(s$pooled_vals)
      raw_max <- s$pooled_vals[max_idx]
      max_country <- s$pooled_countries[max_idx]
      
      ## counts how many values were in the pool for the indicator, lines up with
      ## number of countries given one row per code per country structure
      n_countries <- length(s$pooled_vals)
      
      ## computes floor-ceiling range
      den_range <- p_high - p_low
      
      ## only computes normalised score if range is real (greater than 0)
      if (!is.na(den_range) && den_range > 0) {
        
        ## rescale every pooled raw value on scale 0 to 1 against p low(5th percentile) and p high (95th percentile)
        ## clamp outliers to the 0-1 edges
        ## flip the scale if good is 0, e.g., when a more desirable outcome would reflect a lower score
        norm_array <- sapply(s$pooled_vals, function(v) {
          norm_val <- min(max((v - p_low) / den_range, 0), 1)
          if (gd == 0) (1 - norm_val) else norm_val
        })
        
        ## calculates the global average of each indicators normalised score
        avg_norm <- mean(norm_array, na.rm = TRUE)
        col_label <- get_color(avg_norm)
      }
    }
  }
  
  ## populates the inital environment strucutre BenchmarkStruct with the values needed
  ## to normalise a countrys raw value - the floor and ceilings
  assign(key, list(min = p_low, max = p_high), envir = BenchmarkStruct)
  
  ## apends a row to the analysis report, overviewing the variables in the brackets
  AnalysisReport[nrow(AnalysisReport) + 1, ] <- list(
    ind, gd, raw_min, min_country, p_low, raw_max, max_country, p_high, n_countries, avg_norm, col_label
  )
}

# --- 6. PASS 2: SCORING (SUBSTANCE MATRIX) ---

## creates an empty results sheet per category x sex combo - filled with one entry
## per country as the loop runs
cat('\n--- PASS 2: SCORING SUBSTANCE MATRIX ---\n')
results <- list()
for (cat_idx in seq_along(run_categories)) {
  results[[paste0('all_', run_categories[cat_idx])]] <- list()
  results[[paste0('male_', run_categories[cat_idx])]] <- list()
  results[[paste0('female_', run_categories[cat_idx])]] <- list()
}

detailed_data <- list()

## one country at a time
for (i in seq_along(CountryList)) {
  C <- CountryList[[i]]
  data_clean <- C$Data
  
  ## for one category at a time, for Total use every weighted indicator (Composite index)
  ## for specific substance category, filter wights dwon to just that substance's code
  for (sub_iter in seq_along(run_categories)) {
    current_category <- run_categories[sub_iter]
    
    if (current_category == 'Total') {
      curr_w_data <- w_data
    } else {
      real_code <- sub_codes[which(sub_names == current_category)]
      curr_w_data <- w_data[which(w_data$substance == real_code), ]
    }
    
    if (nrow(curr_w_data) == 0) next
    
    ##running totals for each country+category - start at 0 and build up as loop runs
    ## num_* - weighted sum of normalised scores
    ## den_* - sum of weights (e.g., weighted-average demoninator for aggregate indices)
    ## cnt_* - number indicators that contributed
    ## final score = num/den
    num_all=0; den_all=0; cnt_all=0
    num_male=0; den_male=0; cnt_male=0
    num_fem=0; den_fem=0; cnt_fem=0
    
    ## placeholder to record each indicator's contribution to country's overall score
    ## used in the Index_Global_All breakdown column exported later on
    if (current_category == 'Total') temp_col <- rep(NA_real_, nrow(data_clean))
    
    ## for each indicator, fetch floor and ceiling - skip if no entry exists or if
    ## it exists but no real min/max
    for (j in 1:nrow(data_clean)) {
      ind <- data_clean$indicator[j]
      key <- paste0("I", ind)
      
      if (!exists(key, envir = BenchmarkStruct)) next
      bench <- get(key, envir = BenchmarkStruct)
      b_min <- bench$min
      b_max <- bench$max
      if (is.na(b_min) || is.na(b_max)) next
      
      ## find indicator's good direction, recheck floor-ceiling range is not degenerate (like Section 5)
      ## calc_sc captures the row/indicator's b_min, b_max, and good
      row_meta <- which(BaseMeta$indicator == ind)[1]
      if (is.na(row_meta)) next
      gd <- BaseMeta$good[row_meta]
      
      den_range <- b_max - b_min
      if (is.na(den_range) || den_range == 0) next
      
      calc_sc <- function(v) {
        norm_v <- min(max((v - b_min) / den_range, 0), 1)
        if (gd == 0) (1 - norm_v) else norm_v
      }
      
      ## find indicator's weight row within current weight table
      w_row <- which(curr_w_data$indicator == ind)[1]
      if (is.na(w_row)) next
      
      ##effective weigth applied in scoring, after any of the group redistribution below
      ## again, defaults to 0 until actually computed
      eff_w_all=0; eff_w_male=0; eff_w_female=0
      gid <- curr_w_data$group_id[w_row]
      
      ## next part asks the question - how much weight should an indicator carry for a country
      ## given what else is or isn't present in the group(in column E in the indicator list input excel)
      ## basically indicators that belong to some groups (e.g., prevalence of cannabis use over lifetime, month, year)
      ## were assigned different weights all of which within a group should equal 0
      ## therefore, if one indicator was missing, we still want the group to contribute 100 weight to the score
      ## so we redistribute any outstanding weight among the indicators with data
      ## runs seperately for all, male, and female since they have seperate weights
      ## the redistribution is proportional - so the indicator with the highest original weight
      ## will be assigned the greatest proportion of the outstanding weights
      if (!is.na(gid) && gid != "") {
        
        ## find every weight row sharing the group_id, and check which the country reports
        ## data for
        grp_idx <- which(tolower(str_trim(curr_w_data$group_id)) == tolower(gid))
        grp_inds <- curr_w_data$indicator[grp_idx]
        
        is_present <- grp_inds %in% data_clean$indicator
        loc_in_clean <- match(grp_inds, data_clean$indicator)
        
        ## target - groups full intended/theoretical weight (should be 100 but 100 was not hard-coded to allow flexibility)
        ## active - weight usable now (e.g., for indicators that have data)
        target_all <- sum(curr_w_data$w_all[grp_idx], na.rm = TRUE); active_all <- 0
        for (k in seq_along(grp_idx)) {
          if (is_present[k] && !is.na(data_clean$all[loc_in_clean[k]])) active_all <- active_all + curr_w_data$w_all[grp_idx[k]]
        }
        ## rescale this indicator's own weight by (target/active) -- present indicators in the group
        ## will sum to group's full intended weight.. ratio 1 when nothgin in the group is missing
        if (active_all > 0) eff_w_all <- curr_w_data$w_all[w_row] * (target_all / active_all)
        
        ## same as above but for males
        target_male <- sum(curr_w_data$w_male[grp_idx], na.rm = TRUE); active_male <- 0
        for (k in seq_along(grp_idx)) {
          if (is_present[k] && !is.na(data_clean$male[loc_in_clean[k]])) active_male <- active_male + curr_w_data$w_male[grp_idx[k]]
        }
        if (active_male > 0) eff_w_male <- curr_w_data$w_male[w_row] * (target_male / active_male)
        
        ## same as above but for females
        target_fem <- sum(curr_w_data$w_female[grp_idx], na.rm = TRUE); active_fem <- 0
        for (k in seq_along(grp_idx)) {
          if (is_present[k] && !is.na(data_clean$female[loc_in_clean[k]])) active_fem <- active_fem + curr_w_data$w_female[grp_idx[k]]
        }
        if (active_fem > 0) eff_w_female <- curr_w_data$w_female[w_row] * (target_fem / active_fem)
      }
      
      ##score and accumulate whatevr of all/male/female actually has data for this indicator
      ## gives a usuable (i.e., greater than 0) effective weight
      ## sc*effect weight - adds to the weighter numerator
      ## eff_w * alone - adds to the denominator
      ## count - tracks number indicators that contribute
      if (!is.na(data_clean$all[j]) && eff_w_all > 0) {
        sc <- calc_sc(data_clean$all[j])
        num_all <- num_all + (sc * eff_w_all); den_all <- den_all + eff_w_all; cnt_all <- cnt_all + 1
        if (current_category == 'Total') temp_col[j] <- sc * eff_w_all
      }
      if (!is.na(data_clean$male[j]) && eff_w_male > 0) {
        sc <- calc_sc(data_clean$male[j])
        num_male <- num_male + (sc * eff_w_male); den_male <- den_male + eff_w_male; cnt_male <- cnt_male + 1
      }
      if (!is.na(data_clean$female[j]) && eff_w_female > 0) {
        sc <- calc_sc(data_clean$female[j])
        num_fem <- num_fem + (sc * eff_w_female); den_fem <- den_fem + eff_w_female; cnt_fem <- cnt_fem + 1
      }
    }
    
    ## attach the per-indicator breakdown column for total only
    if (current_category == 'Total') data_clean$INDEX_Global_All <- temp_col
    
    ## final score for each index = weighted average (numerator/denominator)
    ## NA if no data
    score_all <- NA_real_; if (den_all > 0) score_all <- num_all / den_all
    score_male <- NA_real_; if (den_male > 0) score_male <- num_male / den_male
    score_fem <- NA_real_; if (den_fem > 0) score_fem <- num_fem / den_fem
    
    ##put the result into the matching results bucket
    results[[paste0('all_', current_category)]][[length(results[[paste0('all_', current_category)]]) + 1]] <- list(C$Name, C$Year, score_all, cnt_all)
    results[[paste0('male_', current_category)]][[length(results[[paste0('male_', current_category)]]) + 1]] <- list(C$Name, C$Year, score_male, cnt_male)
    results[[paste0('female_', current_category)]][[length(results[[paste0('female_', current_category)]]) + 1]] <- list(C$Name, C$Year, score_fem, cnt_fem)
  }
  
  ## after everything scored, lock it in, and tidy up the column order for the table
  col_order <- c('indicator', 'year', 'all', 'male', 'female', 'INDEX_Global_All')
  final_cols <- col_order[col_order %in% colnames(data_clean)]
  data_clean <- data_clean[, final_cols, drop = FALSE]
  
  ## transforms country's sheet name into a valid R identifier
  ## but make.names -- weird text with dots where spaces should be
  ## so the untransformed country sheet name is also stored under a separate entry
  ## untransformed name used later in Section 8 to make the output Excel
  valid_name <- make.names(C$Name)
  detailed_data[[valid_name]] <- data_clean
  detailed_data[[paste0(valid_name, '_original_name')]] <- C$Name
}

# --- 7. BUILD GLOBAL TIERS (BASED ON DATA AVAILABILITY) ---
## calculates tiers of data availability, based on percentiles of available indicators
cat('\n--- STEP 5: CALCULATING RELATIVE TIERS ---\n')

##GlobalTiers - the 33rd, 66th percentile indicator-count threshold, for every
## aggregate/category - used to assign tiers later on
GlobalTiers <- list()
sx_list <- c('all', 'male', 'female')

## for every aggreate x category combination (remember that total is also a category)
## note, total relates to all indicators regardless of substance class (Composite index)
## Countries with 0 valid indicators excluded from percentile calculation
## type = 5 matchs MATLAB quantile convention
## if no country has valid data, both thresholds default to 0 important for downstream comparions
for (sx in sx_list) {
  for (sub_n in run_categories) {
    key <- paste0(sx, '_', sub_n)
    
    if (!is.null(results[[key]]) && length(results[[key]]) > 0) {
      counts <- sapply(results[[key]], function(x) as.numeric(x[[4]]))
      valid_counts <- counts[counts > 0]
      
      if (length(valid_counts) > 0) {
        GlobalTiers[[paste0(key, '_p33')]] <- quantile(valid_counts, 0.33, names = FALSE, type = 5)
        GlobalTiers[[paste0(key, '_p66')]] <- quantile(valid_counts, 0.66, names = FALSE, type = 5)
      } else {
        GlobalTiers[[paste0(key, '_p33')]] <- 0
        GlobalTiers[[paste0(key, '_p66')]] <- 0
      }
    }
  }
}

## classify each country's indicator count into three tiers, relative to
## the p33/p66 threshold for this aggregate/category combination
## default to "None", only entries with count greater 0 eligible for classification below
##
## NOTE: a single rule is now applied to every category - no substance-specific
## exceptions. When p33 and p66 land on adjacent integers (or are equal), this
## default rule can leave Tier-2 mathematically empty for that aggregate/sex
## combination - that is an honest reflection of indicator-count availability
## clustering tightly around one or two values for that category, not a bug,
## and should be documented as such rather than patched around with a
## category-specific carve-out.
get_tier <- function(cnt, p33, p66, category) {
  tiers <- rep("None", length(cnt))
  valid <- (cnt > 0)
  
  tiers[valid & cnt >= p66] <- "Tier-1"
  tiers[valid & cnt < p66 & cnt > p33] <- "Tier-2"
  tiers[valid & cnt <= p33] <- "Tier-3"
  
  return(tiers)
}

# --- 8. EXPORT TO EXCEL ---
cat('\n--- STEP 6: WRITING TO WORKBOOK ---\n')

## create an empty in-memory workbook
## worksheets appended progressively
## derive each category's substance-filtered weigth subset using same logic
## as Section 6
wb <- createWorkbook()

for (cat in run_categories) {
  if (cat == 'Total') {
    curr_w_data <- w_data
  } else {
    curr_w_data <- w_data[which(w_data$substance == sub_codes[which(sub_names == cat)]), ]
  }
  
  ## first, excludes blacklisted indicators at start of script
  ## then counts distinct indicator codes with weigths greater than 0 for each aggregate
  valid_inds <- curr_w_data[!(curr_w_data$indicator %in% excluded_indicators_numeric), ]
  max_ind_all <- length(unique(valid_inds$indicator[valid_inds$w_all > 0]))
  max_ind_male <- length(unique(valid_inds$indicator[valid_inds$w_male > 0]))
  max_ind_fem <- length(unique(valid_inds$indicator[valid_inds$w_female > 0]))
  
  key_all <- paste0('all_', cat)
  
  ## only proceeds if this category accumulate any country results in Section 6
  ## Each result is a list of lists with one list per country
  ## country list - (Name, Year, score, count)
  ## do.call(rbind,,,) - vertically concatenates lists into a single-row data frame, one row per country
  ## repeats for all, male, female
  if (!is.null(results[[key_all]]) && length(results[[key_all]]) > 0) {
    
    rc_all <- do.call(rbind, lapply(results[[key_all]], function(x) data.frame(Country = as.character(x[[1]]), Year = as.numeric(x[[2]]), Score = as.numeric(x[[3]]), Count = as.numeric(x[[4]]), stringsAsFactors = FALSE)))
    rc_male <- do.call(rbind, lapply(results[[paste0('male_', cat)]], function(x) data.frame(Score = as.numeric(x[[3]]), Count = as.numeric(x[[4]]), stringsAsFactors = FALSE)))
    rc_fem <- do.call(rbind, lapply(results[[paste0('female_', cat)]], function(x) data.frame(Score = as.numeric(x[[3]]), Count = as.numeric(x[[4]]), stringsAsFactors = FALSE)))
    
    ## extract each relevant column into own named vector
    ## Index_year is a hardcoded constant (2025) applied to every country - because
    ## data sourced from World Drug Report 2025.. note this is different from latest_data_year
    ## which refers to the most recent year for which data was available for a specific indicator
    ## for each country
    Countries <- rc_all$Country
    Index_Year <- rep(2025, length(Countries))
    Latest_Data_Year <- rc_all$Year
    
    Score_Overall <- rc_all$Score; Ind_Count_Overall <- rc_all$Count
    Score_Male <- rc_male$Score; Ind_Count_Male <- rc_male$Count
    Score_Female <- rc_fem$Score; Ind_Count_Female <- rc_fem$Count
    
    ## classify each country into tiers of data availability using the aggregate' own
    ## count vector against that aggregate's own p33/p66 thresholds (Section 7)
    Tier_Overall <- get_tier(Ind_Count_Overall, GlobalTiers[[paste0('all_', cat, '_p33')]], GlobalTiers[[paste0('all_', cat, '_p66')]], cat)
    Tier_Male <- get_tier(Ind_Count_Male, GlobalTiers[[paste0('male_', cat, '_p33')]], GlobalTiers[[paste0('male_', cat, '_p66')]], cat)
    Tier_Female <- get_tier(Ind_Count_Female, GlobalTiers[[paste0('female_', cat, '_p33')]], GlobalTiers[[paste0('female_', cat, '_p66')]], cat)
    
    ## counts missing indicators, using the theoretical maximums (derived Section 2)
    ## for each aggregate - the actual count.
    ## floors at 0
    Ind_Missing_Overall <- pmax(0, max_ind_all - Ind_Count_Overall)
    Ind_Missing_Male <- pmax(0, max_ind_male - Ind_Count_Male)
    Ind_Missing_Female <- pmax(0, max_ind_fem - Ind_Count_Female)
    
    ## Assemble final data frame - 15 columns, which become header row in exported sheet
    ## sorts by overall score, best country first
    ## countries with no score at the bottom using na.last = TRUE
    T_df <- data.frame(
      Country = Countries, Index_Year = Index_Year, Latest_Data_Year = Latest_Data_Year,
      Score_Overall = Score_Overall, Tier_Overall = Tier_Overall, Ind_Count_Overall = Ind_Count_Overall, Ind_Missing_Overall = Ind_Missing_Overall,
      Score_Male = Score_Male, Tier_Male = Tier_Male, Ind_Count_Male = Ind_Count_Male, Ind_Missing_Male = Ind_Missing_Male,
      Score_Female = Score_Female, Tier_Female = Tier_Female, Ind_Count_Female = Ind_Count_Female, Ind_Missing_Female = Ind_Missing_Female,
      stringsAsFactors = FALSE
    )
    
    T_df <- T_df[order(T_df$Score_Overall, decreasing = TRUE, na.last = TRUE), ]
    
    addWorksheet(wb, cat)
    writeData(wb, cat, T_df)
  }
}

## write Detailed Indicators Tab as per Section 5
addWorksheet(wb, 'detailed_indicators')
writeData(wb, 'detailed_indicators', AnalysisReport)

## write Country Detailed Tabs
## names (detailed_data) holds TWO keys per country - the mangled and unmangled coutry
## names from Section 6
## sorts alphabetically
cnames <- str_sort(names(detailed_data), numeric = TRUE)
for (k in seq_along(cnames)) {
  if (endsWith(cnames[k], '_original_name')) next
  
  ## retrive the lossless orignal country names (stored in Section 6)
  ## string every category not allowed in Excel sheet names (e.g. things
  ## like : / ?)
  orig_name <- detailed_data[[paste0(cnames[k], '_original_name')]]
  sn_scrub <- gsub('[:\\\\/\\?\\*\\[\\]]', ' ', orig_name)
  
  ## truncates to Excel's max 31-character sheet-name limit
  sn_scrub <- substr(sn_scrub, 1, min(nchar(sn_scrub), 31))
  
  addWorksheet(wb, paste0(sn_scrub, '_det'))
  writeData(wb, paste0(sn_scrub, '_det'), detailed_data[[cnames[k]]])
}

saveWorkbook(wb, file_master, overwrite = TRUE)
cat(sprintf('SUCCESS! Master file generated: %s\n', file_master))