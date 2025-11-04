#********************************************************************************************************#
#                  IMTS Data processing and re-shaping for SBS Data portal                               #
#                              August 2025                                                               #  
#                                                                                                        #
#********************************************************************************************************#

# Load required libraries
library(openxlsx) #create, read and write excel files
library(readxl)
library(dplyr)
library(tidyverse)

# Map working directory
repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path))
setwd(repository)

# Sourcing table 1-7 functions scripts
source("functions/imts_functions.R")

#### *********************** IMTS table 1 processing *********************************** ####

# read in table1
table1 <- imts_table1(imts_tab1)

#Create the percentage change between the same month of two consecutive years
table1_ptC <- table1 |>
  filter(TRANSFORMATION == "N" & FREQ == "M") |>
  mutate(X = as.numeric(X), M = as.numeric(M), B = as.numeric(B)) |>
  arrange(TIME_PERIOD) |>
  mutate(
    X_ptC = 100 * (X / lag(X, 12) - 1),
    M_ptC = 100 * (M / lag(M, 12) - 1),
    B_ptC = 100 * (B / lag(B, 12) - 1)
  )

table1_ptC <- table1_ptC |>
  mutate(X = X_ptC, M = M_ptC, B = B_ptC,
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
         ) |>
  select(-c("X_ptC", "M_ptC", "B_ptC"))

# Combine the percentage change with main table
table1_merge_table1_ptC <- rbind(table1, table1_ptC)

# Reshape from wide to long format
table1_long <- table1_merge_table1_ptC %>%
  pivot_longer(
    cols = -c(DATAFLOW:TIME_PERIOD),
    names_to = "TRADE_FLOW",
    values_to = "OBS_VALUE"
  )

# Re-order the columns in the proper order
table1_final <- table1_long |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, TIME_PERIOD,
         OBS_VALUE, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS)
  
table1_final[is.na(table1_final)] <- ""
table1_final <- table1_final |> mutate(OBS_STATUS = ifelse(OBS_STATUS == "NA" | OBS_STATUS == "*", "", OBS_STATUS))

table1_final$TIME_PERIOD <- sapply(table1_final$TIME_PERIOD, as.character)
                                       
#Output table1 to csv output
write.csv(table1_final, "output/imts/DF_IMTS_TABLE1.csv", row.names = FALSE)

#### *********************** IMTS table 2 processing *********************************** ####

# read in the table1
table2 <- imts_table2(imts_tab2)

# Table 2 Annual percentage change 
table2_yr_PT <- table2

# Convert TIME_PERIOD to Date
table2_yr_PT <- table2_yr_PT %>%
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) %>%
  arrange(TIME_PERIOD)

# List of columns to calculate change
hs_cols <- grep("^HS(_\\d+)?$", names(table2_yr_PT), value = TRUE)

# Create a lagged version (12 months before)
table2_pct_change_yr <- table2_yr_PT |>
  mutate(across(all_of(hs_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
         )
  
# table2 monthly percentage change 
table2_mth_PT <- table2

# Convert TIME_PERIOD to Date
table2_mth_PT <- table2_mth_PT |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) |>
  arrange(TIME_PERIOD)

# List of columns to calculate change
hs_cols <- grep("^HS(_\\d+)?$", names(table2_mth_PT), value = TRUE)

# Calculate monthly % change
table2_pct_change_mth <- table2_mth_PT |>
  mutate(across(all_of(hs_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

table2_year <- table2 |>
  mutate(TIME_PERIOD = substr(TIME_PERIOD, 1, 4)
  )

# Calculate the annual amount and percentage change 
#Determine whether year data is complete
year_data_amnt <- table2 |> group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(hs_cols), sum, na.rm = TRUE), .groups = "drop"
            ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE2(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "X", COUNTERPART_AREA = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
         ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(hs_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
         ) |>
  select(-numMonths)

#Drop the numMonths column from the annual amount table
year_data_amnt <- year_data_amnt |> select(-numMonths)

# Combine and re-shape the table 
table2_combine <- rbind(year_data_amnt, year_data_per, table2, table2_pct_change_yr, table2_pct_change_mth) |>
  select(DATAFLOW:DECIMALS, everything())

# Reshape from wide to long format
table2_long <- table2_combine |>
  mutate(across(starts_with("HS_"):`HS`, as.numeric)) |>
  pivot_longer(
    cols = -c(DATAFLOW:TIME_PERIOD),
    names_to = "COMMODITY",
    values_to = "OBS_VALUE"
  )

# Re-order the columns in the proper order
table2_final <- table2_long |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, TIME_PERIOD,
         OBS_VALUE, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" |OBS_VALUE == "NA", "", OBS_VALUE)
         )

#Output table2 to csv output
write.csv(table2_final, "output/imts/DF_IMTS_TABLE2.csv", row.names = FALSE)

#### *********************** IMTS table 3 processing *********************************** ####

# read in table3
table3 <- imts_table3(imts_tab3)

# table 3 Annual percentage change
table3_yr_PT <- table3

# Convert TIME_PERIOD to Date
table3_yr_PT <- table3_yr_PT |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) |>
  arrange(TIME_PERIOD)

# List of columns to calculate change
sitc_cols <- grep("^SITC(_\\d+)?$", names(table3_yr_PT), value = TRUE)

# Create a lagged version (12 months before)
table3_pct_change_yr <- table3_yr_PT |>
  mutate(across(all_of(sitc_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# table3 monthly percentage change
table3_mth_PT <- table3

# Convert TIME_PERIOD to Date
table3_mth_PT <- table3_mth_PT |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) |>
  arrange(TIME_PERIOD)

# Calculate monthly % change
table3_pct_change_mth <- table3_mth_PT |>
  mutate(across(all_of(sitc_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# Calculate the annual amount and percentage change
#Determine whether year data is complete
year_data_amnt <- table3 |> group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(sitc_cols), sum, na.rm = TRUE), .groups = "drop"
  ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE3(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "X", COUNTERPART_AREA = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
  ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(sitc_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
  ) |>
  select(-numMonths)

#Drop the numMonths column from the annual amount table
year_data_amnt <- year_data_amnt |> select(-numMonths)

# Combine and re-shape the table 
table3_combine <- rbind(year_data_amnt, year_data_per, table3, table3_pct_change_yr, table3_pct_change_mth) |>
  select(DATAFLOW:DECIMALS, everything())

# Reshape from wide to long format
table3_long <- table3_combine %>%
  mutate(across(starts_with("SITC_"):`SITC`, as.numeric)) %>%
  pivot_longer(
    cols = -c(DATAFLOW:TIME_PERIOD),
    names_to = "COMMODITY",
    values_to = "OBS_VALUE"
  )

# Re-order the columns in the proper order
table3_final <- table3_long |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, TIME_PERIOD,
         OBS_VALUE, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" |OBS_VALUE == "NA", "", OBS_VALUE)
         )

#Output table3 to csv output
write.csv(table3_final, "output/imts/DF_IMTS_TABLE3.csv", row.names = FALSE)

#### *********************** IMTS table 4 processing *********************************** ####

# Read in table4
table4_all <- imts_table4_all(table4_all)

table4_yr_PT <- table4_all |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) %>%
  arrange(TIME_PERIOD)

# List of region columns which will be used to calculate annual change
start_col <- which(names(table4_all)=="DECIMALS")
counterpart_cols <- names(table4_all)[(start_col + 1):ncol(table4_all)]

# Reformat the TIME_PERIOD column
table4_all <- table4_all |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD),
         across((start_col + 1):ncol(table4_all), ~ as.numeric(.))
  ) |>
  arrange(TIME_PERIOD)

# Read in the counterpart csv file
counterpart <- read.csv("other/counterpart_area.csv")
counterpart <- counterpart |> rename(geography = Name)

# Create a lagged version (12 months before)
table4_pct_change_yr <- table4_all |>
  mutate(across(all_of(counterpart_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# Calculate monthly % change
table4_pct_change_mth <- table4_all |>
  mutate(across(all_of(counterpart_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

#Determine whether year data is complete
year_data_amnt <- table4_all |> 
  group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(counterpart_cols), sum, na.rm = TRUE), .groups = "drop"
  ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE4(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "X", COMMODITY = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
  ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(counterpart_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
  ) |>
  select(-numMonths)

#Drop the numMonths column from the annual amount table
year_data_amnt <- year_data_amnt |> select(-numMonths)

# Combine and re-shape the table 
table4_all <- table4_all |> mutate(TIME_PERIOD = substr(TIME_PERIOD, 1, 7))

table4_combine <- rbind(year_data_amnt, year_data_per, table4_all, table4_pct_change_yr, table4_pct_change_mth) |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, TRANSFORMATION, TIME_PERIOD, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, everything())

# Reshape from wide to long format
table4_long <- table4_combine |>
  pivot_longer(
    cols = -c(DATAFLOW:DECIMALS),
    names_to = "geography",
    values_to = "OBS_VALUE"
  )

table4_final <- merge(table4_long, counterpart, by = "geography") |>
  select(-geography) |>
  rename(COUNTERPART_AREA = Id) |>
  relocate(COUNTERPART_AREA, .before = TRANSFORMATION) |>
  relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" |OBS_VALUE == "NA", "", OBS_VALUE)
         )

#Output table4 to csv output
write.csv(table4_final, "output/imts/DF_IMTS_TABLE4.csv", row.names = FALSE)


#### *********************** IMTS table 5 processing *********************************** ####

# Read in table5
table5 <- imts_table5(imts_tab5)

# Table 5 Annual percentage change

table5_yr_PT <- table5 |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) |>
  arrange(TIME_PERIOD)

# List of columns to calculate change
hs_cols <- grep("^HS(_\\d+)?$", names(table5_yr_PT), value = TRUE)

# Create a lagged version (12 months before)
table5_pct_change_yr <- table5_yr_PT |>
  mutate(across(all_of(hs_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

#table5 monthly percentage change
table5_mth_PT <- table5 |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) |>
  arrange(TIME_PERIOD)

# List of columns to calculate change
hs_cols <- grep("^HS(_\\d+)?$", names(table5_mth_PT), value = TRUE)

# Calculate monthly % change
table5_pct_change_mth <- table5_mth_PT |>
  mutate(across(all_of(hs_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# Calculate the annual amount and percentage change
# Determine whether year data is complete
year_data_amnt <- table5 |> group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(hs_cols), sum, na.rm = TRUE), .groups = "drop"
  ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE5(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "M", COUNTERPART_AREA = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
  ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(hs_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
  ) |>
  select(-numMonths)

year_data_amnt <- year_data_amnt |> select(-numMonths)

# Combine and re-shape the table
table5_combine <- rbind(year_data_amnt, year_data_per, table5, table5_pct_change_yr, table5_pct_change_mth) |>
  select(DATAFLOW:DECIMALS, everything())

# Reshape from wide to long format
table5_long <- table5_combine |>
  mutate(across(starts_with("HS_"):`HS`, as.numeric)) %>%
  pivot_longer(
    cols = -c(DATAFLOW:TIME_PERIOD),
    names_to = "COMMODITY",
    values_to = "OBS_VALUE"
  )

# Re-order the columns in the proper order
table5_final <- table5_long |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, TIME_PERIOD,
         OBS_VALUE, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" |OBS_VALUE == "NA", "", OBS_VALUE)
  )

#Output table5 to csv output
write.csv(table5_final, "output/imts/DF_IMTS_TABLE5.csv", row.names = FALSE)

#### *********************** IMTS table 6 processing *********************************** ####

# Read in table6
table6 <- imts_table6(imts_tab6)
# Table 6 Annual percentage change

table6_yr_PT <- table6
# Convert TIME_PERIOD to Date
table6_yr_PT <- table6_yr_PT |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD),
         across(c(SITC, starts_with("SITC_")), as.numeric)) |>
  arrange(TIME_PERIOD)

# List of columns to calculate change
sitc_cols <- grep("^SITC(_\\d+)?$", names(table6_yr_PT), value = TRUE)

# Create a lagged version (12 months before)
table6_pct_change_yr <- table6_yr_PT |>
  mutate(across(all_of(sitc_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# table6 monthly percentage change
table6_mth_PT <- table6

# Convert TIME_PERIOD to Date
table6_mth_PT <- table6_mth_PT |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD),
         across(c(SITC, starts_with("SITC_")), as.numeric)) |>
  arrange(TIME_PERIOD)

# Calculate monthly % change
table6_pct_change_mth <- table6_mth_PT |>
  mutate(across(all_of(sitc_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# Calculate the annual amount and percentage change
#Determine whether year data is complete
year_data_amnt <- table6 |>
  mutate(across(c(SITC, starts_with("SITC_")), as.numeric))

year_data_amnt <- year_data_amnt |> 
  group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(sitc_cols), sum, na.rm = TRUE), .groups = "drop"
  ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE6(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "M", COUNTERPART_AREA = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
  ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(sitc_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
  ) |>
  select(-numMonths)

#Drop the numMonths column from the annual amount table
year_data_amnt <- year_data_amnt |> select(-numMonths)

# Combine and re-shape the table 
table6_combine <- rbind(year_data_amnt, year_data_per, table6, table6_pct_change_yr, table6_pct_change_mth) |>
  select(DATAFLOW:DECIMALS, everything())

# Reshape from wide to long format
table6_long <- table6_combine %>%
  mutate(across(starts_with("SITC_"):`SITC`, as.numeric)) %>%
  pivot_longer(
    cols = -c(DATAFLOW:TIME_PERIOD),
    names_to = "COMMODITY",
    values_to = "OBS_VALUE"
  )

# Re-order the columns in the proper order
table6_final <- table6_long |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, TIME_PERIOD,
         OBS_VALUE, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" |OBS_VALUE == "NA", "", OBS_VALUE)
  )

#Output table6 to csv output
write.csv(table6_final, "output/imts/DF_IMTS_TABLE6.csv", row.names = FALSE)

#### *********************** IMTS table 7 processing *********************************** ####

# Read in table7
table7_all <- imts_table7_all(table7_all)

table7_yr_PT <- table7_all |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD)) %>%
  arrange(TIME_PERIOD)

# List of region columns which will be used to calculate annual change
start_col <- which(names(table7_all)=="DECIMALS")
counterpart_cols <- names(table7_all)[(start_col + 1):ncol(table7_all)]

# Reformat the TIME_PERIOD column
table7_all <- table7_all |>
  mutate(TIME_PERIOD = ym(TIME_PERIOD),
         across((start_col + 1):ncol(table7_all), ~ as.numeric(.))
  ) |>
  arrange(TIME_PERIOD)

# Read in the counterpart csv file
counterpart <- read.csv("other/counterpart_area.csv")
counterpart <- counterpart |> rename(geography = Name)

# Create a lagged version (12 months before)
table7_pct_change_yr <- table7_all |>
  mutate(across(all_of(counterpart_cols), #apply following calculations to the multiple columns
                ~ (.-lag(., 12)) / lag(., 12) * 100, 
                .names = "{.col}"),
         TRANSFORMATION = "G1Y", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

# Calculate monthly % change
table7_pct_change_mth <- table7_all |>
  mutate(across(all_of(counterpart_cols),
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         
         TRANSFORMATION = "G1M", UNIT_MEASURE = "PT", UNIT_MULT = "",
         TIME_PERIOD = substr(TIME_PERIOD, 1, 7)
  )

#Determine whether year data is complete
year_data_amnt <- table7_all |> group_by(substr(TIME_PERIOD, 1, 4)) |> 
  summarise(numMonths = n(),
            across(all_of(counterpart_cols), sum, na.rm = TRUE), .groups = "drop"
  ) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE7(1.0)", FREQ = "A", REF_AREA = "WS", TRADE_FLOW = "M", COMMODITY = "_T", TRANSFORMATION = "N",
         UNIT_MEASURE = "WST", UNIT_MULT = "3", OBS_STATUS = "", COMMENT = "", DECIMALS = "1"
  ) |>
  rename(TIME_PERIOD = 1) |>
  filter(numMonths == 12)

# Calculate the percentage change of the years
year_data_per <- year_data_amnt |>
  mutate(across(all_of(counterpart_cols), 
                ~ (.-lag(.)) / lag(.) * 100,
                .names = "{.col}"),
         TRANSFORMATION = "G1Y",
         UNIT_MEASURE = "PT",
         UNIT_MULT = ""
  ) |>
  select(-numMonths)

#Drop the numMonths column from the annual amount table
year_data_amnt <- year_data_amnt |> select(-numMonths)

#Combine and re-shape the table
table7_all <- table7_all |> mutate(TIME_PERIOD = substr(TIME_PERIOD, 1, 7))

table7_combine <- rbind(year_data_amnt, year_data_per, table7_all, table7_pct_change_yr, table7_pct_change_mth) |>
  select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, COMMODITY, TRANSFORMATION, TIME_PERIOD, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, everything())

# Reshape from wide to long format
table7_long <- table7_combine |>
  pivot_longer(
    cols = -c(DATAFLOW:DECIMALS),
    names_to = "geography",
    values_to = "OBS_VALUE"
  )

table7_final <- merge(table7_long, counterpart, by = "geography") |>
  select(-geography) |>
  rename(COUNTERPART_AREA = Id) |>
  relocate(COUNTERPART_AREA, .before = TRANSFORMATION) |>
  relocate(OBS_VALUE, .before = UNIT_MEASURE) |>
  mutate(across(everything(), ~replace(., is.na(.), "")),
         OBS_VALUE = ifelse(OBS_VALUE == "Inf" | OBS_VALUE == "NA" , "", OBS_VALUE),
         OBS_STATUS = ifelse(OBS_STATUS == "NA", "", OBS_STATUS)
         )

#Output table7 to csv output
write.csv(table7_final, "output/imts/DF_IMTS_TABLE7.csv", row.names = FALSE)

#### *********************** Clear Working directory (src) ***************************** ####

# Drop downloaded excel file
excel_files <- list.files(pattern = "\\.xlsx?$")
file.remove(excel_files)


