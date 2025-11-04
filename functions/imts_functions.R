# Load libraries
library(openxlsx)
library(readxl)
library(dplyr)
library(tidyverse)
library(ISOcodes)

#### ***************************** Process annual commodity prices ********************************************* ####

url <- "https://www.sbs.gov.ws/documents/economics/Merchandise-Trade/2025/Monthly_trade_tab_Sep_2025.xlsx"
destfile <- "Monthly_trade_tab_Sep_2025.xlsx"
download.file(url, destfile, mode = "wb")

#### ***************************** IMTS table1 processing function ********************************************* ####

imts_table1 <- function(imts_tab1){
 
  data <- read_excel(destfile, sheet = "Table1")
  
  # Drop row 1 and 3
  data <- data[-c(1,3), ]
  
  # Drop last 9 records
  data <- data[1:(nrow(data) - 9), ]
  
  # Replace the column headings with the first record
  colnames(data) <- as.character(unlist(data[1, ]))
  
  # Remove the first row (now used as column names)
  data <- data[-1, ]
  
  # Drop blank columns
  data <- data[, -c(2, 7)]
  
  # Rename percentage columns by adding _pt to the names and rename month column to month
  colnames(data)[6:8] <- paste0(colnames(data)[6:8], "_pt")
  colnames(data)[2] <- "month"
  
  data <- data |> filter(!is.na(Exports) & !is.na(Imports) & !is.na(Balance))
  
  data$YearTracker <- NA
  data_v1 <- data |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        grepl("\\(P\\)", Period) ~ "P",
        grepl("\\(R\\)", Period) ~ "R",
        grepl("\\(\\*\\)", Period) ~ "*",
        TRUE ~ "NA"
      ),
      
      month = gsub("\\(P\\)|\\(R\\)|\\(\\*\\)", "", month),
      Period = gsub("\\(P\\)|\\(R\\)|\\(\\*\\)", "", Period),
    )
  
  data_v2 <- data_v1 |>
    mutate(monthId = case_when(
      grepl("January", month) ~ "01",
      grepl("February", month) ~ "02",
      grepl("March", month) ~ "03",
      grepl("April", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("June", month) ~ "06",
      grepl("July", month) ~ "07",
      grepl("August", month) ~ "08",
      grepl("September", month) ~ "09",
      grepl("October", month) ~ "10",
      grepl("November", month) ~ "11",
      grepl("December", month) ~ "12",
      TRUE ~ "NA"
    ))
  
  for (i in 1:nrow(data_v2)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data_v2[i, 1])) {
      data_v2$YearTracker[i:nrow(data_v2)] <- data_v2[i, 1]
    }
  }
  
  data_v2 <- data_v2 |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  table1_amt <- data_v2 |>
    select(TIME_PERIOD, Exports, Imports, Balance, OBS_STATUS) |>
    rename(X = Exports,
           M = Imports,
           B = Balance
    ) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE1(1.0)",
           FREQ = if_else(str_detect(TIME_PERIOD, "^\\d{4}-\\d{2}$"), "M", "A"),
           REF_AREA = "WS",
           COMMODITY = "_T",
           COUNTERPART_AREA = "_T",
           TRANSFORMATION = "N",
           UNIT_MEASURE = "WST",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    )
  
  table1_pct <- data_v2 |>
    select(TIME_PERIOD, Export_pt, Imports_pt, Balance_pt, OBS_STATUS) |>
    rename(X = Export_pt,
           M = Imports_pt,
           B = Balance_pt
    ) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE1(1.0)",
           FREQ = if_else(str_detect(TIME_PERIOD, "^\\d{4}-\\d{2}$"), "M", "A"),
           REF_AREA = "WS",
           COMMODITY = "_T",
           COUNTERPART_AREA = "_T",
           TRANSFORMATION = if_else(str_detect(TIME_PERIOD, "^\\d{4}-\\d{2}$"), "G1M", "G1Y"),
           UNIT_MEASURE = "PT",
           UNIT_MULT = "",
           COMMENT = "",
           DECIMALS = 1
    )
  
  
  table1 <- rbind(table1_amt, table1_pct)
  
  imts_tab1 <- table1 |>
    select(DATAFLOW, FREQ, REF_AREA, COMMODITY, COUNTERPART_AREA, TRANSFORMATION, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, TIME_PERIOD, X, M, B)
  
  return(imts_tab1)
  
}


#### ***************************** IMTS table2 processing function ********************************************* ####

imts_table2 <- function(imts_tab2){
  data <- read_excel(destfile, sheet = "Table 2")
  
  # Transpose the read in table
  data <- t(data)
  
  # Drop columns 1 and 2 and row 2
  data <- data[, -c(1,2)]
  data <- data[-c(2), ]
  
  # Drop last 9 columns
  data <- data[, 1:(ncol(data) - 9)]
  
  # Replace the column headings with the first record
  colnames(data) <- as.character(unlist(data[1, ]))
  
  # Remove the first row (now used as column names)
  data <- data[-1, ]
  
  # Rename month and year columns
  colnames(data)[1] <- "year"
  colnames(data)[2] <- "month"
  
  # Clear row headings
  rownames(data) <- NULL
  
  # Reformat data to a dataframe format
  data <- as.data.frame(data)
  
  # Define year tracker column
  data$YearTracker <- NA
  data <- data |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        TRUE ~ "NA"
      )
    )
  
  # Create a monthid column based on the month name
  data_v1 <- data |>
    mutate(monthId = case_when(
      grepl("Jan", month) ~ "01",
      grepl("Feb", month) ~ "02",
      grepl("Mar", month) ~ "03",
      grepl("Apr", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("Jun", month) ~ "06",
      grepl("Jul", month) ~ "07",
      grepl("Aug", month) ~ "08",
      grepl("Sep", month) ~ "09",
      grepl("Oct", month) ~ "10",
      grepl("Nov", month) ~ "11",
      grepl("Dec", month) ~ "12",
      TRUE ~ "NA"
    ))
  
  for (i in 1:nrow(data_v1)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data_v1[i, 1])) {
      data_v1$YearTracker[i:nrow(data_v1)] <- data_v1[i, 1]
    }
  }
  
  data_v1 <- data_v1 |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  table2_amt <- data_v1 |>
    select(TIME_PERIOD, OBS_STATUS, 1:97, `01-98`) |>
    select(-c(month, year)) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE2(1.0)",
           FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
           REF_AREA = "WS",
           COUNTERPART_AREA = "_T",
           TRANSFORMATION = "N",
           UNIT_MEASURE = "WST",
           TRADE_FLOW = "X",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    ) |>
    rename(`HS` = `01-98`)
  
  #Drop last three records
  table2_amt <- table2_amt[1:(nrow(table2_amt) - 3), ]
  
  table2 <- table2_amt
  
  imts_tab2 <- table2 |>
    select(DATAFLOW, FREQ, REF_AREA, COUNTERPART_AREA, TRANSFORMATION, TRADE_FLOW, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, TIME_PERIOD, c(1:`HS`))
  
  imts_tab2 <- imts_tab2 |>
    mutate(across(`1`:`HS`, ~ round(as.numeric(.), 1)))
  
  # Renaming column names to include HS_ infront fot the code numbers
  colnames(imts_tab2) <- sapply(colnames(imts_tab2), function(name) {
    if (name == "_T") {
      return(name)  # leave "_T" unchanged
    } else if (grepl("^[0-9]$", name)) {
      paste0("HS_0", name)  # e.g., "1" -> "HS_01"
    } else if (grepl("^[0-9]+$", name)) {
      paste0("HS_", name)   # e.g., "10" -> "HS_10"
    } else {
      return(name)
    }
  })
  
  return(imts_tab2)
  
}


#### ***************************** IMTS table3 processing function ********************************************* ####

imts_table3 <- function(imts_tab3){
  data <- read_excel(destfile, sheet = "Table 3")
  
  # Transpose the read in table
  data <- t(data)
  
  # Drop columns 1 and 2 and row 2
  data <- data[, -c(1,2)]
  
  # Replace the column headings with the first record
  colnames(data) <- as.character(unlist(data[1, ]))
  
  # Remove the first row (now used as column names)
  data <- data[-c(1,2), ]
  
  # Drop last 5 records
  data <- data[1:(nrow(data) - 5), ]
  
  # Clear row headings
  rownames(data) <- NULL
  
  # Drop columns last 10 columns
  data <- data[, 1:(ncol(data) - 10)]
  
  # Reformat data to a dataframe format and rename the columns
  data <- as.data.frame(data)
  names(data)[ncol(data)] <- "SITC"
  colnames(data)[1] <- "year"
  colnames(data)[2] <- "month"
  
  # Define year tracker column
  data$YearTracker <- NA
  data <- data |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        TRUE ~ "NA"
      )
    )
  
  # Create a monthid column based on the month name
  data <- data |>
    mutate(monthId = case_when(
      grepl("Jan", month) ~ "01",
      grepl("Feb", month) ~ "02",
      grepl("Mar", month) ~ "03",
      grepl("Apr", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("Jun", month) ~ "06",
      grepl("Jul", month) ~ "07",
      grepl("Aug", month) ~ "08",
      grepl("Sep", month) ~ "09",
      grepl("Oct", month) ~ "10",
      grepl("Nov", month) ~ "11",
      grepl("Dec", month) ~ "12",
      TRUE ~ "NA"
    ))
  
  for (i in 1:nrow(data)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data[i, 1])) {
      data$YearTracker[i:nrow(data)] <- data[i, 1]
    }
  }
  
  data <- data |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  table3_amt <- data |>
    select(TIME_PERIOD, OBS_STATUS, `0`:`SITC`) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE3(1.0)",
           FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
           REF_AREA = "WS",
           COUNTERPART_AREA = "_T",
           TRANSFORMATION = "N",
           UNIT_MEASURE = "WST",
           TRADE_FLOW = "X",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    )
  
  # Re-ordeing the columns
  imts_tab3 <- table3_amt |>
    select(DATAFLOW, FREQ, REF_AREA, COUNTERPART_AREA, TRANSFORMATION, TRADE_FLOW, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, TIME_PERIOD, c(`0`:`SITC`))
  
  # Round off the number into 1 decimal place
  imts_tab3 <- imts_tab3 |>
    mutate(across(`0`:`SITC`, ~ round(as.numeric(.), 1)))
  
  # Renaming column names to include HS_ infront fot the code numbers
  colnames(imts_tab3) <- sapply(colnames(imts_tab3), function(name) {
    if (name == "SICT") {
      return(name)  # leave "_T" unchanged
    } else if (grepl("^[0-9]$", name)) {
      paste0("SITC_", name)  # e.g., "1" -> "HS_01"
    } else {
      return(name)
    }
  })
  
  return(imts_tab3)
  
}


  


#### ***************************** IMTS table4 processing function ********************************************* ####

imts_table4_all <- function(table4_all){
  
  data <- read_excel(destfile, sheet = "Table 4")
  
  # Transpose data
  data_t <- t(data)
  # Remove row headings
  rownames(data_t) <- NULL
  
  # Drop columns 1 and 2 and row 2
  data_t <- data_t[, -c(1,2)]
  
  # Replace the column headings with the first record
  colnames(data_t) <- as.character(unlist(data_t[1, ]))
  
  # Remove the first row (now used as column names)
  data_t <- data_t[-1, ]
  
  data_t <- as.data.frame(data_t)
  data_t$YearTracker <- NA
  
  for (i in 1:nrow(data_t)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data_t[i, 1])) {
      data_t$YearTracker[i:nrow(data_t)] <- data_t[i, 1]
    }
  }
  
  colnames(data_t)[1] <- "year"
  colnames(data_t)[2] <- "month"
  
  data_t <- data_t[, !is.na(colnames(data_t)) & colnames(data_t) != ""]
  
  # Create a monthid column based on the month name
  data_t <- data_t |>
    mutate(monthId = case_when(
      grepl("Jan", month) ~ "01",
      grepl("Feb", month) ~ "02",
      grepl("Mar", month) ~ "03",
      grepl("Apr", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("Jun", month) ~ "06",
      grepl("Jul", month) ~ "07",
      grepl("Aug", month) ~ "08",
      grepl("Sep", month) ~ "09",
      grepl("Oct", month) ~ "10",
      grepl("Nov", month) ~ "11",
      grepl("Dec", month) ~ "12",
      TRUE ~ "NA"
    ))  
  
  data_t <- data_t |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  data_t <- data_t |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        TRUE ~ "NA"
      )
    )
  
  table4_amt <- data_t |>
    select(TIME_PERIOD, OBS_STATUS, `Africa`:`Total`) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE4(1.0)",
           FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
           REF_AREA = "WS",
           TRADE_FLOW = "X",
           TRANSFORMATION = "N",
           COMMODITY = "_T",
           UNIT_MEASURE = "WST",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    )
  
    #Drop last 3 columns
    table4_amt <- table4_amt[1:(nrow(table4_amt) - 3), ]
    #Drop blank column and rename other region and other country
    table4_all <- table4_amt |>
      select(- 'Top Countries') |>
      rename(`Other regions` = Others,
             `Other countries` = `Others.1`
      )
    
    # Reorder the columns
    table4_all <- table4_all |>
      select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, TRANSFORMATION, TIME_PERIOD, UNIT_MEASURE, COMMODITY, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, everything())
    
    return(table4_all)
}

#### ***************************** IMTS table5 processing function ********************************************* ####

imts_table5 <- function(imts_tab5){
  
data <- read_excel(destfile, sheet = "Table 5")

# Transpose the read in table
data <- t(data)

# Drop columns 1 and 2 and row 2 and 3
data <- data[, -c(1,2)]
data <- data[-c(2,3), ]

# Drop last 9 columns
data <- data[, 1:(ncol(data) - 9)]

# Replace the column headings with the first record
colnames(data) <- as.character(unlist(data[1, ]))

# Remove the first row (now used as column names)
data <- data[-1, ]

# Rename month and year columns
colnames(data)[1] <- "year"
colnames(data)[2] <- "month"

# Clear row headings
rownames(data) <- NULL

# Reformat data to a dataframe format
data <- as.data.frame(data)

#Drop last three records
data <- data[1:(nrow(data) - 3), ]

# Define year tracker column
data$YearTracker <- NA
data <- data |>
  mutate(
    OBS_STATUS = case_when(
      grepl("\\(P\\)", month) ~ "P",
      grepl("\\(R\\)", month) ~ "R",
      grepl("\\(\\*\\)", month) ~ "*",
      TRUE ~ "NA"
    )
  )

# Create a monthid column based on the month name
data <- data |>
  mutate(monthId = case_when(
    grepl("Jan", month) ~ "01",
    grepl("Feb", month) ~ "02",
    grepl("Mar", month) ~ "03",
    grepl("Apr", month) ~ "04",
    grepl("Apr", month) ~ "04",
    grepl("May", month) ~ "05",
    grepl("Jun", month) ~ "06",
    grepl("Jul", month) ~ "07",
    grepl("Aug", month) ~ "08",
    grepl("Sep", month) ~ "09",
    grepl("Oct", month) ~ "10",
    grepl("Nov", month) ~ "11",
    grepl("Dec", month) ~ "12",
    TRUE ~ "NA"
  ))

for (i in 1:nrow(data)) {
  # If the Period column is a full year (4 digits or with (P)), save it
  if (grepl("^[0-9]{4}", data[i, 1])) {
    data$YearTracker[i:nrow(data)] <- data[i, 1]
  }
}

data <- data |>
  mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))

table5_amt <- data |>
  select(TIME_PERIOD, OBS_STATUS, `1`:`01-98`) |>
  mutate(DATAFLOW = "SBS:DF_IMTS_TABLE5(1.0)",
         FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
         REF_AREA = "WS",
         COUNTERPART_AREA = "_T",
         TRANSFORMATION = "N",
         UNIT_MEASURE = "WST",
         TRADE_FLOW = "M",
         UNIT_MULT = 3,
         COMMENT = "",
         DECIMALS = 1
  ) |>
  rename(HS = `01-98`)

imts_tab5 <- table5_amt |>
  select(DATAFLOW, FREQ, REF_AREA, COUNTERPART_AREA, TRANSFORMATION, TRADE_FLOW, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, TIME_PERIOD, c(1:HS))

imts_tab5 <- imts_tab5 |>
  mutate(across(`1`:`HS`, ~ round(as.numeric(.), 1)))

# Renaming column names to include HS_ infront fot the code numbers
colnames(imts_tab5) <- sapply(colnames(imts_tab5), function(name) {
  if (name == "HS") {
    return(name)  # leave "_T" unchanged
  } else if (grepl("^[0-9]$", name)) {
    paste0("HS_0", name)  # e.g., "1" -> "HS_01"
  } else if (grepl("^[0-9]+$", name)) {
    paste0("HS_", name)   # e.g., "10" -> "HS_10"
  } else {
    return(name)
  }
})

return(imts_tab5)

}
#### ***************************** IMTS table6 processing function ********************************************* ####

imts_table6 <- function(imts_tab6){
  
  data <- read_excel(destfile, sheet = "Table 6")
  
  # Transpose the read in table
  data <- t(data)
  
  # Drop columns 1 and 2 and row 2
  data <- data[, -c(1,2)]
  data <- data[-c(2), ]
  
  # Replace the column headings with the first record
  colnames(data) <- as.character(unlist(data[1, ]))
  
  # Remove the first row (now used as column names)
  data <- data[-1, ]
  
  # Drop last 3 columns
  data <- data[, 1:(ncol(data) - 3)]
  
  # Rename month and year columns and last column to _T
  colnames(data)[1] <- "year"
  colnames(data)[2] <- "month"
  colnames(data)[ncol(data)] <- "SITC"
  
  # Clear row headings
  rownames(data) <- NULL
  
  # Reformat data to a dataframe format
  data <- as.data.frame(data)
  
  # Define year tracker column
  data$YearTracker <- NA
  data <- data |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        TRUE ~ "NA"
      )
    )
  
  # Create a monthid column based on the month name
  data_v1 <- data |>
    mutate(monthId = case_when(
      grepl("Jan", month) ~ "01",
      grepl("Feb", month) ~ "02",
      grepl("Mar", month) ~ "03",
      grepl("Apr", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("Jun", month) ~ "06",
      grepl("Jul", month) ~ "07",
      grepl("Aug", month) ~ "08",
      grepl("Sep", month) ~ "09",
      grepl("Oct", month) ~ "10",
      grepl("Nov", month) ~ "11",
      grepl("Dec", month) ~ "12",
      TRUE ~ "NA"
    ))
  
  for (i in 1:nrow(data_v1)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data_v1[i, 1])) {
      data_v1$YearTracker[i:nrow(data_v1)] <- data_v1[i, 1]
    }
  }
  
  data_v1 <- data_v1 |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  table6_amt <- data_v1 |>
    select(TIME_PERIOD, OBS_STATUS, `0`:`SITC`) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE3(1.0)",
           FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
           REF_AREA = "WS",
           COUNTERPART_AREA = "_T",
           TRANSFORMATION = "N",
           UNIT_MEASURE = "WST",
           TRADE_FLOW = "M",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    )
  
  # Drop the last 3 records including the percentage records
  table6_amt <- table6_amt[1:(nrow(table6_amt) - 3), ]
  
  imts_tab6 <- table6_amt |>
    select(DATAFLOW, FREQ, REF_AREA, COUNTERPART_AREA, TRANSFORMATION, TRADE_FLOW, UNIT_MEASURE, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, TIME_PERIOD, c(`0`:`SITC`))
  
  # Renaming column names to include HS_ infront fot the code numbers
  colnames(imts_tab6) <- sapply(colnames(imts_tab6), function(name) {
    if (name == "SITC") {
      return(name)  # leave "_T" unchanged
    } else if (grepl("^[0-9]$", name)) {
      paste0("SITC_", name)  # e.g., "1" -> "HS_01"
    } else {
      return(name)
    }
  })
  
  return(imts_tab6)
  
}

#### ***************************** IMTS table7 processing function ********************************************* ####

imts_table7_all <- function(table7_all){
  
  data <- read_excel(destfile, sheet = "Table 7")
  
  # Transpose data
  data_t <- t(data)
  # Remove row headings
  rownames(data_t) <- NULL
  
  # Drop columns 1 and 2 and row 2
  data_t <- data_t[, -c(1,2)]
  
  # Remove the first three row
  data_t <- data_t[-c(1:3), ]
  
  # Replace the column headings with the first record
  colnames(data_t) <- as.character(unlist(data_t[1, ]))
  
  # Remove the first row (now used as column names)
  data_t <- data_t[-1, ]
  
  data_t <- as.data.frame(data_t)
  data_t$YearTracker <- NA
  
  for (i in 1:nrow(data_t)) {
    # If the Period column is a full year (4 digits or with (P)), save it
    if (grepl("^[0-9]{4}", data_t[i, 1])) {
      data_t$YearTracker[i:nrow(data_t)] <- data_t[i, 1]
    }
  }
  
  colnames(data_t)[1] <- "year"
  colnames(data_t)[2] <- "month"
  
  data_t <- data_t[, !is.na(colnames(data_t)) & colnames(data_t) != ""]
  
  # Create a monthid column based on the month name
  data_t <- data_t |>
    mutate(monthId = case_when(
      grepl("Jan", month) ~ "01",
      grepl("Feb", month) ~ "02",
      grepl("Mar", month) ~ "03",
      grepl("Apr", month) ~ "04",
      grepl("Apr", month) ~ "04",
      grepl("May", month) ~ "05",
      grepl("Jun", month) ~ "06",
      grepl("Jul", month) ~ "07",
      grepl("Aug", month) ~ "08",
      grepl("Sep", month) ~ "09",
      grepl("Oct", month) ~ "10",
      grepl("Nov", month) ~ "11",
      grepl("Dec", month) ~ "12",
      TRUE ~ "NA"
    ))  
  
  data_t <- data_t |>
    mutate(TIME_PERIOD = ifelse(monthId != "NA", paste0(YearTracker, "-", monthId), YearTracker))
  
  data_t <- data_t |>
    mutate(
      OBS_STATUS = case_when(
        grepl("\\(P\\)", month) ~ "P",
        grepl("\\(R\\)", month) ~ "R",
        grepl("\\(\\*\\)", month) ~ "*",
        TRUE ~ "NA"
      )
    )
  
  table7_amt <- data_t |>
    select(TIME_PERIOD, OBS_STATUS, `Asia`:`Total`) |>
    mutate(DATAFLOW = "SBS:DF_IMTS_TABLE7(1.0)",
           FREQ = ifelse(length(TIME_PERIOD) > 5, "M", "A"),
           REF_AREA = "WS",
           TRANSFORMATION = "N",
           COMMODITY = "_T",
           UNIT_MEASURE = "WST",
           TRADE_FLOW = "M",
           UNIT_MULT = 3,
           COMMENT = "",
           DECIMALS = 1
    )
  
    # Drop last three records
    table7_amt <- table7_amt[1:(nrow(table7_amt) - 3), ]
    table7_all <- table7_amt |>
      select(- 'Top Countries') |>
      rename(`Other regions` = Others,
             `Other countries` = `Others.1`
      )
    
    # Reorder the columns
    table7_all <- table7_all |>
      select(DATAFLOW, FREQ, REF_AREA, TRADE_FLOW, TRANSFORMATION, TIME_PERIOD, UNIT_MEASURE, COMMODITY, UNIT_MULT, OBS_STATUS, COMMENT, DECIMALS, everything())
    
    return(table7_all)
  }
  