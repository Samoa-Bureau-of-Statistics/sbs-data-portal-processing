
# Load libraries.
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(janitor)

# Map working directory
repository <- file.path(dirname(rstudioapi::getSourceEditorContext()$path))
setwd(repository)

#data <- read_excel("../data/cpiData.xlsx", sheet = "DF_CPI_TABLE1")

# Path to your Excel file
file_path <- "data/cpiData.xlsx"

# Get list of sheet names
sheet_names <- excel_sheets(file_path)

i = 1
numsheet = length(sheet_names)

while (i <= numsheet) {
  message("📄 Processing sheet: ", sheet_names[i])
  if(sheet_names[i] == "DF_CPI_TABLE3"){
    
    table <- read_excel(file_path, sheet = sheet_names[i])
    table$ITEM <- iconv(table$ITEM, from = "", to = "UTF-8", sub = "")
    table$ITEM <- tolower(trimws(table$ITEM))
    
    cpi_items <- read.csv("other/cpi_items.csv")
    cpi_items$ITEM <- iconv(cpi_items$ITEM, from = "", to = "UTF-8", sub = "")
    cpi_items$ITEM <- tolower(trimws(cpi_items$ITEM))
    
    diff_table_not_cpi_items <- table[!table$ITEM %in% cpi_items$ITEM, ]
    
    table <- merge(table, cpi_items) |>
      select(-c("ITEM", "Unit")) |>
      select(ITEM_Code, everything())
    
    table_new <- as.data.frame(t(table))
    colnames(table_new) <- as.character(unlist(table_new[1, ]))
    table_new <- table_new[-1, ]
    
    # Ensure unique names
    colnames(table_new) <- make.names(colnames(table_new), unique = TRUE)
    
    table_new <- cbind(TIME_PERIOD = rownames(table_new), table_new)
    rownames(table_new) <- NULL
    
    table_new <- table_new |>
      mutate(
        DATAFLOW = "SBS:DF_CPI_TABLE3(1.0)",
        FREQ = "M",
        REF_AREA = "WS",
        INDICATOR_TYPE = "ARPSEL",
        INDICATOR = ifelse(TIME_PERIOD == "Wt", "WGT", "PRI"),
        UNIT_MEASURE = case_when(
          grepl("\\(A\\)", TIME_PERIOD) ~ "PT",
          grepl("\\(M\\)", TIME_PERIOD) ~ "PT",
          TIME_PERIOD == "Wt" ~ "N",
          TRUE ~ "WST"
        ),
        BASE_PER = ifelse(TIME_PERIOD == "Wt", "", "Average Prices February 2016 = 100"),
        OBS_STATUS = "",
        COMMENT = "",
        DECIMALS = 2,
        TRANSFORMATION = case_when(
          grepl("\\(A\\)", TIME_PERIOD) ~ "G1Y",
          grepl("\\(M\\)", TIME_PERIOD) ~ "G1M",
          TRUE ~ "N"
        ),
        TIME_PERIOD = trimws(gsub("\\((A|M)\\)", "", TIME_PERIOD))
      )
    
    table_new <- table_new |>
      mutate(TIME_PERIOD = ifelse(TIME_PERIOD == "Wt", "2020-01-01/P6Y", TIME_PERIOD)) |>
      select(DATAFLOW, FREQ, REF_AREA, INDICATOR_TYPE, INDICATOR, TRANSFORMATION, TIME_PERIOD, UNIT_MEASURE, BASE_PER, OBS_STATUS, COMMENT, DECIMALS,everything())
    
    # Create table long
    
      table_long <- table_new |>
        pivot_longer(
          cols = -c(DATAFLOW:DECIMALS),
          names_to = "ITEM",
          values_to = "OBS_VALUE"
        )
      
    # Re-organizing the columns
      
    table_long_final <- table_long |>
      select(DATAFLOW, FREQ, REF_AREA, INDICATOR_TYPE, INDICATOR, ITEM, TRANSFORMATION, TIME_PERIOD, OBS_VALUE, UNIT_MEASURE, BASE_PER, OBS_STATUS, COMMENT, DECIMALS)
    
    #table_long_final[is.na(table_long_final)] <- ""
    
    myfile <- paste0("output/cpi/", sheet_names[i],".csv")
    
    # Export if desired
    write.csv(table_long_final, myfile , row.names = FALSE)
    
  }else{
    
  # Process tables
  table <- read_excel(file_path, sheet = sheet_names[i])
  # Reshape from wide to long format
  table_long <- table |>
    pivot_longer(
      cols = -c(DATAFLOW:DECIMALS),
      names_to = "ITEM",
      values_to = "OBS_VALUE"
    )
  
  # Keep only CPI index observations
  df_idx <- table_long |>
    filter(TRANSFORMATION == "N", UNIT_MEASURE == "INDEX")
  
  #Convert TIME_PERIOD into a proper date
  df_idx <- df_idx %>%
    mutate(
      date = case_when(
        FREQ == "A" ~ ymd(paste0(TIME_PERIOD, "-01-01")),
        FREQ == "M" ~ ymd(paste0(TIME_PERIOD, "-01")),
        TRUE ~ NA
      )
    )
  
  # Calculate % changes by REF_AREA + ITEM + FREQ
  df_changes <- df_idx %>%
    arrange(REF_AREA, ITEM, date) %>%
    group_by(REF_AREA, ITEM, FREQ) %>%
    mutate(
      MCHG = (OBS_VALUE / lag(OBS_VALUE) - 1) * 100,
      ACHG = (OBS_VALUE / lag(OBS_VALUE, 12) - 1) * 100
    ) %>%
    ungroup()
  
  #Reshape into SDMX long format
  df_long <- df_changes %>%
    select(DATAFLOW, FREQ, REF_AREA, INDICATOR_TYPE, INDICATOR,
           TIME_PERIOD, UNIT_MEASURE, BASE_PER, COMMENT,
           ITEM, OBS_VALUE, MCHG, ACHG) %>%
    tidyr::pivot_longer(
      cols = c(OBS_VALUE, MCHG, ACHG),
      names_to = "TRANSFORMATION",
      values_to = "OBS_VALUE"
    ) %>%
    mutate(
      TRANSFORMATION = recode(TRANSFORMATION,
                              "OBS_VALUE" = "IDX",
                              "MCHG" = "G1M",
                              "ACHG" = "G1Y"),
      UNIT_MEASURE = case_when(
        TRANSFORMATION == "IDX" ~ "INDEX",
        TRUE ~ "PT"
      ),
      DECIMALS = 1,
      OBS_STATUS = ""
    )
  
  # Final tidy SDMX-style output
  final_sdmx <- df_long |>
    arrange(REF_AREA, ITEM, FREQ, TRANSFORMATION)
    
  final_sdmx <- final_sdmx |> 
    filter((TRANSFORMATION == "G1M" | TRANSFORMATION == "G1Y") & !is.na(OBS_VALUE)) |>
    mutate(TRANSFORMATION = ifelse(FREQ == "A" & TRANSFORMATION == "G1M", "G1Y", TRANSFORMATION))
    
  final_sdmx <- rbind(table_long, final_sdmx) |>
    select(DATAFLOW, FREQ, REF_AREA, INDICATOR_TYPE, INDICATOR, ITEM, TRANSFORMATION, TIME_PERIOD, OBS_VALUE, UNIT_MEASURE, BASE_PER, OBS_STATUS, COMMENT, DECIMALS)
  
  #final_sdmx[is.na(final_sdmx)] <- ""
  
  myfile <- paste0("output/cpi/", sheet_names[i],".csv")
  
  
  # Export if desired
  write.csv(final_sdmx, myfile , row.names = FALSE)
  
  } # End the if statement
  message("😎 Completed processing table", sheet_names[i])
  
  i <- i + 1
  
} # End the While loop

message("🎉 All sheets processed successfully!")
