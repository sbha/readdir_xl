library(tidyverse)
library(readxl)

dir_path <- "~/test_dir/"         # target directory path where the xlsx files are located. 
re_file <- "^test[0-9]\\.xlsx"    # regex pattern to match the file name format, in this case 'test1.xlsx', 'test2.xlsx' etc, but could simply be 'xlsx'.

read_sheets <- function(dir_path, file){
  xlsx_file <- paste0(dir_path, file)
  xlsx_file %>%
    excel_sheets() %>%
    set_names() %>%
    map_df(read_xlsx, path = xlsx_file, .id = 'sheet_name') %>% 
    mutate(file_name = file) %>% 
    select(file_name, sheet_name, everything())
}

df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~ read_sheets(dir_path, .))

# reformat column names
read_sheets <- function(dir_path, file){
  xlsx_file <- paste0(dir_path, file)
  xlsx_file %>%
    excel_sheets() %>%
    set_names() %>%
    map_df(read_xlsx, path = xlsx_file, .id = 'sheet_name') %>% 
    mutate(file_name = file) %>% 
    select(file_name, sheet_name, everything()) %>% 
    rename_all(~tolower(.)) %>% 
    rename_all(~str_replace_all(., '\\s+', '_'))
}

# reformat column names 
# filter data
# create new column
read_sheets <- function(dir_path, file){
  xlsx_file <- paste0(dir_path, file)
  xlsx_file %>%
    excel_sheets() %>%
    set_names() %>%
    map_df(read_xlsx, path = xlsx_file, .id = 'sheet_name') %>% 
    mutate(file_name = file) %>% 
    select(file_name, sheet_name, everything()) %>% 
    rename_all(~tolower(.)) %>% 
    rename_all(~str_replace_all(., '\\s+', '_')) %>% 
    filter(a > 1) %>% 
    mutate(a_plus_b = a + b)
}