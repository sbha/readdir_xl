#install.packages(tidyverse)
#install.packages(readxl)
library(tidyverse)
library(readxl)

dir_path <- "~/test_dir/"         # target directory path where the xlsx files are located. 
# dir_path <- '/Users/stuartharty/Documents/Data/test_dir/'

re_file <- "^sample_[0-9]{4}-[0-9]{2}-[0-9]{2}\\.xlsx"    # regex pattern to match the file name format, in this case 'test1.xlsx', 'test2.xlsx' etc, but could simply be 'xlsx'.

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
# drop a column
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
    select(-col_3) %>% 
    filter(cat != 'b') %>% 
    mutate(col_1_plus_col_2 = col_1 + col_2)
}

# select files by date
df_dir <- data_frame(file = list.files(dir_path, re_file)) %>% 
  mutate(date = as.Date(str_extract(file, '[0-9]{4}-[0-9]{2}-[0-9]{2}'))) %>% 
  mutate(year_mon = format(date, '%Y-%m')) %>% 
  filter(year_mon == '2019-01') 

df_xl <- df_dir$file %>% 
  map_df(~ read_sheets(dir_path, .))



# similar function for csv files
read_csv_files <- function(dir_path, file){
  read_csv(paste0(dir_path, file)) %>% 
    mutate(file_name = file) %>% 
    select(file_name, everything())
}

df_csv <- list.files(dir_path, '\\.csv') %>% 
  map_df(~ read_csv_files(dir_path, .))

