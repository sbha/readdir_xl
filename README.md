## Introduction

R can quickly and easily aggregrate data from multiple sheets from multiple Excel workbooks into a single data frame. A custom function combining functions from the `tidyverse` and `readxl` packages provides a template that can be adapted and expanded to handle many different data specific needs. 


## Usage

Import the packages that have the functions that do all the heavy lifting:

```
#install.packages(tidyverse)
#install.packages(readxl)
library(tidyverse)
library(readxl)
```

Define the target directory as a variable. This is the path to the directory where the Excel files are stored:

`dir_path <- "~/path/to/test_dir/"`  


In this example all the Excel files we're interested in have the same naming convention, so we can use a regular expression to ensure we're reading only the files we need. This is useful if there are files in the same directory that we don't want to read. In this case the files are named `sample_2019-01-09.xlsx`, `sample_2019-01-15.xlsx` and so on, but we could simply use `xlsx` if we knew we needed every Excel file, or leave that arguement empty if we knew the directory contained only Excel files. The regular expression can be defined as a variable:

`re_file <- "^sample_[0-9]{4}-[0-9]{2}-[0-9]{2}\\.xlsx" `    

This regular expression will match file names that begin with `sample_`, a date formatted as `xxxx-xx-xx`, and end with the file extension `.xlsx`. You can check that this matches all the files you're expecting with:

`list.files(dir_path, re_file)`

Next we can write the custom function that will read the individual Excel files and combine the different sheets within each workbook into a single data frame. The two function inputs will be the path to the files, which was defined earlier with `dir_path` and then the individual file name itself. The function will read the individual file using `excel_sheets()`, get the individual sheet names with `set_names()` and set them as a variable called `sheet_name`, and then import the data from each sheet using `read_xlsx()` and `map_df()`. `mutate()` is used to set the file name as a variable so that we know which file the data comes from, and then `select()` rearranges the columns so that the file name and sheet names are the first two columns and `everything()` includes the remaining of the columns without having to name them specifically: 

```
read_sheets <- function(dir_path, file){
  xlsx_file <- paste0(dir_path, file)
  xlsx_file %>%
    excel_sheets() %>%
    set_names() %>%
    map_df(read_xlsx, path = xlsx_file, .id = 'sheet_name') %>% 
    mutate(file_name = file) %>% 
    select(file_name, sheet_name, everything())
}
```

With the custom function defined, we can now use it to aggregate all the sheets from all the individual files matching the regular expession in the directory into a single data frame:

```
df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~read_sheets(dir_path, .))
```

Just like when we checked to make sure that we were reading from the correct directory and the expected file names, we'll use `list.files()` to get the names for all the individual file names we're interested in importing. The resulting file names will be used as the file name argument along side `dir_path` in the custom function. Using `map_df()` we'll apply the `read_sheets()` custom function to each file name returned by `list.files()` and combine them into a single data frame:  

``` 
> df_xl
# A tibble: 88 x 7
   file_name              sheet_name cat   `Col 1` `Col 2` `Col 3` `Col 4`
   <chr>                  <chr>      <chr>   <dbl>   <dbl>   <dbl>   <dbl>
 1 sample_2019-01-09.xlsx Sheet1     b           5       9       8      NA
 2 sample_2019-01-09.xlsx Sheet1     b           8       1       9      NA
 3 sample_2019-01-09.xlsx Sheet1     c           9       2       1      NA
 4 sample_2019-01-09.xlsx Sheet1     a           1       5       1      NA
 5 sample_2019-01-09.xlsx Sheet1     b           2       8       2      NA
 6 sample_2019-01-09.xlsx Sheet1     b           5       4       5      NA
 7 sample_2019-01-09.xlsx Sheet1     c           8       8       8      NA
 8 sample_2019-01-09.xlsx Sheet1     a           9       9       2      NA
 9 sample_2019-01-09.xlsx Sheet1     b           1       1       5      NA
10 sample_2019-01-09.xlsx Sheet1     b           2       2       8      NA
```

To get a better sense of everything in the data frame, we can use `count()` to see the number of observations by file and sheet name:

```
> df_xl %>% count(file_name, sheet_name)
# A tibble: 8 x 3
  file_name              sheet_name     n
  <chr>                  <chr>      <int>
1 sample_2019-01-09.xlsx Sheet1         6
2 sample_2019-01-09.xlsx Sheet2         7
3 sample_2019-01-15.xlsx Sheet1         6
4 sample_2019-01-15.xlsx Sheet2         5
5 sample_2019-01-15.xlsx Sheet3         3
6 sample_2019-02-02.xlsx Sheet1         9
7 sample_2019-02-02.xlsx Sheet2         5
8 sample_2019-02-02.xlsx Sheet3         2
```

In this example, not every file has the same sheets or columns. The files need not have exactly the same structure; only `sample_2019-01-15.xlsx` has `Col 4`, and `sample_2019-01-09.xlsx` has only two sheets. 

This custom function aggregates every file, but we can improve the output by extending the function to handle more specific needs of this data. For our example, the column names can be reformatted so that they are always in a more `R` friendly format. In the modified function below, all column names are converted to lower case and spaces are replaced with a single underscore using the `rename_all()` and `str_replace_all()` functions:

```
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
```

Going further, if the original Excel files contains data you don't need, you could remove it using `filter()` or by negating specific columns with `select()`. Or a column could be created using `mutate()` with something like the following:

```
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
```

In this expanding function, we're removing any rows where column `cat` equals `b`, dropping column `col_3`, and then creating a new column that adds columns `col_1` and `col_2`. Running the same Excel files through the updated custom function gives us a different data frame, more specific to our needs:

```
df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~ read_sheets(dir_path, .))
  
> df_xl
# A tibble: 48 x 7
   file_name              sheet_name cat   col_1 col_2 col_1_plus_col_2 col_4
   <chr>                  <chr>      <chr> <dbl> <dbl>            <dbl> <dbl>
 1 sample_2019-01-09.xlsx Sheet1     c         9     2               11    NA
 2 sample_2019-01-09.xlsx Sheet1     a         1     5                6    NA
 3 sample_2019-01-09.xlsx Sheet1     c         8     8               16    NA
 4 sample_2019-01-09.xlsx Sheet1     a         9     9               18    NA
 5 sample_2019-01-09.xlsx Sheet1     c         4     5                9    NA
 6 sample_2019-01-09.xlsx Sheet1     a        NA     8               NA    NA
 7 sample_2019-01-09.xlsx Sheet2     a         5     4                9    NA
 8 sample_2019-01-09.xlsx Sheet2     c         9     4               13    NA
 9 sample_2019-01-09.xlsx Sheet2     a         1     1                2    NA
10 sample_2019-01-09.xlsx Sheet2     c         8     5               13    NA
# ... with 38 more rows
```

The modifications with this example might seem trivial and something that can be done seperately after the data has been aggregrated, which they can, of course, but they quickly become useful when dealing with a large number of files that might get near a machine's memory limits or as part of a process that will be repeated. 

If there are many files in the directory and we only want a specific smaller subset, we can be a bit more selective with the files we import. In our example, the sample files have a date in the name and we can use that to read only files from a given month or before a given date. To get only those select files, we can create a data frame from the file names that match our naming structure and the `list.files()` function, extract the date from the file name with `str_extract()` and then filter to the month we need:

```
df_dir <- data_frame(file = list.files(dir_path, re_file)) %>% 
  mutate(date = as.Date(str_extract(file, '[0-9]{4}-[0-9]{2}-[0-9]{2}'))) %>% 
  mutate(year_mon = format(date, '%Y-%m')) %>% 
  filter(year_mon == '2019-01') 
```  

In this example, we can say we're only interested in files from January 2019. We would then read the files using:

```
df_xl <- df_dir$file %>% 
  map_df(~read_sheets(dir_path, .))
```
  


Finally, if the files aren't `.xlsx`, a similar method can be used for `.csv` or other delimited files using functions from the `readr` package:

```
read_csv_files <- function(dir_path, file){
  read_csv(paste0(dir_path, file)) %>% 
    mutate(file_name = file) %>% 
    select(file_name, everything())
}

df_csv <- list.files(dir_path, '\\.csv') %>% 
  map_df(~ read_csv_files(dir_path, .))
```




### Summary

Using functions from the `tidyverse` and `readxl` packages we can define a custom function that can combine multiple Excel files into a single data frame. We can modifiy and extend that custom function further to handle more specific needs for formatting, reorganizing, and combining the data. This is a process that can be adapted as needed. 
