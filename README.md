## Introduction

R can quickly and easily combine multiple sheets from multiple Excel files into a single data frame. A custom function combining functions from the `tidyverse` and `readxl` packages provides a template that can be adapted and expanded to handle many different data specific situations. 


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


In this example all the Excel files we're interested in have the same naming convention. We can use a regular expression to ensure we're reading only the files we need. This is useful if there are files in the same directory that we don't want to read. In this case the files are named `sample_2019-01-09.xlsx`, `sample_2019-01-15.xlsx` and so on, but we could simply use `xlsx` if we knew we needed every Excel file. The regular expression can be defined as a variable:

`re_file <- "^sample_[0-9]{4}-[0-9]{2}-[0-9]{2}\\.xlsx" `    

This regular expression will match file names that begin with 'sample_', then a data, and finally ends with the file extension '.xlsx'. You can check that this matches all the files you're expecting with:

`list.files(dir_path, re_file)`

Next we can define the custom function that will read the individual Excel files and combine the different sheets within each workbook. The two function inputs will be the path to the files, which was defined earlier with `dir_path` and then the individual file name itself. The function will read the individual file using `excel_sheets()`, get the individual sheet names with `set_names()` and set them as a variable called `sheet_name`, and then import the data from each sheet using `read_xlsx()` and `map_df()`. `mutate()` is used to set the file name as a variable so that we know which file the data comes from, and then `select()` rearranges the columns so that the file name and sheet names are the first two columns and `everything()` includes the rest of the columns without having to name them specifically: 

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

With the custom function defined, we can now use it to combine all the sheets from all the individual files matching the regular expession in the directory into a single data frame. We'll again use the `map_df()`:

```
df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~ read_sheets(dir_path, .))
```

Just like when we checked to make sure that we were reading from the correct directory and the correct file names, we'll use `list.files()` to get all the individual file names we're interested in importing. These will be used as the file name argument along side the `dir_path` in our custom function. Then using `map_df()` we'll apply our `read_sheets()` custom function to read each Excel file in that directory and combine them into a single data frame:  

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

To get a better sense of everything in the data frame we can use `count()` to see the number of observations by file and sheet names:

```
> df_xl %>% count(file_name, sheet_name)
# A tibble: 9 x 3
  file_name              sheet_name     n
  <chr>                  <chr>      <int>
1 sample_2019-01-09.xlsx Sheet1        12
2 sample_2019-01-09.xlsx Sheet2        11
3 sample_2019-01-09.xlsx Sheet3         8
4 sample_2019-01-15.xlsx Sheet1        12
5 sample_2019-01-15.xlsx Sheet2         8
6 sample_2019-01-15.xlsx Sheet3         7
7 sample_2019-02-02.xlsx Sheet1        17
8 sample_2019-02-02.xlsx Sheet2         8
9 sample_2019-02-02.xlsx Sheet3         5
```

In this example, not every file has the same sheets or columns. The files need not have exactly the same format; only `sample_2019-01-15.xlsx` has `Col 4`, and `sample_2019-01-09.xlsx` has only two sheets.  

The custom function defined above can be modified to handle additional specific needs depending on the data being import. For example, the column names can be reformated so that they are always in a consistent format. In the modified function below, everything is converted to lower case and spaces are replaced with a single underscore using the `rename_all()` and `str_replace_all()` functions:

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

If the original Excel files contains data you don't need, you could remove it using `filter()` or negating it in a `select()` call. If a field is missing you could create it using `mutate()` with something like the following:

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

In this case, we're removing any rows where column `cat` is `b`, removing column `col_3`, and then creating a new column that adds columns `col_1` and `col_2`. Running the same Excel files through the updated custom function gives us a different data frame:

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

### Summary

Using functions from the `tidyverse` and `readxl` packages we can define a custom function that can combine multiple Excel files into a single data frame. We can modifiy that custom function further to handle more specific needs for formatting, reorganizing, and combining the data. 
