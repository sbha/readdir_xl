
```{r setup, include=FALSE}
knitr::opts_chunk$set(echo = TRUE)
```

## Introduction

R can quickly and easily combine multiple sheets from multiple Excel files into a single data frame. Here's a `tidyverse` and `readxl` driven custom function that returns a data frame with columns for file and sheet names for each file.


## Usage

Import the packages:

```
library(tidyverse)
library(readxl)
```

Define the target directory as a variable. This is the path to the directory where the Excel files are stored:

`dir_path <- "~/path/to/test_dir/"`  


In this example all the Excel files we're interested in have the same naming convention. We can use a regular expression to ensure we're reading only the files we need. This is useful if there are files in the same directory that we don't want to read. In this case the files are named `test1.xlsx`, `test2.xlsx` and so on, but we could simply use `xlsx` if we knew we needed every Excel file. The regular expression can be defined in a variable:

`re_file <- "^test[0-9]\\.xlsx"`    

You can check that this matches alll the files you're expecting with:

`list.files(dir_path, re_file)`

Next we can define the custom function that will read the individual Excel files. The two function inputs will be the path to the files, which was defined earlier with `dir_path` and then the individual file name itself. The function will read the individual file using `excel_sheets()`, get the individual sheet names with `set_names()`, and then import the data from each sheet using `read_xlsx()` and `map_df()`. `mutate()` is used to define the file name and then `select()` rearranges the columns so that the file name and sheet names are the first two columns and `everything()` includes the rest of the columns without having to name them specifically: 

```{r message = FALSE, results='hide'}
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

With the custom function defined, we can then use it to combine the individual files into a single data frame. We'll again use the `map_df()` function from the `purrr` package:

```
df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~ read_sheets(dir_path, .))
```

Just like when we checked to make sure that we were reading from the correct directory and the correct file names, we'll use `list.files()` to get all the individual file names. Then using `map_df()` we'll apply our `read_sheets()` function to each Excel file in that directory and combine them into a single data frame:  

``` 
# A tibble: 15 x 5
   file_name  sheet_name  col1  col2  col3
   <chr>      <chr>      <dbl> <dbl> <dbl>
 1 test1.xlsx Sheet1         1     2     4
 2 test1.xlsx Sheet1         3     2     3
 3 test1.xlsx Sheet1         2     4     4
 4 test1.xlsx Sheet2         3     3     1
 5 test1.xlsx Sheet2         2     2     2
 6 test1.xlsx Sheet2         4     3     4
 7 test2.xlsx Sheet1         1     3     5
 8 test2.xlsx Sheet1         4     4     3
 9 test2.xlsx Sheet1         1     2     2
10 test3.xlsx Sheet1         3     9    NA
11 test3.xlsx Sheet1         4     7    NA
12 test3.xlsx Sheet1         5     3    NA
13 test3.xlsx Sheet2         1     3     4
14 test3.xlsx Sheet2         2     5     9
15 test3.xlsx Sheet2         4     3     1
```

In this example, not every file has the same sheets or columns. File `test2.xlsx` has only one sheet, and `Sheet1` in `test3.xlsx` has only the first two columns.  


This custom function defined above can be modified to handle specific needs. For example, the column names can converted to lower case and spaces replaces with underscores:

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

If the original Excel files contains data you don't need, you could remove it using `filter()` or if a field is missing you could create it using `mutate()` with something like the following:

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
    filter(a > 1) %>% 
    mutate(a_plus_b = a + b)
}
```

In this case, we're removing any rows where column a is 1 or less and then creating a new column that adds columns a and b:

```
df_xl <- list.files(dir_path, re_file) %>% 
  map_df(~ read_sheets(dir_path, .))
  
# A tibble: 10 x 6
   file_name  sheet_name     a     b c     a_plus_b
   <chr>      <chr>      <dbl> <dbl> <chr>    <dbl>
 1 test1.xlsx Sheet1         2     5 b            7
 2 test1.xlsx Sheet1         3     6 c            9
 3 test2.xlsx Sheet1         4    11 d           15
 4 test2.xlsx Sheet1         5    10 e           15
 5 test2.xlsx Sheet1         6     9 f           15
 6 test2.xlsx Sheet1         7     8 g           15
 7 test3.xlsx Sheet1         8    11 d           19
 8 test3.xlsx Sheet1         7    11 a           18
 9 test3.xlsx Sheet1         6    10 f           16
10 test3.xlsx Sheet1         5    10 d           15
```

### Summary

Using functions from the `tidyverse` and `readxl` packages we can define a custom function that can combine multiple Excel files into a single data frame. We can modifiy that custom function further to handle more specific needs of the data. 
