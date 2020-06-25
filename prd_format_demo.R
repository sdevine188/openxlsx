library(dplyr)
library(scales)
library(openxlsx)
library(readxl)
library(stringr)
library(purrr)
library(tidyr)


# setwd

# load prd_format function
source("prd_format.R")

# load table_1
table_1 <- read_excel(path = "example_table.xlsx", sheet = "table_1") %>% 
        mutate(Filed = comma(Filed), Grants = comma(Grants), Grant_rate = percent(Grant_rate), Denials = comma(Denials),
               Denial_rate = percent(Denial_rate), Closed = comma(Closed), Closed_rate = percent(Closed_rate)) %>%
        rename("Grant rate with looooooooooong title" = Grant_rate, "Denial rate\nwith\nlots\nof\nlinebreaks" = Denial_rate, "Closed rate" = Closed_rate, "Example text" = Example_text)
table_1

# load table_2
table_2 <- read_excel(path = "example_table.xlsx", sheet = "table_2") %>% 
        mutate(Filed = comma(Filed), Grants = comma(Grants), Grant_rate = percent(Grant_rate), Denials = comma(Denials),
               Denial_rate = percent(Denial_rate), Closed = comma(Closed), Closed_rate = percent(Closed_rate)) %>%
        rename("Grant rate" = Grant_rate, "Denial rate" = Denial_rate, "Closed rate" = Closed_rate, "Example text" = Example_text)
table_2


######################################################################################################################################


# create superscript_table
superscript_table_1 <- tibble(section = c("header", "body", "footnote"), row = c(1, 5, 1), col = c(2, 3, 1), 
                              superscript_text = c("1", "*", "a"), superscript_position = c(6, 2, 1), 
                              prior_text = c("Filed", "200,000", "This is footnote 1 for table 1."))
superscript_table_2 <- tibble(section = c("header", "body", "footnote"), row = c(1, 5, 1), col = c(2, 3, 1),
                              superscript_text = c("1", "*", "a"), superscript_position = c(6, 2, 1),  
                              prior_text = c("Filed", "200,000", "This is footnote 1 for table 2."))
superscript_table <- list(superscript_table_1, superscript_table_2)
superscript_table


######################################################################################################################################


# create footnote_table
footnote_table_1 <- tibble(text = c("This is footnote 1 for table 1.", str_c("This is footnote 2 - it is long",
                                                                             "long long long long long long long long long long long long long long",
                                                                             "long long long long long long long long long long long long long long",
                                                                             "long long long long long long long long long long long long long long",
                                                                             "long long long long long long long long long long long long long long")))
footnote_table_2 <- tibble(text = c("This is footnote 1 for table 2.", str_c("This is footnote 2 - it is reallllllllllllllll",
                                                                             "llllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllllly long")))
footnote_table <- list(footnote_table_1, footnote_table_2)
footnote_table 


######################################################################################################################################


# create style_table
style_table_1 <- tibble(row_type = c("header", "body", "body", "footnote"), rows_from = c(1, 1, 4, 1), rows_to = c(1, 3, 7, 1), 
                        cols_from = c(1, 1, 1, 1), cols_to = c(1, 2, ncol(table_1), 1),
                        font_name = c("Arial", "Arial", NA, "Arial"), font_size = c(NA, 15, 7, NA), font_color = c(NA, "#cc0099", NA, NA),
                        border = c(NA, "top_bottom_left_right", NA, NA), border_color = c(NA, "#ff0000", NA, NA),
                        border_style = c(NA, "thick", NA, NA), background_fill = c(NA, "#66ff66", NA, NA),
                        horizontal_alignment = c("right", "left", NA, "right"), vertical_alignment = c(NA, "top", NA, NA),
                        text_decoration = c(NA, "bold", NA, NA), wrap_text = c(FALSE, TRUE, NA, NA),
                        text_rotation = c(45, NA, NA, NA))

style_table_2 <- tibble(row_type = c("header", "body", "body", "footnote"), rows_from = c(1, 1, 4, 1), rows_to = c(1, 3, 7, 1), 
                        cols_from = c(1, 1, 1, 1), cols_to = c(1, 2, ncol(table_1), 1),
                        font_name = c("Arial", "Arial", NA, "Arial"), font_size = c(NA, 15, 7, NA), font_color = c(NA, "#cc0099", NA, NA),
                        border = c(NA, "blah_top_blah_bottom_blah_right", NA, NA), border_color = c(NA, "#ff0000", NA, NA),
                        border_style = c(NA, "thick", NA, NA), background_fill = c(NA, "#66ff66", NA, NA),
                        horizontal_alignment = c("right", "left", NA, "right"), vertical_alignment = c(NA, "top", NA, NA),
                        text_decoration = c(NA, "bold", NA, NA), wrap_text = c(FALSE, TRUE, NA, NA),
                        text_rotation = c(45, NA, NA, NA))

style_table <- list(style_table_1, style_table_2)


######################################################################################################################################
######################################################################################################################################
######################################################################################################################################


# output single table        
workbook <- createWorkbook() %>% prd_format(tables = table_1, text_cols = 9)
openXL(workbook)
saveWorkbook(wb = workbook, file = "example_table_output.xlsx", overwrite = TRUE)

# output multiple tables, increased col_width_padding    
previous_workbook <- loadWorkbook(file = "example_table_output.xlsx")
openXL(previous_workbook)
prd_format(workbook = previous_workbook, tables = list(table_1, table_2), output_sheet_names = "table_1_sheet", 
           text_cols = list(9, 2), custom_col_width = NULL, custom_row_height = NULL, 
           style_table = NULL, footnote_table = NULL, superscript_table = NULL, 
           col_width_padding = 4, min_col_width = 8.43) %>% openXL()

# output multiple tables w/ custom_col_width and custom_row_height, style_table  
createWorkbook() %>% prd_format(tables = list(table_1, table_2), output_sheet_names = c("table_1_sheet", "table_2_sheet"), 
                                style_table = style_table, footnote_table = footnote_table, superscript_table = superscript_table, 
                                text_cols = list(c(9), c(2, 5)), custom_col_width = list(c(8, NA, 10, NA, 10, 10, 10, 10, 12), c(12, 10, 10, 10, 10, 10, 10, 10, 8)), 
                                custom_row_height = list(c(NA, rep(15, times = nrow(table_1))), c(50, rep(15, times = nrow(table_1)), 15, 25)), 
                                col_width_padding = 1, min_col_width = 8.43) %>% openXL()

# with summary_rows/cols, with footnote table as a single tibble, not a list
createWorkbook() %>% prd_format(tables = list(table_1, table_2), output_sheet_names = c("table_1_sheet", "table_2_sheet"), 
                                summary_rows = list(10, NULL), summary_cols = list(NULL, 9), 
                                style_table = NULL, footnote_table = footnote_table[[1]], superscript_table = superscript_table,
                                text_cols = list(9, 2), custom_col_width = list(c(8, 10, 10, 10, 10, 10, 10, 10, 12), c(12, 10, 10, 10, 10, 10, 10, 10, 8)), 
                                custom_row_height = list(c(40, rep(15, times = nrow(table_1))), c(50, rep(15, times = nrow(table_1)), 15, 25)), 
                                col_width_padding = 1, min_col_width = 8.43) %>% openXL()