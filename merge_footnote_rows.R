# # load merge_footnote_rows function
# current_wd <- getwd()
# setwd("H:/R/openxlsx")
# source("merge_footnote_rows.R")
# setwd(current_wd)


library(tidyverse)
library(openxlsx)
library(stringr)

# https://cran.r-project.org/web/packages/openxlsx/vignettes/Introduction.pdf

# create merge_footnote_rows function, that will merge each row for a footnote that is already written to the workbook
merge_footnote_rows <- function(current_row, workbook, sheet, table, list_of_preceding_footnotes, ...) {
        
        # get total_preceding_footnote_rows
        if(!is.null(list_of_preceding_footnotes)) {
                total_preceding_footnote_nrow <- list_of_preceding_footnotes %>% enframe() %>% select(value) %>% unnest() %>% nrow() + 
                        (1 * length(list_of_preceding_footnotes))
        }
        if(is.null(list_of_preceding_footnotes)) {
                total_preceding_footnote_nrow <- 0
        }
        
        # merge cells for each row of footnote
        mergeCells(wb = workbook, sheet = sheet, cols = 1:(table %>% ncol()), 
                   rows = table %>% nrow() + 2 + total_preceding_footnote_nrow + current_row)
}


#############################


# # test
# 
# # create table
# table <- tibble(var1 = c(1, 2, 3), var2 = c("red", "blue", "green"))
# table
# 
# # create title
# title <- "this is the title"
# 
# 
# ###########
# 
# 
# # create footnotes
# footnote_tbl_1 <- tibble(footnote_text = c("1) this is footnote 1", "", "2) this is footnote 2"))
# footnote_tbl_1
# 
# footnote_tbl_2 <- tibble(footnote_text = c("3) this is footnote 3, maybe with different styling so needs to be added separately", "",
#                                            "4) this is footnote 4"))
# footnote_tbl_2
# 
# footnote_tbl_3 <- tibble(footnote_text = c("5) this is footnote 5", "", "6) this is footnote 6"))
# footnote_tbl_3
# 
# # create preceding_footnotes_nested_tbl
# list_of_preceding_footnotes_for_footnote_1 <- NULL
# list_of_preceding_footnotes_for_footnote_2 <- c(footnote_tbl_1)
# # can also pass as a traditional list (if you're a cop)
# # preceding_footnotes <- list(footnote_tbl_1, footnote_tbl_2)
# list_of_preceding_footnotes_for_footnote_3 <- c(footnote_tbl_1, footnote_tbl_2)
# 
# 
# ###############
# 
# 
# # create workbook and write data
# workbook <- createWorkbook()
# sheet <- "test"
# addWorksheet(wb = workbook, sheet = sheet)
# writeData(wb = workbook, sheet = sheet, x = title, startRow = 1, startCol = 1)
# writeData(wb = workbook, sheet = sheet, x = table, startRow = 2, startCol = 1)
# writeData(wb = workbook, sheet = sheet, x = footnote_tbl_1 %>% rename(" " = footnote_text), startRow = table %>% nrow() + 3, startCol = 1)
# writeData(wb = workbook, sheet = sheet, x = footnote_tbl_2 %>% rename(" " = footnote_text),
#           startRow = table %>% nrow() + 3 + (footnote_tbl_1 %>% nrow()) + 1, startCol = 1)
# writeData(wb = workbook, sheet = sheet, x = footnote_tbl_3 %>% rename(" " = footnote_text),
#           startRow = table %>% nrow() + 3 + (footnote_tbl_1 %>% nrow()) + 1 + (footnote_tbl_2 %>% nrow() + 1), startCol = 1)
# 
# # inspect
# openXL(workbook)
# 
# 
# #############
# 
# 
# # for testing
# # list_of_preceding_footnotes <- list_of_preceding_footnotes_for_footnote_3
# # current_row <- 1
# 
# # merge title
# mergeCells(wb = workbook, sheet = sheet, cols = 1:(table %>% ncol()), rows = 1)
# # removeCellMerge(wb = workbook, sheet = sheet, cols = 1:(test_table %>% ncol()), rows = 1)
# 
# # call merge_footnote_rows for footnote_tbl_1
# map(.x = seq(from = 1, to = footnote_tbl_1 %>% nrow() + 1, by = 1),
#     .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = table,
#                                list_of_preceding_footnotes = list_of_preceding_footnotes_for_footnote_1))
# 
# # call merge_footnote_rows for footnote_tbl_2
# map(.x = seq(from = 1, to = footnote_tbl_2 %>% nrow() + 1, by = 1),
#     .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = table,
#                                list_of_preceding_footnotes = list_of_preceding_footnotes_for_footnote_2))
# 
# # call merge_footnote_rows for footnote_tbl_3
# map(.x = seq(from = 1, to = footnote_tbl_3 %>% nrow() + 1, by = 1),
#     .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = table,
#                                list_of_preceding_footnotes = list_of_preceding_footnotes_for_footnote_3))
# 
# # inspect
# openXL(workbook)





