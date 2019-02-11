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
merge_footnote_rows <- function(current_row, workbook, sheet, table, skip_lines = 1, ...) {
        
        # merge cells for each row of footnote
        mergeCells(wb = workbook, sheet = sheet, cols = 1:(table %>% ncol()), 
                   rows = table %>% nrow() + skip_lines + 1 + current_row)
}


#############################


# # test
# table <- tibble(var1 = c(1, 2, 3), var2 = c("red", "blue", "green"))
# table
# 
# footnote_tbl <- tibble(footnote_text = c("1) this is footnote 1", "", "2) this is footnote 2"))
# footnote_tbl
# 
# workbook <- createWorkbook()
# sheet <- "test"
# addWorksheet(wb = workbook, sheet = sheet)
# writeData(wb = workbook, sheet = sheet, x = table, startRow = 1, startCol = 1)
# writeData(wb = workbook, sheet = sheet, x = footnote_tbl, startRow = table %>% nrow() + 3, startCol = 1)
# 
# openXL(workbook)
# 
# # call merge_footnote_rows
# map(.x = seq(from = 1, to = footnote_tbl %>% nrow(), by = 1), 
#     .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = table, skip_lines = 1))
# 
# openXL(workbook)





