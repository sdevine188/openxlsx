library(tidyverse)
library(openxlsx)
library(stringr)

# https://cran.r-project.org/web/packages/openxlsx/vignettes/Introduction.pdf

# load merge_footnote_rows function
current_wd <- getwd()
setwd("H:/R/openxlsx")
source("merge_footnote_rows.R")
setwd(current_wd)

# load add_superscript_to_cell function
current_wd <- getwd()
setwd("H:/R/openxlsx")
source("add_superscript_to_cell.R")
setwd(current_wd)


setwd("H:/R/openxlsx")


################


# create test_table
test_table <- starwars %>% select(name, species, mass, height) %>% head()
test_table

# create test_table_title
test_table_title <- "This is the title of the starwars table"

# create footnotes
test_table_footnote_1 <- tibble(footnote_text = c("this is the first row of footnote 1", "", "this is the third row of footnote 1 after skipping a line", "",
                                           str_c("this is long part of footnote 1: dddddddddddddddddddddddddddddddddddddddddddddddddddddd",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")))
test_table_footnote_1

test_table_footnote_2 <- tibble(footnote_text = c("this is footnote 2", "", "footnote 2 third row after skipping line"))
test_table_footnote_2

test_table_footnote_3 <- tibble(footnote_text = c("this is footnote 3", "footnote 3 second row, with no skipped line"))
test_table_footnote_3

# get list_of_all_footnotes
list_of_all_footnotes <- c(test_table_footnote_1, test_table_footnote_2, test_table_footnote_3) 
list_of_all_footnotes


################


# create title_style
title_style <- createStyle(fontName = "Calibri", fontSize = 13, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                           border = NULL,
                           borderColour = NULL,
                           borderStyle = NULL,
                           halign = "left", valign = "center", textDecoration = "Bold",
                           wrapText = TRUE)


# create header_style
header_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#ffffff", fgFill = "#003399", numFmt = "GENERAL",
                            border = c("Top", "Bottom", "Left", "Right"),
                            borderColour = "#000000",
                            borderStyle = c("thin", "thin", "thin", "thin"),
                            halign = "center", valign = "center", textDecoration = "Bold",
                            wrapText = TRUE)

# create body_style if you want all body to be styled the same way, but probably better to use body_string_style and body_numeric_style
# body_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
#                           border = c("Top", "Bottom", "Left", "Right"),
#                           borderColour = "#000000",
#                           borderStyle = c("thin", "thin", "thin", "thin"),
#                           halign = "center", valign = "center", textDecoration = NULL,
#                           wrapText = TRUE)

# create body_string_style
body_string_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                          border = c("Top", "Bottom", "Left", "Right"),
                          borderColour = "#000000",
                          borderStyle = c("thin", "thin", "thin", "thin"),
                          halign = "left", valign = "center", textDecoration = NULL,
                          wrapText = TRUE)

# create body_numeric_style
body_numeric_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                 border = c("Top", "Bottom", "Left", "Right"),
                                 borderColour = "#000000",
                                 borderStyle = c("thin", "thin", "thin", "thin"),
                                 halign = "right", valign = "center", textDecoration = NULL,
                                 wrapText = TRUE)

# create footnote_style
footnote_1_style <- createStyle(fontName = "Calibri", fontSize = 9, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                              border = NULL,
                              borderColour = NULL,
                              borderStyle = NULL,
                              halign = "left", valign = "center", textDecoration = NULL,
                              wrapText = TRUE)

footnote_2_style <- createStyle(fontName = "Calibri", fontSize = 9, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                              border = NULL,
                              borderColour = NULL,
                              borderStyle = NULL,
                              halign = "left", valign = "center", textDecoration = "italic",
                              wrapText = TRUE)

footnote_3_style <- createStyle(fontName = "Calibri", fontSize = 9, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                              border = NULL,
                              borderColour = NULL,
                              borderStyle = NULL,
                              halign = "left", valign = "center", textDecoration = "bold",
                              wrapText = TRUE)


###########


# create workbook
workbook <- createWorkbook()
sheet <- "test_sheet"
addWorksheet(wb = workbook, sheet = sheet)

# add test_table_title
writeData(wb = workbook, sheet = sheet, x = test_table_title, 
          borders = "all", borderStyle = "thin", startRow = 1, startCol = 1)

# add test_table
writeData(wb = workbook, sheet = sheet, x = test_table, headerStyle = NULL,
          borders = "all", borderStyle = "thin", startRow = 2, startCol = 1)

# add test_table_footnote_1
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_1 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 3, startCol = 1)

# add test_table_footnote_2
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_2 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1, startCol = 1)

# add test_table_footnote_3
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_3 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1, startCol = 1)

# inspect workbook
# openXL(workbook)


###############


# merge title and footnotes

# merge title
mergeCells(wb = workbook, sheet = sheet, cols = 1:(test_table %>% ncol()), rows = 1)
# removeCellMerge(wb = workbook, sheet = sheet, cols = 1:(test_table %>% ncol()), rows = 1)

# call merge_footnote_rows for test_table_footnote_1
map(.x = seq(from = 1, to = test_table_footnote_1 %>% nrow() + 1, by = 1),
    .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = test_table,
                               list_of_preceding_footnotes = NULL))

# call merge_footnote_rows for test_table_footnote_2
map(.x = seq(from = 1, to = test_table_footnote_2 %>% nrow() + 1, by = 1),
    .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = test_table,
                               list_of_preceding_footnotes = c(test_table_footnote_1)))

# call merge_footnote_rows for test_table_footnote_3
map(.x = seq(from = 1, to = test_table_footnote_3 %>% nrow() + 1, by = 1),
    .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = test_table,
                               list_of_preceding_footnotes = c(test_table_footnote_1, test_table_footnote_2)))


#############


# add superscripts
add_superscript_to_cell(workbook = workbook, sheet = sheet, row = test_table %>% nrow() + 4, 
                         col = 1, text = " test_superscript",
                         superscript_text = "(a) ", position = 1, is_superscript = TRUE,
                         size = 9, color = "#000000", font = "Calibri", font_family = 1, bold = FALSE, 
                         italic = FALSE, underlined = FALSE)


################


# add title_style
addStyle(wb = workbook, sheet = sheet, style = title_style, 
         rows = 1, cols = seq(from = 1, to = test_table %>% ncol(), by = 1), gridExpand = TRUE)

# add header_style
addStyle(wb = workbook, sheet = sheet, style = header_style, 
         rows = 2, 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1), gridExpand = TRUE)

# add body_style 
# addStyle(wb = workbook, sheet = sheet, style = body_style, 
#          rows = seq(from = 3, to = test_table %>% nrow() + 2, by = 1), 
#          cols = seq(from = 1, to = test_table %>% ncol(), by = 1), gridExpand = TRUE)

# add body_string_style
body_string_vars <- test_table %>% select_if(.predicate = is.character) %>% names()
body_string_col_index <- test_table %>% names() %>% as_tibble() %>% mutate(row_number = row_number()) %>% 
        filter(value == body_string_vars) %>% pull(row_number)
addStyle(wb = workbook, sheet = sheet, style = body_string_style, 
         rows = seq(from = 3, to = test_table %>% nrow() + 2, by = 1), 
         cols = body_string_col_index, gridExpand = TRUE)

# add body_numeric_style
body_numeric_vars <- test_table %>% select_if(.predicate = is.numeric) %>% names()
body_numeric_col_index <- test_table %>% names() %>% as_tibble() %>% mutate(row_number = row_number()) %>% 
        filter(value == body_numeric_vars) %>% pull(row_number)
addStyle(wb = workbook, sheet = sheet, style = body_numeric_style, 
         rows = seq(from = 3, to = test_table %>% nrow() + 2, by = 1), 
         cols = body_numeric_col_index, gridExpand = TRUE)

# add footnote_style to all footnotes 
# addStyle(wb = workbook, sheet = sheet, style = footnote_1_style, 
#          rows = seq(from = test_table %>% nrow() + 1 + 1, 
#                     to = test_table %>% nrow() + 1 + 
#                             (list_of_all_footnotes %>% enframe() %>% select(value) %>% unnest() %>% nrow()) + 
#                             (1 * length(list_of_all_footnotes)), by = 1), 
#          cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
#          gridExpand = TRUE)

# add custom footnote_1_style
addStyle(wb = workbook, sheet = sheet, style = footnote_1_style, 
         rows = seq(from = test_table %>% nrow() + 3, 
                    to = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)

# add custom footnote_2_style
addStyle(wb = workbook, sheet = sheet, style = footnote_2_style, 
         rows = seq(from = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1, 
                    to = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)

# add custom footnote_3_style
addStyle(wb = workbook, sheet = sheet, style = footnote_3_style, 
         rows = seq(from = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1, 
                    to = test_table %>% nrow() + 3 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1 +
                            (test_table_footnote_3 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)


#################


# set row heights for title manually
setRowHeights(wb = workbook, sheet = sheet, rows = 1, heights = 20)

# set row heights for header manually
setRowHeights(wb = workbook, sheet = sheet, rows = 2, heights = 60)

# set row heights for all non-header table rows and for all footnote rows at 15
setRowHeights(wb = workbook, sheet = sheet, rows = seq(from = 3, to = test_table %>% nrow() + 1 + 
                               (list_of_all_footnotes %>% enframe() %>% select(value) %>% unnest() %>% nrow()) + 
                               (1 * length(list_of_all_footnotes))), heights = 15)


# set row heights for individual text_wrapped cells manually
setRowHeights(wb = workbook, sheet = sheet, rows = 12, heights = 30)
setRowHeights(wb = workbook, sheet = sheet, rows = 14, heights = 70)
setRowHeights(wb = workbook, sheet = sheet, rows = 18, heights = 30)
setRowHeights(wb = workbook, sheet = sheet, rows = 21, heights = 30)

# set col width
setColWidths(wb = workbook, sheet = sheet, cols = seq(from = 1, to = test_table %>% ncol(), by = 1), 
             widths = rep(25, times = test_table %>% ncol()))


################


# inspect workbook
openXL(workbook)

# write to file
saveWorkbook(wb = workbook, "test_workbook.xlsx", overwrite = TRUE)

