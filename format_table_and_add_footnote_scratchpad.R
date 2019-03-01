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
test_table <- starwars %>% select(name, species, homeworld) %>% head()
test_table

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

# create header_style
header_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#ffffff", fgFill = "#003399", numFmt = "GENERAL",
                            border = c("Top", "Bottom", "Left", "Right"),
                            borderColour = "#000000",
                            borderStyle = c("thin", "thin", "thin", "thin"),
                            halign = "center", valign = "center", textDecoration = "Bold",
                            wrapText = TRUE)

# create body_style
body_style <- createStyle(fontName = "Calibri", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                          border = c("Top", "Bottom", "Left", "Right"),
                          borderColour = "#000000",
                          borderStyle = c("thin", "thin", "thin", "thin"),
                          halign = "center", valign = "center", textDecoration = NULL,
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

# create workbook
workbook <- createWorkbook()
sheet <- "test_sheet"
addWorksheet(wb = workbook, sheet = sheet)


##############


# add test_table
writeData(wb = workbook, sheet = sheet, x = test_table, headerStyle = header_style,
          borders = "all", borderStyle = "thin")


################


# add test_table_footnote_1
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_1 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 2, startCol = 1)

# add test_table_footnote_2
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_2 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 2 + (test_table_footnote_1 %>% nrow()) + 1, startCol = 1)

# add test_table_footnote_3
writeData(wb = workbook, sheet = sheet, x = test_table_footnote_3 %>% rename(" " = footnote_text), 
          startRow = test_table %>% nrow() + 2 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1, startCol = 1)

# inspect workbook
openXL(workbook)


###############


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
add_superscript_to_cell(workbook = workbook, sheet = sheet, row = test_table %>% nrow() + 3, 
                         col = 1, text = " test_superscript",
                         superscript_text = "(a) ", position = 1, is_superscript = TRUE,
                         size = 9, color = "#000000", font = "Calibri", font_family = 1, bold = FALSE, 
                         italic = FALSE, underlined = FALSE)


################


# add body_style
addStyle(wb = workbook, sheet = sheet, style = body_style, 
         rows = seq(from = 2, to = test_table %>% nrow() + 1, by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1), gridExpand = TRUE)

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
         rows = seq(from = test_table %>% nrow() + 1 + 1, 
                    to = test_table %>% nrow() + 1 + 1 + (test_table_footnote_1 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)

# add custom footnote_2_style
addStyle(wb = workbook, sheet = sheet, style = footnote_2_style, 
         rows = seq(from = test_table %>% nrow() + 1 + 1 + (test_table_footnote_1 %>% nrow()) + 1, 
                    to = test_table %>% nrow() + 1 + 1 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)

# add custom footnote_3_style
addStyle(wb = workbook, sheet = sheet, style = footnote_3_style, 
         rows = seq(from = test_table %>% nrow() + 1 + 1 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1, 
                    to = test_table %>% nrow() + 1 + 1 + (test_table_footnote_1 %>% nrow()) + 1 + (test_table_footnote_2 %>% nrow()) + 1 +
                            (test_table_footnote_3 %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)


#################


# set row heights for header manually
setRowHeights(wb = workbook, sheet = sheet, rows = 1, heights = 60)

# set row heights for all non-header table rows and for all footnote rows at 15
setRowHeights(wb = workbook, sheet = sheet, rows = seq(from = 2, to = test_table %>% nrow() + 1 + 
                               (list_of_all_footnotes %>% enframe() %>% select(value) %>% unnest() %>% nrow()) + 
                               (1 * length(list_of_all_footnotes))), heights = 15)


# set row heights for individual text_wrapped cells manually
setRowHeights(wb = workbook, sheet = sheet, rows = 11, heights = 30)
setRowHeights(wb = workbook, sheet = sheet, rows = 13, heights = 70)
setRowHeights(wb = workbook, sheet = sheet, rows = 17, heights = 30)
setRowHeights(wb = workbook, sheet = sheet, rows = 20, heights = 30)

# set col width
setColWidths(wb = workbook, sheet = sheet, cols = c(1, 2, 3), widths = c(25, 25, 25))


################


# inspect workbook
openXL(workbook)

# write to file
saveWorkbook(wb = workbook, "test_workbook.xlsx", overwrite = TRUE)

