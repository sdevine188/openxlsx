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

# create 
test_table_footnote <- tibble(footnote = c("this is a test footnote", "", "this is the third row after skipping a line", "",
                                           str_c("this is a really long footnote dddddddddddddddddddddddddddddddddddddddddddddddddddddd",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                                                 "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")))
test_table_footnote

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
footnote_style <- createStyle(fontName = "Calibri", fontSize = 9, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                              border = NULL,
                              borderColour = NULL,
                              borderStyle = NULL,
                              halign = "left", valign = "center", textDecoration = NULL,
                              wrapText = TRUE)

# create workbook
workbook <- createWorkbook()
sheet <- "test_sheet"
addWorksheet(wb = workbook, sheet = sheet)
writeData(wb = workbook, sheet = sheet, x = test_table, headerStyle = header_style,
          borders = "all", borderStyle = "thin")
writeData(wb = workbook, sheet = sheet, x = test_table_footnote %>% pull(footnote), 
          startRow = test_table %>% nrow() + 3, startCol = 1)

# add footnote
map(.x = seq(from = 1, to = test_table_footnote %>% nrow(), by = 1),
    .f = ~ merge_footnote_rows(current_row = .x, workbook = workbook, sheet = sheet, table = test_table, skip_lines = 1))

add_super_script_to_cell(workbook = workbook, sheet = sheet, row = test_table %>% nrow() + 3, 
                         col = 1, text = " green_w_superscript",
                         superscript_text = "(a) ", position = 1, is_superscript = TRUE,
                         size = 9, color = "#000000", font = "Calibri", font_family = 1, bold = FALSE, 
                         italic = FALSE, underlined = FALSE)

# add body_style
addStyle(wb = workbook, sheet = sheet, style = body_style, 
         rows = seq(from = 2, to = test_table %>% nrow() + 1, by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1), gridExpand = TRUE)

# add footnote_style
addStyle(wb = workbook, sheet = sheet, style = footnote_style, 
         rows = seq(from = test_table %>% nrow() + 2, 
                    to = test_table_footnote %>% nrow() + 2 + (test_table %>% nrow()), by = 1), 
         cols = seq(from = 1, to = test_table %>% ncol(), by = 1),
         gridExpand = TRUE)

# set row heights for header manually
setRowHeights(wb = workbook, sheet = sheet, rows = 1, heights = 60)

# set row heights for individual text_wrapped cells manually
setRowHeights(wb = workbook, sheet = sheet, rows = 13, heights = 70)

# set col width
setColWidths(wb = workbook, sheet = sheet, cols = c(1, 2, 3), widths = c(25, 25, 25))

# inspect workbook
openXL(workbook)

# write to file
saveWorkbook(wb = workbook, "test_workbook.xlsx", overwrite = TRUE)

