# # load add_super_script_to_cell function
# current_wd <- getwd()
# setwd("H:/R/openxlsx")
# source("add_super_script_to_cell.R")
# setwd(current_wd)


library(tidyverse)
library(openxlsx)
library(stringr)

# https://github.com/awalker89/openxlsx/issues/407
# https://cran.r-project.org/web/packages/openxlsx/vignettes/Introduction.pdf

# create add_superscript_to_cell function
add_superscript_to_cell <- function(workbook, sheet, row, col, text, superscript_text, position = nchar(text), is_superscript = TRUE,
                                     size = 11, color = "#000000", font = "Calibri", font_family = 1, bold = FALSE, italic = FALSE, underlined = FALSE) {
        
        # create placeholder_text
        placeholder_text <- 'This is placeholder text that should not appear anywhere in your document.'
        
        # add placeholder_text to workbook in specified cell that will contain superscript
        writeData(wb = workbook, sheet = sheet, x = placeholder_text, startRow = row, startCol = col)
        
        # find the workbook$sharedstring that you want to update
        shared_string_to_update <- workbook$sharedStrings %>% enframe() %>% unnest() %>%
                filter(str_detect(string = value, pattern = placeholder_text)) %>% pull(name)
        
        # get pre_text before superscript
        pre_text <- str_sub(string = text, start = 1, end = position) 
        
        # get post_text after superscript
        post_text <- str_sub(string = text, start = position + 1, end = nchar(text)) 
        
        
        #############
        
        
        # handle properties formatting; note "rPR" is a formatting tag that excel reads
        prop_size <- paste('<sz val =\"',size,'\"/>',
                       sep = '')
        
        prop_color <- paste('<color rgb =\"',color,'\"/>',
                       sep = '')
        
        prop_font <- paste('<rFont val =\"',font,'\"/>',
                       sep = '')
        
        prop_font_family <- paste('<family val =\"', font_family,'\"/>',
                       sep = '')
        
        if(bold){
                prop_bold <- '<b/>'
        } else {prop_bold <- ''}
        
        if(italic){
                prop_italic <- '<i/>'
        } else {prop_italic <- ''}
        
        if(underlined){
                prop_underlined <- '<u/>'
        } else {prop_underlined <- ''}
        
        if(is_superscript){
                prop_vertical_align <- '<vertAlign val=\"superscript\"/>'
        } else {prop_vertical_align <- '<vertAlign val=\"subscript\"/>'}

        # get prop_text and prop_superscript_text
        prop_text <- str_c(prop_size, prop_color, prop_font, prop_font_family, prop_bold, prop_italic, prop_underlined, sep = "")
        prop_superscript_text <- str_c(prop_vertical_align, prop_size, prop_color, prop_font, 
                                       prop_font_family, prop_bold, prop_italic, prop_underlined, sep = "")
        
        
        ##############
        
        
        # create new_shared_string
        new_shared_string <- str_c('<si><r><rPr>',
                                   prop_text,
                                   '</rPr><t xml:space="preserve">',
                                   pre_text,
                                   '</t></r><r><rPr>',
                                   prop_superscript_text,
                                   '</rPr><t xml:space="preserve">',
                                   superscript_text,
                                   '</t></r><r><rPr>',
                                   prop_text,
                                   '</rPr><t xml:space="preserve">',
                                   post_text,
                                   '</t></r></si>',
                                   sep = '')
        
        # update sharedStrings with new_shared_string
        workbook$sharedStrings[shared_string_to_update] <- new_shared_string
}


##################################################################################


# test
# workbook <- createWorkbook()
# addWorksheet(wb = workbook, sheet = "superscript_test")
# writeData(wb = workbook, sheet = "superscript_test", x = tibble(var1 = c(1, 2, 3), var2 = c("red", "blue", "green")))
# row <- 3
# col <- 2
# sheet <- "superscript_test"
# position <- 3
# text <- "green_w_superscript"
# is_superscript <- TRUE
# superscript_text <- "a"
# size = 11
# color = "#000000"
# font = "Calibri"
# font_family = 1
# bold = FALSE
# italic = FALSE
# underlined = FALSE
# 
# openXL(workbook)
# workbook
# workbook$sharedStrings
# 
# # note position = 1 actually places superscript text starting at position 2 of text for some reason
# # so the best i can do is text = " text here" with a leading space, and then use position = 1
# add_superscript_to_cell(workbook = workbook, sheet = "superscript_test", row = row, col = col, text = text,
#                          superscript_text = superscript_text, position = position, is_superscript = is_superscript,
#                          size = size, color = color, font = font, font_family = font_family, bold = bold,
#                          italic = italic, underlined = underlined)
# 
# workbook
# workbook$sharedStrings
# openXL(workbook)
















