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
# https://stackoverflow.com/questions/40234742/superscript-from-r-to-excel-table

# create add_superscript_to_cell function
add_superscript_to_cell <- function(workbook, sheet, row, col, text, superscript_text, superscript_position) {
        
        # create placeholder_text
        placeholder_text <- 'This is placeholder text that should not appear anywhere in your document.'
        
        # add placeholder_text to workbook in specified cell that will contain superscript
        writeData(wb = workbook, sheet = sheet, x = placeholder_text, startRow = row, startCol = col)
        
        # find the workbook$sharedstring that you want to update
        shared_string_to_update <- workbook$sharedStrings %>% enframe() %>% unnest(value) %>%
                filter(str_detect(string = value, pattern = placeholder_text)) %>% pull(name)
        
        # get pre_text before superscript
        # note that blanks must be replaced with a space, because blanks will cause error when compiling into xml
        pre_text <- str_sub(string = text, start = 1, end = superscript_position - 1) 
        pre_text <- ifelse(pre_text == "", " ", pre_text)
        
        # get post_text after superscript
        # note that blanks must be replaced with a space, because blanks will cause error when compiling into xml
        post_text <- str_sub(string = text, start = superscript_position, end = nchar(text)) 
        post_text <- ifelse(post_text == "", " ", post_text)
        
        
        #############
        
        
        # create get_style_tbl function
        get_style_tbl <- function(current_style_object, current_style_index) {
                
                # get style info
                font_color <- ifelse(is.null(current_style_object$style$fontColour$rgb), NA, current_style_object$style$fontColour$rgb)
                font_name <- ifelse(is.null(current_style_object$style$fontName$val), NA, current_style_object$style$fontName$val)
                font_size <- ifelse(is.null(current_style_object$style$fontSize$val), NA, current_style_object$style$fontSize$val)
                text_decoration <- ifelse(is.null(current_style_object$style$fontDecoration), NA, current_style_object$style$fontDecoration)
                text_decoration <- ifelse(length(text_decoration) > 1, 
                                          str_c(text_decoration, collapse = ", "), text_decoration)
                sheet <- current_style_object$sheet
                rows <- current_style_object$rows
                cols <- current_style_object$cols
                
                # get and return current_style_object_tbl
                current_style_object_tbl <- tibble(style_index = rep(current_style_index, times = length(rows)), 
                                                style_sheet = rep(sheet, times = length(rows)), style_row = rows, style_col = cols, 
                                                   font_color = rep(font_color, times = length(rows)), 
                                                   font_name = rep(font_name, times = length(rows)), 
                                                   font_size = rep(font_size, times = length(rows)),
                                                text_decoration = rep(text_decoration, times = length(rows))) %>%
                        mutate(text_decoration = case_when(text_decoration == "BOLD" ~ "<b/>",
                                                           text_decoration == "ITALIC" ~ "<i/>",
                                                           str_detect(string = text_decoration, pattern = "BOLD") & 
                                                                   str_detect(string = text_decoration, pattern = "ITALIC") ~ "<b/><i/>",
                                                           TRUE ~ text_decoration))
                return(current_style_object_tbl)                                   
        }
        
        # get style_tbl
        style_tbl <- map2(.x = workbook$styleObjects, .y = 1:length(workbook$styleObjects), 
             .f = ~ get_style_tbl(current_style_object = .x, current_style_index = .y)) %>%
                bind_rows()
        
        # get target_cell_style_tbl, filtering down style_tbl to current row/col targeted for superscript, and creating xml property strings
        target_cell_style_tbl <- style_tbl %>% filter(style_sheet == sheet, style_row == row, style_col == col) %>%
                arrange(desc(style_index)) %>% slice(1) %>% 
                mutate(font_color = case_when(is.na(font_color) ~ "", !is.na(font_color) ~ str_c('<color rgb =\"', font_color, '\"/>')),
                       font_name = case_when(is.na(font_name) ~ "", !is.na(font_name) ~ str_c('<rFont val =\"', font_name, '\"/>')),
                       font_size = case_when(is.na(font_size) ~ "", !is.na(font_size) ~ str_c('<sz val=\"', font_size, '\"/>')),
                       text_decoration = case_when(is.na(text_decoration) ~ "", TRUE ~ text_decoration))
  
        
        ##############

        
        new_shared_string <- str_c('<si>',
                                   
                                   # handle pre-text
                                   '<r>',
                                   '<rPr>',
                                   target_cell_style_tbl %>% pull(text_decoration),
                                   target_cell_style_tbl %>% pull(font_size),
                                   target_cell_style_tbl %>% pull(font_color),
                                   target_cell_style_tbl %>% pull(font_name),
                                   '</rPr>',
                                   '<t xml:space="preserve">',
                                   pre_text,
                                   '</t>',
                                   '</r>',
                                   
                                   # handle superscript_text
                                   '<r>',
                                   '<rPr>',
                                   '<vertAlign val=\"superscript\"/>',
                                   target_cell_style_tbl %>% pull(text_decoration),
                                   target_cell_style_tbl %>% pull(font_size),
                                   target_cell_style_tbl %>% pull(font_color),
                                   target_cell_style_tbl %>% pull(font_name),
                                   '</rPr>',
                                   '<t xml:space="preserve">',
                                   superscript_text,
                                   '</t>',
                                   '</r>',
                                   
                                   # handle post-text
                                   '<r>',
                                   '<rPr>',
                                   target_cell_style_tbl %>% pull(text_decoration),
                                   target_cell_style_tbl %>% pull(font_size),
                                   target_cell_style_tbl %>% pull(font_color),
                                   target_cell_style_tbl %>% pull(font_name),
                                   '</rPr>',
                                   '<t xml:space="preserve">',
                                   post_text,
                                   '</t>',
                                   '</r>',
                                   '</si>',
                                   sep = "")
        
        # update sharedStrings with new_shared_string
        workbook$sharedStrings[shared_string_to_update] <- new_shared_string
}


##################################################################################


# # test version with pre-existing formatting
# workbook <- createWorkbook()
# addWorksheet(wb = workbook, sheet = "superscript_test")
# writeData(wb = workbook, sheet = "superscript_test", x = tibble(var1 = c(1, 2, 3), var2 = c("red", "blue", "green")))
# test_style_1 <- createStyle(fontName = "Source Sans Pro", fontSize = 14, fontColour = "#ffffff", fgFill = "#1F497D", numFmt = "GENERAL",
#                           border = "top_bottom_left_right", borderColour = "#ffffff", borderStyle = "thick",
#                           halign = "center", valign = "center", textDecoration = "bold", wrapText = TRUE)
# addStyle(wb = workbook, sheet = "superscript_test", style = test_style_1, 
#          rows = 1:3, cols = 1:2, gridExpand = TRUE, stack = TRUE)
# test_style_2 <- createStyle(fontName = "Arial", fontColour = "#00cc00", halign = "left", textDecoration = "italic",
#                             border = "right", borderColour = "#ff0000", borderStyle = "dashed")
# addStyle(wb = workbook, sheet = "superscript_test", style = test_style_2, rows = 2, cols = 2, gridExpand = TRUE, stack = TRUE)
# 
# # inspect
# openXL(workbook)
# workbook
# workbook$sharedStrings
# workbook$styleObjects
# 
# # set args
# row <- 2
# col <- 2
# sheet <- "superscript_test"
# text <- "green_w_superscript"
# superscript_position <- 5
# superscript_text <- "a"
# 
# 
# #########################
# 
# 
# # call add_superscript_to_cell
# add_superscript_to_cell(workbook = workbook, sheet = "superscript_test", row = row, col = col, text = text,
#                         superscript_text = superscript_text, superscript_position = position)
# 
# # inspect
# workbook
# workbook$sharedStrings
# workbook$styleObjects
# openXL(workbook)

      
      

      
      
      
      
      
      
