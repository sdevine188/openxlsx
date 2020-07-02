library(tidyverse)
library(scales)
library(openxlsx)
library(readxl)


######################################################################################################################################
######################################################################################################################################
######################################################################################################################################



# create prd_format function
prd_format <- function(workbook, tables, output_sheet_names = NULL, text_cols = NULL, summary_rows = NULL, summary_cols = NULL,
                       style_table = NULL, footnote_table = NULL, superscript_table = NULL,
                       custom_col_width = NULL, custom_row_height = NULL,
                       col_width_padding = 1, min_col_width = 8.43) {
        
        # create is_even function
        is_even <- function(x) {
                ifelse(x %% 2 == 0, TRUE, FALSE)
        }
        
        
        # create format_current_table function
        format_current_table <- function(workbook, current_table, current_output_sheet_name, 
                                         current_text_cols, current_summary_rows, current_summary_cols,
                                         current_style_table, current_footnote_table, current_superscript_table,
                                         current_custom_col_width, current_custom_row_height, 
                                         current_tailor_header_width, current_tailor_footnote_height, 
                                         current_col_width_padding, current_min_col_width) {
                
                
                ##############################
                
                
                # create styles
                # note that creating several different styles probably wasn't the best method, since i could have used the stack = TRUE argument
                # but at first it seemed like an easy enough option...
                
                # create header_style
                header_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#ffffff", fgFill = "#1F497D", numFmt = "GENERAL",
                                            border = NULL, borderColour = NULL, borderStyle = NULL,
                                            halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                ###################
                
                
                # create body_string_odd_row_style
                body_string_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                         border = NULL, borderColour = NULL, borderStyle = NULL,
                                                         halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_string_even_row_style
                body_string_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                          border = NULL, borderColour = NULL, borderStyle = NULL,
                                                          halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                ######################
                
                
                # create body_numeric_odd_row_style
                body_numeric_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                          border = NULL, borderColour = NULL, borderStyle = NULL,
                                                          halign = "right", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_numeric_even_row_style
                body_numeric_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                           border = NULL, borderColour = NULL, borderStyle = NULL,
                                                           halign = "right", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                ######################
                
                
                
                # create body_first_col_odd_row_style
                body_first_col_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                            border = "right", borderColour = "#B2B2B2", borderStyle = "thick",
                                                            halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_first_col_even_row_style
                body_first_col_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                             border = "right", borderColour = "#B2B2B2", borderStyle = "thick",
                                                             halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                ######################
                
                
                # create body_last_row_first_col_odd_row_style
                body_last_row_first_col_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                                     border = c("right", "bottom"), borderColour = "#B2B2B2", borderStyle = "thick",
                                                                     halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_last_row_first_col_even_row_style
                body_last_row_first_col_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                                      border = c("right", "bottom"), borderColour = "#B2B2B2", borderStyle = "thick",
                                                                      halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_last_row_numeric_odd_row_style
                body_last_row_numeric_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                                   border = "bottom", borderColour = "#B2B2B2", borderStyle = "thick",
                                                                   halign = "right", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_last_row_numeric_even_row_style
                body_last_row_numeric_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                                    border = "bottom", borderColour = "#B2B2B2", borderStyle = "thick",
                                                                    halign = "right", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_last_row_string_odd_row_style
                body_last_row_string_odd_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                                                  border = "bottom", borderColour = "#B2B2B2", borderStyle = "thick",
                                                                  halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                # create body_last_row_string_even_row_style
                body_last_row_string_even_row_style <- createStyle(fontName = "Source Sans Pro", fontSize = 11, fontColour = "#000000", fgFill = "#D6E4F5", numFmt = "GENERAL",
                                                                   border = "bottom", borderColour = "#B2B2B2", borderStyle = "thick",
                                                                   halign = "center", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                ########################
                
                
                # create footnote_style
                footnote_style <- createStyle(fontName = "Source Sans Pro", fontSize = 10, fontColour = "#000000", fgFill = "#ffffff", numFmt = "GENERAL",
                                              border = NULL, borderColour = NULL, borderStyle = NULL,
                                              halign = "left", valign = "center", textDecoration = NULL, wrapText = TRUE)
                
                
                #######################
                
                
                # create summary_rows_style
                summary_rows_style <- createStyle(fgFill = "#DDDDDD", border = c("top", "bottom"), borderColour = "#B2B2B2", borderStyle = "thick")
                
                # create summary_cols_style
                summary_cols_style <- createStyle(fgFill = "#DDDDDD", border = c("left", "right"), borderColour = "#B2B2B2", borderStyle = "thick")
                
                # create summary_cols_last_row_style
                summary_cols_last_row_style <- createStyle(fgFill = "#DDDDDD", border = c("bottom", "left", "right"), borderColour = "#B2B2B2", borderStyle = "thick")
                
                
                ###############################################################################################################################
                
                
                # add sheet
                sheet <- current_output_sheet_name
                addWorksheet(wb = workbook, sheet = sheet)
                
                # add table
                writeData(wb = workbook, sheet = sheet, x = current_table, 
                          borders = "all", borderStyle = "thin", startRow = 2, startCol = 2)
                
                
                ##########################################################################################################
                
                
                # tailor_col_width 
                
                # create tailor_col_widths function
                tailor_col_widths <- function(table, col_width_padding, min_col_width) {
                        
                        # get header_col_width_tbl 
                        header_col_width_tbl <- str_replace_all(string = names(table), pattern = "\n", replacement = " ") %>% 
                                str_split(string = ., pattern = " ") %>% 
                                map(.x = ., .f = ~ nchar(.x)) %>%
                                map(.x = ., .f = ~ max(.x)) %>%
                                tibble(header_max_word_length = .) %>% unnest(header_max_word_length) %>%
                                bind_cols(table %>% names() %>% tibble(var_name = .), .) %>%
                                mutate(header_best_col_width = case_when(header_max_word_length > (8 - (2 * col_width_padding)) ~ header_max_word_length + (2 * col_width_padding), 
                                                                         TRUE ~ min_col_width))
                        
                        # get body_col_width_tbl
                        body_col_width_tbl <- table %>% map(.x = ., .f = ~ max(nchar(.x))) %>% as_tibble() %>%
                                pivot_longer(cols = everything(), names_to = "var_name", values_to = "body_max_nchar") %>%
                                mutate(body_best_col_width = case_when(body_max_nchar > (8 - (2 * col_width_padding)) ~ body_max_nchar + (2 * col_width_padding), 
                                                                       TRUE ~ min_col_width))
                        
                        # get overall_col_width_tbl
                        overall_col_width_tbl <- header_col_width_tbl %>% left_join(., body_col_width_tbl, by = "var_name") %>%
                                mutate(overall_best_col_width = case_when(body_best_col_width > header_best_col_width ~ body_best_col_width, 
                                                                          TRUE ~ header_best_col_width))
                        
                        # set col widths to best_col_width
                        walk2(.x = overall_col_width_tbl %>% pull(overall_best_col_width), .y = seq(from = 2, to = 1 + ncol(current_table), by = 1),
                              .f = ~ setColWidths(wb = workbook, sheet = sheet, cols = .y, widths = .x))
                }
                
                
                ######################
                
                
                # call tailor_col_widths 
                current_table %>% tailor_col_widths(col_width_padding = current_col_width_padding, min_col_width = current_min_col_width)
                
                
                ######################
                
                
                # handle custom_col_width
                # note that custom_row_height is handled below after footnotes are added, and tailor_footnote_height is applied
                if(!is.null(current_custom_col_width)) {
                        
                        walk2(.x = 1:length(current_custom_col_width), 
                              .y = workbook$colWidths[[1]] %>% enframe() %>% rename(current_col_width = value) %>% pull(current_col_width),
                              .f = ~ setColWidths(wb = workbook, sheet = sheet, 
                                                  cols = 1 + .x, widths = ifelse(is.na(current_custom_col_width[.x]), .y, current_custom_col_width[.x])))
                }
                
                
                ################################################################################################################
                
                
                # create tailor_header_height function
                tailor_header_height <- function(table, col_width_padding, min_col_width) {
                        
                        # get header_col_width_tbl 
                        header_col_width_tbl <- str_replace_all(string = names(table), pattern = "\n", replacement = " ") %>%
                                str_split(string = ., pattern = " ") %>%
                                map(.x = ., .f = ~ nchar(.x)) %>%
                                map(.x = ., .f = ~ max(.x)) %>%
                                tibble(header_max_word_length = .) %>% unnest(header_max_word_length) %>%
                                bind_cols(table %>% names() %>% tibble(var_name = .), .) %>%
                                mutate(header_best_col_width = case_when(header_max_word_length > (8 - (2 * col_width_padding)) ~ header_max_word_length + (2 * col_width_padding), 
                                                                         TRUE ~ min_col_width))
                        
                        # get body_col_width_tbl
                        body_col_width_tbl <- table %>% map(.x = ., .f = ~ max(nchar(.x))) %>% as_tibble() %>%
                                pivot_longer(cols = everything(), names_to = "var_name", values_to = "body_max_nchar") %>%
                                mutate(body_best_col_width = case_when(body_max_nchar > (8 - (2 * col_width_padding)) ~ body_max_nchar + (2 * col_width_padding), TRUE ~ min_col_width))
                        
                        # get best_header_col_width
                        best_header_col_width <- header_col_width_tbl %>% left_join(., body_col_width_tbl, by = "var_name") %>%
                                mutate(overall_best_col_width = case_when(body_best_col_width > header_best_col_width ~ body_best_col_width, 
                                                                          TRUE ~ header_best_col_width))
                        
                        # get organic_wrapped_lines_required based on best_header_col_width, excluding newline characters
                        organic_wrapped_lines_required <- best_header_col_width %>% 
                                pull(var_name) %>% str_split(string = ., pattern = "\n") %>%
                                map2_dfr(.x = ., .y = best_header_col_width %>% pull(overall_best_col_width),
                                         .f = ~ tibble(var_name_segments_btw_newlines = .x) %>% 
                                                 mutate(overall_best_col_width = .y,
                                                        overall_best_col_width_minus_padding = overall_best_col_width - (2 * col_width_padding), 
                                                        organic_wrapped_lines_required = (nchar(var_name_segments_btw_newlines) / 
                                                                                                  overall_best_col_width_minus_padding) - .1,
                                                        organic_wrapped_lines_required = case_when(organic_wrapped_lines_required < 1 ~ 0, 
                                                                                                   TRUE ~ ceiling(organic_wrapped_lines_required))) %>%
                                                 summarize(organic_wrapped_lines_required_sum = sum(organic_wrapped_lines_required)))
                        
                        # get best_header_height, which incorporates organic_wrapped_lines_required, as well as newline characters, 
                        # and includes a fixed 5 pt padding on both top and bottom of header row
                        best_header_height <- body_col_width_tbl %>% bind_cols(., organic_wrapped_lines_required) %>% 
                                mutate(newline_character_count = str_count(string = var_name, pattern = "\n"),
                                       total_wrapped_lines_required = newline_character_count + 1 + organic_wrapped_lines_required_sum) %>%
                                filter(total_wrapped_lines_required == max(total_wrapped_lines_required)) %>%
                                mutate(best_header_height = ((total_wrapped_lines_required * 15) + 10)) %>% distinct(best_header_height) %>%
                                pull(best_header_height)
                        
                        # set header row to best_header_height
                        setRowHeights(wb = workbook, sheet = sheet, rows = 2, heights = best_header_height)
                }
                
                
                ######################
                
                
                # call tailor_header_height
                current_table %>% tailor_header_height(col_width_padding = current_col_width_padding, min_col_width = current_min_col_width)
                
                
                ################################################################################################################
                
                
                # for now, manually set current_header_row_count, since support for multi-row headers hasn't been added yet
                current_header_row_count <- 1
                
                
                ################################################################################################################
                
                
                # add styles
                
                # add header_style
                addStyle(wb = workbook, sheet = sheet, style = header_style, 
                         rows = 2, cols = 2:(ncol(current_table) + 1), gridExpand = TRUE)
                
                # add body_string_style
                if(!is.null(current_text_cols)) {
                        
                        # apply body_string_odd_row_style 
                        addStyle(wb = workbook, sheet = sheet, style = body_string_odd_row_style, 
                                 rows = seq(from = 3, to = nrow(current_table) + 2, by = 2), 
                                 cols = as.numeric(current_text_cols) + 1, 
                                 gridExpand = TRUE)
                        
                        # apply body_string_even_row_style 
                        addStyle(wb = workbook, sheet = sheet, style = body_string_even_row_style, 
                                 rows = seq(from = 4, to = nrow(current_table) + 2, by = 2), 
                                 cols = as.numeric(current_text_cols) + 1, 
                                 gridExpand = TRUE)
                }
                
                
                ###################
                
                
                if(!is.null(current_text_cols)) {
                        
                        # add body_numeric_odd_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_numeric_odd_row_style, 
                                 rows = seq(from = 3, to = nrow(current_table) + 2, by = 2), 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-current_text_cols], 
                                 gridExpand = TRUE)
                        
                        # add body_numeric_even_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_numeric_even_row_style, 
                                 rows = seq(from = 4, to = nrow(current_table) + 2, by = 2), 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-current_text_cols], 
                                 gridExpand = TRUE)
                }
                
                if(is.null(current_text_cols)) {
                        
                        # add body_numeric_odd_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_numeric_odd_row_style, 
                                 rows = seq(from = 3, to = nrow(current_table) + 2, by = 2), 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1), 
                                 gridExpand = TRUE)
                        
                        # add body_numeric_even_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_numeric_even_row_style, 
                                 rows = seq(from = 4, to = nrow(current_table) + 2, by = 2), 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1), 
                                 gridExpand = TRUE)
                }
                
                
                ###################
                
                
                # apply body_first_col_odd_row_style 
                addStyle(wb = workbook, sheet = sheet, style = body_first_col_odd_row_style, 
                         rows = seq(from = 3, to = nrow(current_table) + 2, by = 2), 
                         cols = 2, 
                         gridExpand = TRUE)
                
                # apply body_first_col_even_row_style 
                addStyle(wb = workbook, sheet = sheet, style = body_first_col_even_row_style, 
                         rows = seq(from = 4, to = nrow(current_table) + 2, by = 2), 
                         cols = 2, 
                         gridExpand = TRUE)
                
                
                ####################
                
                
                # apply last_row_first_col_style
                if(is_even(current_table %>% nrow())) {
                        
                        # apply body_last_row_first_col_even_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_last_row_first_col_even_row_style, 
                                 rows = nrow(current_table) + 2, 
                                 cols = 2, 
                                 gridExpand = TRUE)
                } else {
                        
                        # apply body_last_row_first_col_odd_row_style
                        addStyle(wb = workbook, sheet = sheet, style = body_last_row_first_col_odd_row_style, 
                                 rows = nrow(current_table) + 2, 
                                 cols = 2, 
                                 gridExpand = TRUE)
                }
                
                
                ####################
                
                
                # apply last_row_string_style
                if(!is.null(current_text_cols)) {
                        
                        if(is_even(current_table %>% nrow())) {
                                
                                # apply body_last_row_string_even_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_string_even_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = current_text_cols + 1, 
                                         gridExpand = TRUE)
                        } else {
                                
                                # apply body_last_row_string_odd_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_string_odd_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = current_text_cols + 1, 
                                         gridExpand = TRUE)
                        } 
                }
                
                
                
                ####################
                
                
                # apply last_row_numeric_style if current_text_cols is not NULL
                if(!is.null(current_text_cols)) {
                        
                        if(is_even(current_table %>% nrow())) {
                                
                                # apply body_last_row_numeric_even_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_numeric_even_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-c(1, current_text_cols)], 
                                         gridExpand = TRUE)
                        } else {
                                
                                # apply body_last_row_numeric_odd_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_numeric_odd_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-c(1, current_text_cols)], 
                                         gridExpand = TRUE)
                        }
                }
                
                # apply last_row_numeric_style if current_text_cols is NULL
                if(is.null(current_text_cols)) {
                        
                        # apply last_row_numeric_style
                        if(is_even(current_table %>% nrow())) {
                                
                                # apply body_last_row_numeric_even_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_numeric_even_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-1], 
                                         gridExpand = TRUE)
                        } else {
                                
                                # apply body_last_row_numeric_odd_row_style
                                addStyle(wb = workbook, sheet = sheet, style = body_last_row_numeric_odd_row_style, 
                                         rows = nrow(current_table) + 2, 
                                         cols = seq(from = 2, to = ncol(current_table) + 1, by = 1)[-1], 
                                         gridExpand = TRUE)
                        }
                }
                
                
                ###################
                
                
                # apply summary_rows_style
                if(!(is.null(current_summary_rows))) {
                        
                        addStyle(wb = workbook, sheet = sheet, style = summary_rows_style, 
                                 rows = current_summary_rows + 2, 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1), 
                                 gridExpand = TRUE, stack = TRUE)
                }
                
                # apply summary_cols_style
                if(!(is.null(current_summary_cols))) {
                        
                        addStyle(wb = workbook, sheet = sheet, style = summary_cols_style, 
                                 rows = seq(from = 3, to = nrow(current_table) + 2, by = 1), 
                                 cols = current_summary_cols + 1, 
                                 gridExpand = TRUE, stack = TRUE)
                }
                
                # apply summary_cols_last_row_style
                if(!(is.null(current_summary_cols))) {
                        
                        addStyle(wb = workbook, sheet = sheet, style = summary_cols_last_row_style, 
                                 rows = nrow(current_table) + 2, 
                                 cols = current_summary_cols + 1, 
                                 gridExpand = TRUE, stack = TRUE)
                }
                
                
                ###########################################################################################################
                
                
                # add footnotes
                if(!is.null(current_footnote_table)) {
                        
                        # add test_table_footnote_1
                        writeData(wb = workbook, sheet = sheet, x = current_footnote_table %>% pull(text), 
                                  startRow = current_table %>% nrow() + 3, startCol = 2)
                        
                        
                        ####################
                        
                        
                        # merge footnote cells
                        walk(.x = 1:nrow(current_footnote_table),
                             .f = ~ mergeCells(wb = workbook, sheet = sheet, cols = 2:(current_table %>% ncol() + 1),
                                               rows = current_table %>% nrow() + 2 + .x))
                        
                        
                        ####################
                        
                        
                        # get total_col_width
                        total_col_width <- workbook$colWidths[[1]] %>% tibble(col_width = .) %>%
                                mutate(col_width = as.numeric(col_width)) %>%
                                summarize(col_width_sum = sum(col_width)) %>% pull(col_width_sum)
                        
                        # get max_characters_in_footnote_row
                        # note that 8.43 is the default column width,
                        # and 10 is the chosen character count for footnote-size-10 source sans pro font that can fit comfortably in 8.43 width columns
                        # note that more characters can fit in merged cell since there's no need for padding
                        max_characters_in_footnote_row <- (total_col_width / 8.43) * 12
                        
                        # update current_footnote_table to get required_rows and row_height
                        # note that 15 is the default row_height
                        current_footnote_table <- current_footnote_table %>% mutate(nchar_text = nchar(text),
                                                                                    required_rows = ceiling(nchar_text / max_characters_in_footnote_row),
                                                                                    row_height = 15 * required_rows)
                        
                        walk(.x = 1:length(current_footnote_table %>% pull(row_height)),
                             .f = ~ setRowHeights(wb = workbook, sheet = sheet, rows = nrow(current_table) + 2 + .x,
                                                  heights = current_footnote_table %>% slice(.x) %>% pull(row_height)))
                        
                        
                        ######################
                        
                        
                        # apply footnote_style
                        addStyle(wb = workbook, sheet = sheet, style = footnote_style, 
                                 rows = (nrow(current_table) + 3):((nrow(current_table) + 2) + nrow(current_footnote_table)), 
                                 cols = seq(from = 2, to = ncol(current_table) + 1, by = 1), 
                                 gridExpand = TRUE)
                        
                }
                
                
                #################################################################################################################
                
                
                # handle current_style_table
                if(!is.null(current_style_table)) {
                        
                        # create apply_current_style_table()
                        apply_current_style_table <- function(data, ...) {
                                
                                eval(parse(text = str_c("addStyle(wb = workbook, sheet = sheet, style = createStyle(",
                                                        data %>% pull(format_element), " = ", data %>% pull(value), "), rows = ",
                                                        data %>% pull(rows), ", cols = ", data %>% pull(cols), ", gridExpand = TRUE, stack = TRUE)")))
                        }
                        
                        
                        ############################
                        
                        
                        # call apply_current_style_table for string values
                        current_style_table %>% 
                                select(row_type, rows_from, rows_to, cols_from, cols_to, font_name, font_color, border, border_color,
                                       border_style, background_fill, horizontal_alignment, vertical_alignment, text_decoration) %>%
                                pivot_longer(cols = -c(row_type, rows_from, rows_to, cols_from, cols_to),
                                             names_to = "format_element", values_to = "value") %>% 
                                filter(!is.na(value)) %>%
                                mutate(format_element = case_when(format_element == "font_name" ~ "fontName",
                                                                  format_element == "font_color" ~ "fontColour",
                                                                  format_element == "border_color" ~ "borderColour",
                                                                  format_element == "border_style" ~ "borderStyle",
                                                                  format_element == "background_fill" ~ "fgFill",
                                                                  format_element == "horizontal_alignment" ~ "halign",
                                                                  format_element == "vertical_alignment" ~ "valign",
                                                                  format_element == "text_decoration" ~ "textDecoration", 
                                                                  TRUE ~ format_element),
                                       value = str_c("'", value, "'"),
                                       rows_from = case_when(row_type == "header" ~ rows_from + 1,
                                                             row_type == "body" ~ rows_from + 1 + current_header_row_count,
                                                             row_type == "footnote" ~ rows_from + 1 + current_header_row_count + nrow(current_table)),
                                       rows_to = case_when(row_type == "header" ~ rows_from + 1,
                                                           row_type == "body" ~ rows_to + 1 + current_header_row_count,
                                                           row_type == "footnote" ~ rows_to + 1 + current_header_row_count + nrow(current_table)),
                                       cols_from = cols_from + 1, cols_to = cols_to + 1,
                                       rows = str_c(rows_from, ":", rows_to), cols = str_c(cols_from, ":", cols_to)) %>%
                                nest(data = everything()) %>%
                                pwalk(.l = ., .f = apply_current_style_table)
                        
                        
                        ########################
                        
                        
                        # call apply_current_style_table for numeric values
                        current_style_table %>% 
                                select(row_type, rows_from, rows_to, cols_from, cols_to, font_size, text_rotation) %>%
                                pivot_longer(cols = -c(row_type, rows_from, rows_to, cols_from, cols_to),
                                             names_to = "format_element", values_to = "value") %>% 
                                filter(!is.na(value)) %>%
                                mutate(format_element = case_when(format_element == "text_rotation" ~ "textRotation",
                                                                  format_element == "font_size" ~ "fontSize"),
                                       rows_from = case_when(row_type == "header" ~ rows_from + 1,
                                                             row_type == "body" ~ rows_from + 1 + current_header_row_count,
                                                             row_type == "footnote" ~ rows_from + 1 + current_header_row_count + nrow(current_table)),
                                       rows_to = case_when(row_type == "header" ~ rows_from + 1,
                                                           row_type == "body" ~ rows_to + 1 + current_header_row_count,
                                                           row_type == "footnote" ~ rows_to + 1 + current_header_row_count + nrow(current_table)),
                                       cols_from = cols_from + 1, cols_to = cols_to + 1,
                                       rows = str_c(rows_from, ":", rows_to), cols = str_c(cols_from, ":", cols_to)) %>%
                                nest(data = everything()) %>%
                                pwalk(.l = ., .f = apply_current_style_table)
                        
                        
                        #######################
                        
                        
                        # call apply_current_style_table for logical values
                        current_style_table %>% 
                                select(row_type, rows_from, rows_to, cols_from, cols_to, wrap_text) %>%
                                pivot_longer(cols = -c(row_type, rows_from, rows_to, cols_from, cols_to),
                                             names_to = "format_element", values_to = "value") %>% 
                                filter(!is.na(value)) %>%
                                mutate(format_element = "wrapText",
                                       rows_from = case_when(row_type == "header" ~ rows_from + 1,
                                                             row_type == "body" ~ rows_from + 1 + current_header_row_count,
                                                             row_type == "footnote" ~ rows_from + 1 + current_header_row_count + nrow(current_table)),
                                       rows_to = case_when(row_type == "header" ~ rows_from + 1,
                                                           row_type == "body" ~ rows_to + 1 + current_header_row_count,
                                                           row_type == "footnote" ~ rows_to + 1 + current_header_row_count + nrow(current_table)),
                                       cols_from = cols_from + 1, cols_to = cols_to + 1,
                                       rows = str_c(rows_from, ":", rows_to), cols = str_c(cols_from, ":", cols_to)) %>%
                                nest(data = everything()) %>%
                                pwalk(.l = ., .f = apply_current_style_table)
                }
                
                
                #########################################################################################################
                
                
                # handle current_superscript_table
                if(!is.null(current_superscript_table)) {
                        
                        # update rows/cols on current_superscript_table to account for table placement offset
                        current_superscript_table <- current_superscript_table %>% 
                                mutate(row = case_when(section == "header" ~ row + 1,
                                                       section == "body" ~ row + 1 + current_header_row_count,
                                                       section == "footnote" ~ row + 1 + current_header_row_count + nrow(current_table)),
                                       col = col + 1)
                        
                        
                        #####################
                        
                        
                        # loop through current_superscript_table calling add_superscript_to_cell()
                        current_superscript_table %>% 
                                pwalk(.l = ., .f = function(section, row, col, superscript_text, superscript_position, prior_text, ...) {
                                        add_superscript_to_cell(workbook = workbook, sheet = sheet, row = row, col = col,
                                                                text = prior_text, superscript_text = superscript_text,
                                                                superscript_position = superscript_position)        
                                })
                }
                
                
                ##########################################################################################################
                
                
                # handle custom_row_height
                if(!is.null(current_custom_row_height)) {
                        
                        walk2(.x = 1:length(current_custom_row_height), 
                              .y = current_custom_row_height,
                              .f = ~ ifelse(is.na(.y), NA, setRowHeights(wb = workbook, sheet = sheet, rows = 1 + .x, heights = current_custom_row_height[.x])))
                }
        }
        
        
        ###################################################################################################################
        
        
        # handle arguments
        
        # if only single table is passed, add it into a list
        if(sum(class(tables) != "list") > 0) {
                tables <- list(tables) 
        } 
        
        # if no output_sheet_names are passed, set defaults
        if(is.null(output_sheet_names)) {
                output_sheet_names <- str_c("table_", seq(from = 1, to = length(tables), by = 1))
        }
        
        # if no text_cols are passed, set defaults
        if(is.null(text_cols)) {
                text_cols <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        # if no summary_rows are passed, set defaults
        if(is.null(summary_rows)) {
                summary_rows <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        # if no summary_cols are passed, set defaults
        if(is.null(summary_cols)) {
                summary_cols <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        # if no footnote_table is passed, set defaults
        if(is.null(footnote_table)) {
                footnote_table <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        # if no style_table is passed, set defaults
        if(is.null(style_table)) {
                style_table <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        # if no superscript_table is passed, set defaults
        if(is.null(superscript_table)) {
                superscript_table <- map(.x = 1:length(tables), .f = ~ NULL)
        }
        
        
        ###################
        
        
        # handle output_sheet_names if multiple tables are passed, but arg is only a single value
        if(length(tables) > 1 & length(output_sheet_names) == 1) {
                output_sheet_names <- map(.x = 1:length(tables), .f = ~ str_c(output_sheet_names, "_", .x))
        }
        
        # handle custom_col_width if multiple tables are passed, but arg is only a single value
        if(length(tables) > 1 & length(custom_col_width) == 1) {
                custom_col_width <- map(.x = 1:length(tables), .f = ~ custom_col_width)
        }
        
        # handle custom_row_height if multiple tables are passed, but arg is only a single value
        if(length(tables) > 1 & length(custom_row_height) == 1) {
                custom_row_height <- map(.x = 1:length(tables), .f = ~ custom_row_height)
        }
        
        # handle col_width_padding if multiple tables are passed, but arg is only a single value
        if(length(tables) > 1 & length(col_width_padding) == 1) {
                col_width_padding <- map(.x = 1:length(tables), .f = ~ col_width_padding)
        }
        
        # handle min_col_width if multiple tables are passed, but arg is only a single value
        if(length(tables) > 1 & length(min_col_width) == 1) {
                min_col_width <- map(.x = 1:length(tables), .f = ~ min_col_width)
        }
        
        # handle footnote_table if it is not passed as a list 
        # (note if not a list, a data.frame/tibble is assumed - currently no support for passing footnote_table as a vector)
        # note the sum(class(footnote_table) == "list") == 0 code is needed to handle when class() is called on tibble and length of output is 3
        if(!is.null(footnote_table) & sum(class(footnote_table) == "list") == 0) {
                footnote_table <- list(footnote_table)
        }
        
        # handle footnote_table if footnote_table is not NULL, but more tables are passed than foonote_tables
        if(length(tables) > 1 & length(footnote_table) > 0 & length(footnote_table) < length(tables)) {
                footnote_table <- c(footnote_table, map(.x = (length(footnote_table) + 1):length(tables), .f = ~ NULL))
        }
        
        # handle style_table if it is not passed as a list 
        # (note if not a list, a data.frame/tibble is assumed - currently no support for passing style_table as a vector)
        # note the sum(class(style_table) == "list") == 0 code is needed to handle when class() is called on tibble and length of output is 3
        if(!is.null(style_table) & sum(class(style_table) == "list") == 0) {
                style_table <- list(style_table)
        }
        
        # handle style_table if style_table is not NULL, but more tables are passed than foonote_tables
        if(length(tables) > 1 & length(style_table) > 0 & length(style_table) < length(tables)) {
                style_table <- c(style_table, map(.x = (length(style_table) + 1):length(tables), .f = ~ NULL))
        }
        
        
        ###################################################################################################################
        
        
        # walk over tables calling format_current_table()
        walk(.x = 1:length(tables), 
             .f = ~ format_current_table(workbook = workbook, 
                                         current_table = tables[[.x]],
                                         current_output_sheet_name = output_sheet_names[[.x]],
                                         current_text_cols = text_cols[[.x]], 
                                         current_summary_rows = summary_rows[[.x]],
                                         current_summary_cols = summary_cols[[.x]],
                                         current_style_table = style_table[[.x]],
                                         current_footnote_table = footnote_table[[.x]],
                                         current_superscript_table = superscript_table[[.x]],
                                         current_custom_col_width = custom_col_width[[.x]], 
                                         current_custom_row_height = custom_row_height[[.x]], 
                                         current_tailor_header_width = tailor_header_width,
                                         current_tailor_footnote_height = tailor_footnote_height,
                                         current_col_width_padding = col_width_padding[[.x]], 
                                         current_min_col_width = min_col_width[[.x]]))
        
        
        ##################################################################################################################
        
        
        # return workbook
        return(workbook)
}


###########################################################################################################################
###########################################################################################################################
###########################################################################################################################


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
