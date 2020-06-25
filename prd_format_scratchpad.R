library(tidyverse)
library(scales)
library(openxlsx)
library(readxl)


# load table_1
setwd("C:/users/sjdevine/Work Folders/Desktop/personal_drive/R/openxlsx")
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
######################################################################################################################################
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


#########################


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


#######################


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
                                map2_dfr(.x = ., .y = best_header_col_width %>% filter(str_detect(string = var_name, pattern = "\n")) %>% pull(overall_best_col_width),
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
                                        total_wrapped_lines_required = newline_character_count + organic_wrapped_lines_required_sum) %>%
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


# inspect workbook
openXL(workbook)


# test
text_cols <- list(c(9), c(2))
summary_rows <- NULL
summary_cols <- NULL
tables <- list(table_1, table_2)
output_sheet_names <- c("table_1", "table_2")
col_width_padding <- .5
min_col_width <- c(8.43, 8.43)
output_path <- "example_table_output.xlsx"
custom_row_height = list(c(40, rep(15, times = nrow(table_1)), 15, 25), c(50, rep(15, times = nrow(table_1)), 15, 25))
custom_col_width = list(c(8, NA, 10, 10, 10, 10, 10, 10, 12))


current_output_sheet_name <- output_sheet_names[[1]]
current_table <- tables[[1]]
current_col_width_padding <- col_width_padding[[1]]        
current_min_col_width <- min_col_width[[1]]
current_text_cols <- text_cols[[1]]
current_summary_rows <- summary_rows[[1]]
current_summary_cols <- summary_cols[[1]]
current_custom_col_width <- custom_col_width[[1]]
current_custom_row_height <- custom_row_height[[1]]
current_style_table <- style_table[[1]]
current_footnote_table <- footnote_table[[1]]
current_superscript_table <- superscript_table[[1]]


########################

        
# output single table        
workbook <- createWorkbook() %>% prd_format(tables = table_1, output_sheet_names = NULL, 
           text_cols = 9, custom_col_width = NULL, custom_row_height = NULL, 
           style_table = NULL, footnote_table = NULL, superscript_table = NULL, 
           col_width_padding = 1, min_col_width = 8.43)
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


