# # load add_super_script_to_cell function
# current_wd <- getwd()
# setwd("H:/R/openxlsx")
# source("add_super_script_to_cell.R")
# setwd(current_wd)


library(tidyverse)
library(openxlsx)
library(stringr)

# https://github.com/awalker89/openxlsx/issues/407

# create add_super_script_to_cell function
add_super_script_to_cell <- function(wb, sheet, row, col, text, superScriptText, position = nchar(text), superOrSub = TRUE, size = '10',
                                 colour = '000000', font = 'Arial', family = '2', bold = FALSE, italic = FALSE, underlined = FALSE) {
        
        placeholderText <- 'This is placeholder text that should not appear anywhere in your document.'
        
        openxlsx::writeData(wb = wb,
                            sheet = sheet,
                            x = placeholderText,
                            startRow = row,
                            startCol = col)
        
        #finds the string that you want to update
        stringToUpdate <- which(sapply(wb$sharedStrings,
                                       function(x){
                                               grep(pattern = placeholderText,
                                                    x)
                                       }
        )
        == 1)
        
        #splits the text into before and after the superscript
        
        preText <- stringr::str_sub(text,
                                    1,
                                    position)
        
        postText <- stringr::str_sub(text,
                                     position + 1,
                                     nchar(text))
        
        #formatting instructions
        
        sz    <- paste('<sz val =\"',size,'\"/>',
                       sep = '')
        col   <- paste('<color rgb =\"',colour,'\"/>',
                       sep = '')
        rFont <- paste('<rFont val =\"',font,'\"/>',
                       sep = '')
        fam   <- paste('<family val =\"',family,'\"/>',
                       sep = '')
        if(superOrSub){
                vert <- '<vertAlign val=\"superscript\"/>'
        } else{vert <- '<vertAlign val=\"subscript\"/>'}
        
        if(bold){
                bld <- '<b/>'
        } else{bld <- ''}
        
        if(italic){
                itl <- '<i/>'
        } else{itl <- ''}
        
        if(underlined){
                uld <- '<u/>'
        } else{uld <- ''}
        
        #run properties
        
        rPrText <- paste(sz,
                         col,
                         rFont,
                         fam,
                         bld,
                         itl,
                         uld,
                         sep = '')
        
        rPrSuperText <- paste(vert,
                              sz,
                              col,
                              rFont,
                              fam,
                              bld,
                              itl,
                              uld,
                              sep = '')
        
        newString <- paste('<si><r><rPr>',
                           rPrText,
                           '</rPr><t xml:space="preserve">',
                           preText,
                           '</t></r><r><rPr>',
                           rPrSuperText,
                           '</rPr><t xml:space="preserve">',
                           superScriptText,
                           '</t></r><r><rPr>',
                           rPrText,
                           '</rPr><t xml:space="preserve">',
                           postText,
                           '</t></r></si>',
                           sep = '')
        
        wb$sharedStrings[stringToUpdate] <- newString
}


##################################################################################


# # test
# workbook <- createWorkbook()
# addWorksheet(wb = workbook, sheet = "superscript_test")
# writeData(wb = workbook, sheet = "superscript_test", x = tibble(var1 = c(1, 2, 3), var2 = c("red", "blue", "green")))
# 
# workbook
# workbook$sharedStrings
# 
# add_super_script_to_cell(wb = workbook, sheet = "superscript_test", row = 3, col = 2, text = "green_w_superscript",
#                                  superScriptText = "a", position = 3, superOrSub = FALSE, size = '10', colour = '000000',
#                                  font = 'Arial', family = '2', bold = FALSE, italic = FALSE, underlined = FALSE)
# 
# workbook
# workbook$sharedStrings
# 
# # inspect workbook
# openXL(workbook)

















