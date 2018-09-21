library(tidyverse)
library(openxlsx)
library(viridis)
library(scales)
library(rlang)

# https://cran.r-project.org/web/packages/openxlsx/vignettes/
# https://cran.r-project.org/web/packages/openxlsx/vignettes/Introduction.pdf
# https://cran.r-project.org/web/packages/openxlsx/vignettes/formatting.pdf

# setwd
setwd("C:/Users/Stephen/Desktop/R/openxlsx")

# create basic workbook
wb <- createWorkbook()
addWorksheet(wb, sheetName = "starwars")
addWorksheet(wb, sheetName = "starwars_plot")

# write data 
# note that sheet argument can be the name of sheet, or the number of sheet (eg "starwars" or 1)
writeData(wb, sheet = "starwars", x = starwars)
writeData(wb, sheet = 1, x = starwars)

# insert plot
plot <- starwars %>% ggplot(data = ., aes(x = mass)) + geom_histogram()
plot
insertPlot(wb, sheet = "starwars_plot", xy = c("B", 2)) ## insert plot at cell B2



## set default border Colour and style
wb <- createWorkbook()
options("openxlsx.borderColour" = "#4F80BD")
options("openxlsx.borderStyle" = "thin")
modifyBaseFont(wb, fontSize = 10, fontName = "Arial Narrow")

# add worksheet 
# note that it gets a sheetName, but is still referred to as sheet = 1 in writeDataTable and writeData functions
addWorksheet(wb, sheetName = "starwars", gridLines = FALSE)
addWorksheet(wb, sheetName = "Iris", gridLines = FALSE)

## sheet 1
freezePane(wb, sheet = 1, firstRow = TRUE, firstCol = TRUE) ## freeze first row and column
writeData(wb, sheet = 1, x = starwars,
               colNames = TRUE, rowNames = TRUE)
setColWidths(wb, sheet = 1, cols = "A", widths = 18)


##########


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "Date Formatting.xlsx", overwrite = TRUE)



#############################################################################


# merge cells

## Create a new workbook
wb <- createWorkbook()

## Add a worksheet
addWorksheet(wb, "Sheet 1")
addWorksheet(wb, "Sheet 2")

# add data
writeData(wb, sheet = "Sheet 1", x = iris)

# add a row of data - note that writeData overwrites existing data
new_row <- map_dfc(.x = names(iris), .f = function(.x) {var_name_sym <- sym(.x)
                                                        tibble(!!var_name_sym := .x)})
new_row
writeData(wb, sheet = "Sheet 1", x = new_row, startCol = 1, startRow = 2, colNames = FALSE)

## Merge cells to make second header - note that the upper-left-most cell value is retained for new merged cell
mergeCells(wb, "Sheet 1", cols = 1:2, rows = 1)
mergeCells(wb, "Sheet 1", cols = 3:4, rows = 1)
writeData(wb, "Sheet 1", x = tibble(v1 = "Sepal", v2 = NA, v3 = "Petal"),
          startCol = 1, startRow = 1, colNames = FALSE)

## remove merged cells
removeCellMerge(wb, "Sheet 1", cols = 1:4, rows = 1) # removes any intersecting merges


##########


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "merge_cell_example.xlsx", overwrite = TRUE)


#############################################################################3


# headers and footers

## Create a new workbook
wb <- createWorkbook()

## Add a worksheet
addWorksheet(wb, "Sheet 1")

# add data
writeData(wb, sheet = "Sheet 1", x = head(iris))

# headers and footers
setHeaderFooter(wb, sheet = "Sheet 1",  
                header = c("ODD HEAD LEFT", "ODD HEAD CENTER", "ODD HEAD RIGHT"),
                footer = c("ODD FOOT RIGHT", "ODD FOOT CENTER", "ODD FOOT RIGHT"),
                evenHeader = c("EVEN HEAD LEFT", "EVEN HEAD CENTER", "EVEN HEAD RIGHT"),
                evenFooter = c("EVEN FOOT RIGHT", "EVEN FOOT CENTER", "EVEN FOOT RIGHT"),
                firstHeader = c("TOP", "OF FIRST", "PAGE"),
                firstFooter = c("BOTTOM", "OF FIRST", "PAGE"))


##########


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "headers_and_footers_example.xlsx", overwrite = TRUE)


##############################################################################


# formatting

# create data
df <- data.frame("Date" = Sys.Date()-0:4,
                 "Logical" = c(TRUE, FALSE, TRUE, TRUE, FALSE),
                 "Currency" = paste("$",-2:2),
                 "Accounting" = -2:2,
                 "hLink" = "https://CRAN.R-project.org/",
                 "Percentage" = seq(-1, 1, length.out=5),
                 "TinyNumber" = runif(5) / 1E9, stringsAsFactors = FALSE)
df
glimpse(df)

# set classes of data
class(df$Currency) <- "currency"
class(df$Accounting) <- "accounting"
class(df$hLink) <- "hyperlink"
class(df$Percentage) <- "percentage"
class(df$TinyNumber) <- "scientific"

# optional global settings
options("openxlsx.borderStyle" = "thin")
options("openxlsx.borderColour" = "#4F81BD")

# create a workbook 
wb <- createWorkbook()
# str(wb)
class(wb)

# add a worksheet
addWorksheet(wb, sheetName = "writeData auto-formatting")

# add data
writeData(wb, sheet = 1, x = df, startRow = 2, startCol = 2)
writeData(wb, sheet = 1, x = df, startRow = 9, startCol = 2, borders = "surrounding")
writeData(wb, sheet = 1, x = df, startRow = 16, startCol = 2, borders = "rows")
writeData(wb, sheet = 1, x = df, startRow = 23, startCol = 2, borders ="columns")
writeData(wb, sheet = 1, x = df, startRow = 30, startCol = 2, borders ="all")

# create headerStyle
hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
                   border = "Bottom", fontColour = "white")

# add data with header style
writeData(wb, 1, df, startRow = 16, startCol = 10, headerStyle = hs1,
          borders = "rows", borderStyle = "medium")

## to change the display text for a hyperlink column just write over those cells
writeData(wb, sheet = 1, x = paste("Hyperlink", 1:5), startRow = 17, startCol = 14)

## writing as an Excel Table
addWorksheet(wb, sheetName = "writeDataTable")
writeDataTable(wb, 2, df, startRow = 2, startCol = 2)
writeDataTable(wb, 2, df, startRow = 9, startCol = 2, tableStyle = "TableStyleLight9")
writeDataTable(wb, 2, df, startRow = 16, startCol = 2, tableStyle = "TableStyleLight2")
writeDataTable(wb, 2, df, startRow = 23, startCol = 2, tableStyle = "TableStyleMedium21")


##########


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "Date Formatting.xlsx", overwrite = TRUE)



#############################################################################



# dates

# data.frame of dates
dates <- data.frame("d1" = Sys.Date() - 0:4)
for(i in 1:3) dates <- cbind(dates, dates)
names(dates) <- paste0("d", 1:8)
dates

## Date Formatting
wb <- createWorkbook()
addWorksheet(wb, sheetName = "Date Formatting", gridLines = FALSE)
writeData(wb, sheet = 1, x = dates) ## write without styling

## openxlsx converts columns of class "Date" to Excel dates with the format given by
getOption("openxlsx.dateFormat", "mm/dd/yyyy")

## this can be set via (for example)
options("openxlsx.dateFormat" = "yyyy/mm/dd")
## custom date formats can be made up of any combination of:
## d, dd, ddd, dddd, m, mm, mmm, mmmm, mmmmm, yy, yyyy

# re-write dates using new format
writeData(wb, sheet = 1, x = dates) ## write without styling

# can also create custom styles, as opposed to setting global options
## some custom date format examples
sty <- createStyle(numFmt = "m/d/yy")
addStyle(wb, sheet = 1, style = sty, rows = 2:11, cols = 2, gridExpand = TRUE)

# set column width
setColWidths(wb, sheet = 1, cols = 1:10, widths = 23)


##############


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "Date Formatting v1.xlsx", overwrite = TRUE)


####################################################################################


# conditional formatting

# create workbook
wb <- createWorkbook()

# add worksheets
addWorksheet(wb, sheet = "cellIs")
addWorksheet(wb, sheet = "Moving Row")
addWorksheet(wb, sheet = "Moving Col")
addWorksheet(wb, sheet = "Dependent on 1")
addWorksheet(wb, sheet = "Duplicates")
addWorksheet(wb, sheet = "containsText")
addWorksheet(wb, sheet = "colourScale", zoom = 30)
addWorksheet(wb, sheet = "databar")

# create styles
negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

# rule applies to all each cell in range
writeData(wb, "cellIs", x = -5:5)
writeData(wb, "cellIs", x = LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "cellIs", cols = 1, rows = 1:11, rule = "!=0", style = negStyle)
conditionalFormatting(wb, "cellIs", cols = 1, rows = 1:11, rule = "==0", style = posStyle)

## highlight row dependent on specific column value
writeData(wb, "Moving Row", x = -5:5)
writeData(wb, "Moving Row", x = LETTERS[1:11], startCol = 2)
conditionalFormatting(wb, "Moving Row", cols = 1:2, rows = 1:11, rule = "$A1<0", style = negStyle)
conditionalFormatting(wb, "Moving Row", cols = 1:2, rows = 1:11, rule = "$A1>0", style = posStyle)

## highlight column dependent on specific row value
writeData(wb, "Moving Col", x= -5:5)
writeData(wb, "Moving Col", x = LETTERS[1:11], startCol=2)
conditionalFormatting(wb, "Moving Col", cols = 1:2, rows = 1:11, rule = "A$1<0", style = negStyle)
conditionalFormatting(wb, "Moving Col", cols = 1:2, rows = 1:11, rule = "A$1>0", style = posStyle)

## highlight entire range cols X rows dependent only on cell A1
writeData(wb, "Dependent on 1", -5:5)
writeData(wb, "Dependent on 1", LETTERS[1:11], startCol=2)
conditionalFormatting(wb, "Dependent on 1", cols=1:2, rows=1:11, rule="$A$1<0", style = negStyle)
conditionalFormatting(wb, "Dependent on 1", cols=1:2, rows=1:11, rule="$A$1>0", style = posStyle)

## highlight duplicates using default style
writeData(wb, "Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
conditionalFormatting(wb, "Duplicates", cols = 1, rows = 1:10, type = "duplicates")

## cells containing text
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
writeData(wb, "containsText", sapply(1:10, fn))
conditionalFormatting(wb, "containsText", cols = 1, rows = 1:10, type = "contains", rule = "A")

## colourscale colours cells based on cell value
# having trouble getting colourScale formatting to work
df <- tibble(character_var = c(1, 2, "g", 3), numeric_var1 = c(5, 6, 7, 8), 
             numeric_var2 = c(10, 11, 12, 13), numeric_var3 = c(20, 21, 22, 23))
df

# get color hex values
show_col(viridis_pal()(10))
viridis_pal()(10)

writeData(wb, "colourScale", df) ## write data.frame
## rule is a vector or colours of length 2 or 3 (any hex colour or any of colours())
## If rule is NULL, min and max of cells is used. Rule must be the same length as style or NULL.
conditionalFormatting(wb, "colourScale", cols = 1:(ncol(df) + 1), rows = 1:(nrow(df) + 1),
                      # use viridis min/max 
                      # note i need to remove "FF" on end of viridis hex to 'valid' get 6 digit hex
                      style = c("#440154", "#FDE725"),
                      
                      # style = c(createStyle(bgFill = "#440154"), createStyle(bgFill = "#FDE725")),
                      # rule = c(0, 255),
                      rule = NULL,
                      type = "colourScale")
setColWidths(wb, "colourScale", cols = 1:ncol(df), widths = 10)
setRowHeights(wb, "colourScale", rows = 1:nrow(df), heights = 15)

## Databars
writeData(wb, "databar", -5:5)
conditionalFormatting(wb, "databar", cols = 1, rows = 1:12, type = "databar") ## Default colours


##############


# open a temporary version of workbook in excel
openXL(wb)

# save workbook
saveWorkbook(wb, file = "Date Formatting v2.xlsx", overwrite = TRUE)
