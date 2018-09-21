library(RDCOMClient)

# copy excel data or table as an image
# https://stackoverflow.com/questions/43095645/how-to-export-an-excel-sheet-range-to-a-picture-from-within-r

xlApp <- COMCreate("Excel.Application")
xlWbk <- xlApp$Workbooks()$Open("C:\\Users\\Stephen\\Desktop\\R\\openxlsx\\Date Formatting v2.xlsx")
xlScreen = 1
xlBitmap = 2

xlWbk$Worksheets("cellIs")$Range("A1:B10")$CopyPicture(xlScreen, xlBitmap)

xlApp[['DisplayAlerts']] <- FALSE

oCht <- xlApp[['Charts']]$Add()
oCht$Paste()
oCht$Export("C:\\Users\\Stephen\\Desktop\\R\\openxlsx\\table_image.jpg", "JPG")
oCht$Delete()

# CLOSE WORKBOOK AND APP
xlWbk$Close(FALSE)
xlApp$Quit()

# RELEASE RESOURCES
oCht <- xlWbk <- xlApp <- NULL    
rm(oCht, xlWbk, xlApp)
gc()


