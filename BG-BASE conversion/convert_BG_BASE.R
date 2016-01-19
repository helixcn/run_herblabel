getScriptPath <- function(){
    cmd.args <- commandArgs()
    m <- regexpr("(?<=^--file=).+", cmd.args, perl=TRUE)
    script.dir <- dirname(regmatches(cmd.args, m))
    if(length(script.dir) == 0) stop("can't determine script dir: please call the script with Rscript")
    if(length(script.dir) > 1) stop("can't determine script dir: more than one '--file' argument detected")
    return(script.dir)
}

res <- getScriptPath()
### setwd("O:\\Work team\\Herbarium\\03 herbarium labels\\print herbarium label using herblabel\\CONVERT BG-BASE EXPORTS")
setwd(res)

library(herblabel)
library(openxlsx)
rrr <- bgbase_csv2ht("BG_BASE_EXPORT.csv")
#### setwd("..")
write.xlsx(rrr, "herbarium_specimens_label_data.xlsx")
file.rename(from = "BG_BASE_EXPORT.CSV", to = "CONVERTED_BG_BASE_EXPORT.CSV")
print("Please check and edit the file 'herbarium_specimens_label_data.xlsx' before printing")
