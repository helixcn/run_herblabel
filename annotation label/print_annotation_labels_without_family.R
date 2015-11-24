getScriptPath <- function(){
    cmd.args <- commandArgs()
    m <- regexpr("(?<=^--file=).+", cmd.args, perl=TRUE)
    script.dir <- dirname(regmatches(cmd.args, m))
    if(length(script.dir) == 0) stop("can't determine script dir: please call the script with Rscript")
    if(length(script.dir) > 1) stop("can't determine script dir: more than one '--file' argument detected")
    return(script.dir)
}

### setwd("O:\\Work team\\Herbarium\\03 herbarium labels\\annotation label")

res <- getScriptPath()
setwd(res)

library(openxlsx)
Sys.setlocale(category = "LC_ALL", locale = "Chinese")

if(!file.exists("Annotation_label_data.xlsx")){
   stop("The template file \"Annotation_label_data.xlsx\" can not be found in this directory")   
}
dat_test <- read.xlsx("Annotation_label_data.xlsx")
library(herblabel)
dwc_filled <- fill_sp_dwc(dat_test)
unlink("Annotation_label_data.xlsx")

#### Fill the dataset, edit the herbarium_specimens_label_data file
write.xlsx(x = dwc_filled, file = "Annotation_label_data.xlsx")

dwc_filled2 <- read.xlsx("Annotation_label_data.xlsx")
#### Create the labels for checking or printing
#### 
dwc_filled2$FAMILY <- rep(NA, nrow(dwc_filled))
annotation_label(dat = dwc_filled2, outfile = paste("annotation_labels_to_print.rtf"))

### Save a copy to the history folder
### The time marker
### dat_tag <- gsub(":", "", gsub(" ", "-", paste(Sys.time())))
### 
### if(!dir.exists("xlsx history")){
###     dir.create("xlsx history")
### } 
### 
### file.copy(from = "herbarium_specimens_label_data.xlsx", to = paste("xlsx history/", dat_tag, "_herbarium_specimens_label_data.xlsx", sep = ""))
### 
### 
### if(!dir.exists("RTF history")){
###     dir.create("RTF history")
### } 
### 
### file.copy(from = "herbarium_labels_to_print.rtf", to = paste("RTF history/", dat_tag, "_herbarium_labels_to_print.rtf", sep = ""))

#### Update the data base
if(!dir.exists("Annotation_DB_SAVE")){
    dir.create("Annotation_DB_SAVE")
}

if(!file.exists("Annotation_DB_SAVE/annotation_core_database.xlsx")){
     temppp <- rep(NA, length(colnames(dwc_filled)))
     temppp2 <- t(data.frame(temppp))
     colnames(temppp2) <- colnames(dwc_filled)
     write.xlsx(temppp2, "Annotation_DB_SAVE/annotation_core_database.xlsx")
}

if(file.exists("Annotation_DB_SAVE/annotation_core_database.xlsx")){
    dat_dc_db <- read.xlsx("Annotation_DB_SAVE/annotation_core_database.xlsx")
    if(!all(colnames(dat_dc_db) == colnames(dwc_filled2))){
        stop("Column names of the new data does not match the existing database.")
    } 
    dat_dc_db_char <- paste(dat_dc_db$COLLECTOR, 
          dat_dc_db$COLLECTOR_NUMBER, 
          dat_dc_db$DATE_COLLECTED)
          
    dwc_filled2_char <- paste(dwc_filled2$COLLECTOR, 
          dwc_filled2$COLLECTOR_NUMBER, 
          dwc_filled2$DATE_COLLECTED)
          
    dat_dc_db_GUI <- as.character(dat_dc_db$GLOBAL_UNIQUE_IDENTIFIER) 
    dwc_filled2_GUI <- as.character(dwc_filled2$GLOBAL_UNIQUE_IDENTIFIER)
        
    #### Delete the finded entries, only keep the entries not found in the dwc_filled2 form.
    dat_dc_db_deleted <- dat_dc_db[(!dat_dc_db_GUI %in% dwc_filled2_GUI) | (!dat_dc_db_char %in% dwc_filled2_char), ]
    #### Add all the entries from dwc_filled2
    temp_dat_dc_db <- rbind(dat_dc_db_deleted, dwc_filled2) ## Add the entries not found in the existing database.
}

write.xlsx(temp_dat_dc_db, "Annotation_DB_SAVE/annotation_core_database.xlsx")