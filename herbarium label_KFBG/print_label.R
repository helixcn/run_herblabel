getScriptPath <- function(){
    cmd.args <- commandArgs()
    m <- regexpr("(?<=^--file=).+", cmd.args, perl=TRUE)
    script.dir <- dirname(regmatches(cmd.args, m))
    if(length(script.dir) == 0) stop("can't determine script dir: please call the script with Rscript")
    if(length(script.dir) > 1) stop("can't determine script dir: more than one '--file' argument detected")
    return(script.dir)
}

replace_space <- function(x){gsub("^[[:space:]]+|[[:space:]]+$", "", x)}

res <- getScriptPath()
setwd(res)

#### setwd("/home/jinlong/Documents/github packages/run_herblabel/herbarium label")
#### setwd("C:/github packages/run_herblabel/herbarium label")
#### setwd("C:\\Jinlong\\packages\\run_herblabel\\herbarium label")
invisible(library(openxlsx))
invisible(Sys.setlocale("LC_TIME", "C"))
library(herblabel)


if(!file.exists("herbarium_specimens_label_data.xlsx")){
   stop("The template file \"herbarium_specimens_label_data.xlsx\" can not be found in this directory")   
}

### Save a copy to the history folder
### The time marker
dat_tag <- gsub(":", "", gsub(" ", "-", paste(Sys.time())))

if(!dir.exists("xlsx history")){
    dir.create("xlsx history")
} 

invisible(file.copy(from = "herbarium_specimens_label_data.xlsx", to = paste("xlsx history/", dat_tag, "_herbarium_specimens_label_data.xlsx", sep = "")))

if(!dir.exists("RTF history")){
    dir.create("RTF history")
} 

invisible(file.copy(from = "herbarium_labels_to_print.rtf", to = paste("RTF history/", dat_tag, "_herbarium_labels_to_print.rtf", sep = "")))

library(openxlsx)
#setwd("/home/jinlong/Documents/github packages/run_herblabel/herbarium label")
dat_test <- read.xlsx("herbarium_specimens_label_data.xlsx")
dat_test <- fill_dwc(dat_test)

###########################################################################

### unlink("herbarium_specimens_label_data.xlsx")

herbdat000 <- dat_test
ddd <- rep("", nrow(dat_test) * ncol(dat_test))
dim(ddd) <- c(nrow(dat_test), ncol(dat_test))
colnames(ddd) <- colnames(dat_test)
ddd <- as.data.frame(ddd)
#####     herbdat000[herbdat000 == ""] <- NA
#####     herbdat000$LAT_FLAG <- toupper(herbdat000$LAT_FLAG)
#####     herbdat000$LON_FLAG <- toupper(herbdat000$LON_FLAG)

wb <- createWorkbook()
addWorksheet(wb, "Sheet 1")

##write data to worksheet 1
writeData(wb, sheet = 1, dat_test, rowNames = FALSE)


## style for body
bodyStyle <- createStyle(fontColour = "#FF0000", fgFill = "#FFFF00", border = "TopBottomLeftRight")


#### Check the dataset
if(any(is.na(herbdat000$HERBARIUM))){
    ddd$HERBARIUM[which(is.na(herbdat000$HERBARIUM))] <- "WARNING:HERBARIUM not provided"
    position <- which(ddd == "WARNING:HERBARIUM not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "HERBARIUM not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat000$COLLECTOR))){
    ddd$COLLECTOR[which(is.na(herbdat000$COLLECTOR))] <- "WARNING:COLLECTOR not provided"
    position <- which(ddd == "WARNING:COLLECTOR not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COLLECTOR not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
    
if(any(is.na(herbdat000$COLLECTOR_NUMBER))){
    ddd$COLLECTOR_NUMBER[which(is.na(herbdat000$COLLECTOR_NUMBER))] <- "WARNING:COLLECTOR_NUMBER not provided"
    position <- which(ddd == "WARNING:COLLECTOR_NUMBER not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COLLECTOR_NUMBER not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat000$DATE_COLLECTED))){
    ddd$DATE_COLLECTED[which(is.na(herbdat000$DATE_COLLECTED))] <- "WARNING:DATE_COLLECTED not provided"
    position <- which(ddd == "WARNING:DATE_COLLECTED not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "DATE_COLLECTED not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat000$FAMILY) )){
    ddd$FAMILY[which(is.na(herbdat000$FAMILY))] <- "WARNING:FAMILY not provided"
    position <- which(ddd == "WARNING:FAMILY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "FAMILY not provided"))
    addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat000$GENUS))){
    ddd$GENUS[which(is.na(herbdat000$GENUS))] <- "WARNING:GENUS not provided"
    position <- which(ddd == "WARNING:GENUS not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "GENUS not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
if(any(is.na(herbdat000$COUNTRY))){
    ddd$COUNTRY[which(is.na(herbdat000$COUNTRY))] <- "WARNING:COUNTRY not provided"
    position <- which(ddd == "WARNING:COUNTRY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COUNTRY not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
if(any(is.na(herbdat000$STATE_PROVINCE))){
    ddd$STATE_PROVINCE[which(is.na(herbdat000$STATE_PROVINCE))] <- "WARNING:STATE_PROVINCE not provided"
    position <- which(ddd == "WARNING:STATE_PROVINCE not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "STATE_PROVINCE not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
if(any(is.na(herbdat000$COUNTY))){
    ddd$COUNTY[which(is.na(herbdat000$COUNTY))] <- "WARNING:COUNTY not provided"
    position <- which(ddd == "WARNING:COUNTY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
       writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COUNTY not provided"))
       addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
if(any(is.na(herbdat000$LOCALITY))){
    ddd$LOCALITY[which(is.na(herbdat000$LOCALITY))] <- "WARNING:LOCALITY not provided"
    position <- which(ddd == "WARNING:LOCALITY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "LOCALITY not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}
if(any(is.na(herbdat000$IDENTIFIED_BY))){
    ddd$IDENTIFIED_BY[which(is.na(herbdat000$IDENTIFIED_BY))] <- "WARNING:IDENTIFIED_BY not provided"
    position <- which(ddd == "WARNING:IDENTIFIED_BY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
       writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "IDENTIFIED_BY not provided"))
       addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}


if(any(is.na(herbdat000$DATE_IDENTIFIED))){
    ddd$DATE_IDENTIFIED[which(is.na(herbdat000$DATE_IDENTIFIED))] <- "WARNING:DATE_IDENTIFIED not provided"
    position <- which(ddd == "WARNING:DATE_IDENTIFIED not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
        writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "DATE_IDENTIFIED not provided"))
        addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

#### if( any(as.Date(herbdat000$DATE_IDENTIFIED) <  as.Date(herbdat000$DATE_COLLECTED))){
####     ddd$DATE_IDENTIFIED[which(as.Date(herbdat000$DATE_IDENTIFIED) < as.Date(herbdat000$DATE_COLLECTED))] <- "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED"
####     position <- which(ddd == "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED" , arr.ind = TRUE )
####     for(i in 1:nrow(position)){
####         writeComment(wb, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED"))
####         addStyle(wb, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
####     }
#### } 

####################### Check the vadility of geographical coordinates #######################
col_no_LAT_DEGREE <- which(colnames(herbdat000) == "LAT_DEGREE")
col_no_LAT_MINUTE <- which(colnames(herbdat000) == "LAT_MINUTE")
col_no_LAT_SECOND <- which(colnames(herbdat000) == "LAT_SECOND")
col_no_LAT_FLAG   <- which(colnames(herbdat000) == "LAT_FLAG")
col_no_LON_DEGREE <- which(colnames(herbdat000) == "LON_DEGREE")
col_no_LON_MINUTE <- which(colnames(herbdat000) == "LON_MINUTE")
col_no_LON_SECOND <- which(colnames(herbdat000) == "LON_SECOND")
col_no_LON_FLAG   <- which(colnames(herbdat000) == "LON_FLAG")

herbdat000$LAT_DEGREE  <- as.numeric(herbdat000$LAT_DEGREE)
herbdat000$LAT_MINUTE  <- as.numeric(herbdat000$LAT_MINUTE)
herbdat000$LAT_SECOND  <- as.numeric(herbdat000$LAT_SECOND)
herbdat000$LON_DEGREE  <- as.numeric(herbdat000$LON_DEGREE)
herbdat000$LON_MINUTE  <- as.numeric(herbdat000$LON_MINUTE)
herbdat000$LON_SECOND  <- as.numeric(herbdat000$LON_SECOND)

##################################################################
for (i in 1:nrow(herbdat000)){
    ##### herbdat000$LAT_DEGREE
    ##### col_no_LAT_DEGREE
    if(!is.na(herbdat000$LAT_DEGREE[i])){
        if(herbdat000$LAT_DEGREE[i] < 0 | herbdat000$LAT_DEGREE[i] > 90){
            writeComment(wb, 1, col = col_no_LAT_DEGREE, row = i + 1, comment = createComment(comment = "Degrees must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LAT_DEGREE, rows = i + 1)
        }
    }
    
    ##### herbdat000$LAT_MINUTE col_no_LAT_MINUTE
    if(!is.na(herbdat000$LAT_MINUTE[i])){
        if(herbdat000$LAT_MINUTE[i] < 0 | herbdat000$LAT_MINUTE[i] > 60){
            writeComment(wb, 1, col = col_no_LAT_MINUTE, row = i + 1, comment = createComment(comment = "Minutes must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LAT_MINUTE, rows = i + 1)
        }
    }
    ##### herbdat000$LAT_SECOND  col_no_LAT_SECOND
    if(!is.na(herbdat000$LAT_SECOND[i])){
        if(herbdat000$LAT_SECOND[i] < 0 | herbdat000$LAT_SECOND[i] > 60){
            writeComment(wb, 1, col = col_no_LAT_SECOND, row = i + 1, comment = createComment(comment = "Seconds must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LAT_SECOND, rows = i + 1)
        }
    }
    ##### herbdat000$LAT_FLAG  col_no_LAT_FLAG 
    if(!is.na(herbdat000$LAT_FLAG[i])){
        if(!herbdat000$LAT_FLAG[i] %in% c("N", "S")){
            writeComment(wb, 1, col = col_no_LAT_FLAG, row = i + 1, comment = createComment(comment = "Latitudinal Flag should be either N or S"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LAT_FLAG, rows = i + 1)
        }
    }
    ##### herbdat000$LON_DEGREE col_no_LON_DEGREE 
    if(!is.na(herbdat000$LON_DEGREE[i])){
        if(herbdat000$LON_DEGREE[i] < 0 | herbdat000$LON_DEGREE[i] > 180){
            writeComment(wb, 1, col = col_no_LON_DEGREE, row = i + 1, comment = createComment(comment = "Degrees must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LON_DEGREE, rows = i + 1)
        }
    }
    ##### herbdat000$LON_MINUTE  col_no_LON_MINUTE
    if(!is.na(herbdat000$LON_MINUTE[i])){
        if(herbdat000$LON_MINUTE[i] < 0 | herbdat000$LON_MINUTE[i] > 60){
            writeComment(wb, 1, col = col_no_LON_MINUTE, row = i + 1, comment = createComment(comment = "Minutes must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LON_MINUTE, rows = i + 1)
        }
    }
    ##### herbdat000$LON_SECOND  col_no_LON_SECOND
    if(!is.na(herbdat000$LON_SECOND[i])){
        if(herbdat000$LON_SECOND[i] < 0 | herbdat000$LON_SECOND[i] > 60){
            writeComment(wb, 1, col = col_no_LON_SECOND, row = i + 1, comment = createComment(comment = "Seconds must be between 0 and 90"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LON_SECOND, rows = i + 1)
        }
    }
    ##### herbdat000$LON_FLAG col_no_LON_FLAG 
    if(!is.na(herbdat000$LON_FLAG[i])){
        if(!herbdat000$LON_FLAG[i] %in% c("E", "W")){
            writeComment(wb, 1, col = col_no_LON_FLAG, row = i + 1, comment = createComment(comment = "Longitudinal Flag must be either E or W"))
            addStyle(wb, 1, bodyStyle, cols = col_no_LON_FLAG, rows = i + 1)
        }
    }
}

##################################################################
saveWorkbook(wb, "herbarium_specimens_label_data.xlsx", overwrite = TRUE)

#######################################################
#### which(!is.na(dat_test_res) , arr.ind = TRUE )

#### Fill the dataset, edit the herbarium_specimens_label_data file
#### write.xlsx(x = dwc_filled, file = "herbarium_specimens_label_data.xlsx")

dwc_filled2 <- read.xlsx("herbarium_specimens_label_data.xlsx")
###### dat = dwc_filled2
#### Create the labels for checking or printing
xxx_filled <- herbarium_label(dat = dwc_filled2, outfile = paste("herbarium_labels_to_print.rtf"))

filled_temp <- createWorkbook()
addWorksheet(filled_temp, "Sheet 1")
writeData(filled_temp, sheet = 1, xxx_filled[[1]], rowNames = FALSE)

#### length(xxx_filled)

####################### check the status of genera and families ###################
herbdat_comments <- data.frame(xxx_filled[[2]])
col_genera <- which(colnames(herbdat_comments) == "GENUS")
row_genera <- which(grepl("Empty Species",                                            herbdat_comments$GENUS )| 
                    grepl("This species can not be found",                            herbdat_comments$GENUS )|
                    grepl("Genus not accepted at The Plant List Website",             herbdat_comments$GENUS )|
                    grepl("could also be under",                                      herbdat_comments$GENUS )|
                    grepl("should be under",                                          herbdat_comments$GENUS ))

if(length(row_genera) > 0){
    for(i in 1:length(row_genera)){
        writeComment(filled_temp, 1, col = col_genera, row = row_genera[i] + 1, comment = createComment(comment = gsub(gsub("[ a-zA-Z]", "", herbdat_comments$GENUS[row_genera][i]), "", herbdat_comments$GENUS[row_genera][i])))
        addStyle(filled_temp, 1, bodyStyle, cols = col_genera, rows = row_genera[i] + 1)
        print(gsub(herbdat_comments$GENUS[row_genera][i], "", herbdat_comments$GENUS[row_genera][i]))
    }
}


col_famliy <- which(colnames(herbdat_comments) == "FAMILY")
row_family <- which(grepl("Empty Family", herbdat_comments$FAMILY)|grepl("Family not accepted at The Plant List Website", herbdat_comments$FAMILY ))

if(length(row_family) > 0 ){
    for(i in 1:length(row_family)){
        writeComment(filled_temp, 1, col = col_famliy, row = row_family[i] + 1, comment = createComment(comment = herbdat_comments$FAMILY[row_family][i]))
        addStyle(filled_temp, 1, bodyStyle, cols = col_famliy, rows = row_family[i] + 1)
    }
}

if(any(is.na(herbdat_comments$HERBARIUM))){
    herbdat_comments$HERBARIUM[which(is.na(herbdat_comments$HERBARIUM))] <- "WARNING:HERBARIUM not provided"
    position <- which(herbdat_comments == "WARNING:HERBARIUM not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "HERBARIUM not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$COLLECTOR))){
    herbdat_comments$COLLECTOR[which(is.na(herbdat_comments$COLLECTOR))] <- "WARNING:COLLECTOR not provided"
    position <- which(herbdat_comments == "WARNING:COLLECTOR not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COLLECTOR not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$COLLECTOR_NUMBER))){
    herbdat_comments$COLLECTOR_NUMBER[which(is.na(herbdat_comments$COLLECTOR_NUMBER))] <- "WARNING:COLLECTOR_NUMBER not provided"
    position <- which(herbdat_comments == "WARNING:COLLECTOR_NUMBER not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COLLECTOR_NUMBER not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$DATE_COLLECTED))){
    herbdat_comments$DATE_COLLECTED[which(is.na(herbdat_comments$DATE_COLLECTED))] <- "WARNING:DATE_COLLECTED not provided"
    position <- which(herbdat_comments == "WARNING:DATE_COLLECTED not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "DATE_COLLECTED not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$COUNTRY))){
    herbdat_comments$COUNTRY[which(is.na(herbdat_comments$COUNTRY))] <- "WARNING:COUNTRY not provided"
    position <- which(herbdat_comments == "WARNING:COUNTRY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COUNTRY not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$STATE_PROVINCE))){
    herbdat_comments$STATE_PROVINCE[which(is.na(herbdat_comments$STATE_PROVINCE))] <- "WARNING:STATE_PROVINCE not provided"
    position <- which(herbdat_comments == "WARNING:STATE_PROVINCE not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "STATE_PROVINCE not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$COUNTY))){
    herbdat_comments$COUNTY[which(is.na(herbdat_comments$COUNTY))] <- "WARNING:COUNTY not provided"
    position <- which(herbdat_comments == "WARNING:COUNTY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "COUNTY not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$LOCALITY))){
    herbdat_comments$LOCALITY[which(is.na(herbdat_comments$LOCALITY))] <- "WARNING:LOCALITY not provided"
    position <- which(herbdat_comments == "WARNING:LOCALITY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "LOCALITY not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$IDENTIFIED_BY))){
    herbdat_comments$IDENTIFIED_BY[which(is.na(herbdat_comments$IDENTIFIED_BY))] <- "WARNING:IDENTIFIED_BY not provided"
    position <- which(herbdat_comments == "WARNING:IDENTIFIED_BY not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "IDENTIFIED_BY not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

if(any(is.na(herbdat_comments$DATE_IDENTIFIED))){
    herbdat_comments$DATE_IDENTIFIED[which(is.na(herbdat_comments$DATE_IDENTIFIED))] <- "WARNING:DATE_IDENTIFIED not provided"
    position <- which(herbdat_comments == "WARNING:DATE_IDENTIFIED not provided" , arr.ind = TRUE )
    for(i in 1:nrow(position)){
    writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "DATE_IDENTIFIED not provided"))
    addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
    }
}

#### if( any(as.Date(herbdat_comments$DATE_IDENTIFIED) <  as.Date(herbdat_comments$DATE_COLLECTED))){
####     herbdat_comments$DATE_IDENTIFIED[which(as.Date(herbdat_comments$DATE_IDENTIFIED) < as.Date(herbdat_comments$DATE_COLLECTED))] <- "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED"
####     position <- which(herbdat_comments ==  "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED" , arr.ind = TRUE )
####     for(i in 1:nrow(position)){
####         writeComment(filled_temp, 1, col = position[i,2], row = position[i,1] + 1, comment = createComment(comment = "WARNING:DATE_IDENTIFIED is earlier than DATE_COLLECTED"))
####         addStyle(filled_temp, 1, bodyStyle, cols = position[i,2], rows = position[i,1] + 1)
####     }
#### }

####################### Check the vadility of geographical coordinates #######################
col_no_LAT_DEGREE <- which(colnames(herbdat_comments) == "LAT_DEGREE")
col_no_LAT_MINUTE <- which(colnames(herbdat_comments) == "LAT_MINUTE")
col_no_LAT_SECOND <- which(colnames(herbdat_comments) == "LAT_SECOND")
col_no_LAT_FLAG   <- which(colnames(herbdat_comments) == "LAT_FLAG")
col_no_LON_DEGREE <- which(colnames(herbdat_comments) == "LON_DEGREE")
col_no_LON_MINUTE <- which(colnames(herbdat_comments) == "LON_MINUTE")
col_no_LON_SECOND <- which(colnames(herbdat_comments) == "LON_SECOND")
col_no_LON_FLAG   <- which(colnames(herbdat_comments) == "LON_FLAG")

herbdat_comments$LAT_DEGREE  <- as.numeric(herbdat_comments$LAT_DEGREE)
herbdat_comments$LAT_MINUTE  <- as.numeric(herbdat_comments$LAT_MINUTE)
herbdat_comments$LAT_SECOND  <- as.numeric(herbdat_comments$LAT_SECOND)
herbdat_comments$LON_DEGREE  <- as.numeric(herbdat_comments$LON_DEGREE)
herbdat_comments$LON_MINUTE  <- as.numeric(herbdat_comments$LON_MINUTE)
herbdat_comments$LON_SECOND  <- as.numeric(herbdat_comments$LON_SECOND)

##################################################################
for (i in 1:nrow(herbdat_comments)){
    ##### herbdat_comments$LAT_DEGREE
    ##### col_no_LAT_DEGREE
    if(!is.na(herbdat_comments$LAT_DEGREE[i])){
        if(herbdat_comments$LAT_DEGREE[i] < 0 | herbdat_comments$LAT_DEGREE[i] > 90){
            writeComment(filled_temp, 1, col = col_no_LAT_DEGREE, row = i + 1, comment = createComment(comment = "Degrees must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LAT_DEGREE, rows = i + 1)
        }
    }
    
    ##### herbdat_comments$LAT_MINUTE col_no_LAT_MINUTE
    if(!is.na(herbdat_comments$LAT_MINUTE[i])){
        if(herbdat_comments$LAT_MINUTE[i] < 0 | herbdat_comments$LAT_MINUTE[i] > 60){
            writeComment(filled_temp, 1, col = col_no_LAT_MINUTE, row = i + 1, comment = createComment(comment = "Minutes must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LAT_MINUTE, rows = i + 1)
        }
    }
    ##### herbdat_comments$LAT_SECOND  col_no_LAT_SECOND
    if(!is.na(herbdat_comments$LAT_SECOND[i])){
        if(herbdat_comments$LAT_SECOND[i] < 0 | herbdat_comments$LAT_SECOND[i] > 60){
            writeComment(filled_temp, 1, col = col_no_LAT_SECOND, row = i + 1, comment = createComment(comment = "Seconds must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LAT_SECOND, rows = i + 1)
        }
    }
    ##### herbdat_comments$LAT_FLAG  col_no_LAT_FLAG 
    if(!is.na(herbdat_comments$LAT_FLAG[i])){
        if(!herbdat_comments$LAT_FLAG[i] %in% c("N", "S")){
            writeComment(filled_temp, 1, col = col_no_LAT_FLAG, row = i + 1, comment = createComment(comment = "Latitudinal Flag should be either N or S"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LAT_FLAG, rows = i + 1)
        }
    }
    ##### herbdat_comments$LON_DEGREE col_no_LON_DEGREE 
    if(!is.na(herbdat_comments$LON_DEGREE[i])){
        if(herbdat_comments$LON_DEGREE[i] < 0 | herbdat_comments$LON_DEGREE[i] > 180){
            writeComment(filled_temp, 1, col = col_no_LON_DEGREE, row = i + 1, comment = createComment(comment = "Degrees must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LON_DEGREE, rows = i + 1)
        }
    }
    ##### herbdat_comments$LON_MINUTE  col_no_LON_MINUTE
    if(!is.na(herbdat_comments$LON_MINUTE[i])){
        if(herbdat_comments$LON_MINUTE[i] < 0 | herbdat_comments$LON_MINUTE[i] > 60){
            writeComment(filled_temp, 1, col = col_no_LON_MINUTE, row = i + 1, comment = createComment(comment = "Minutes must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LON_MINUTE, rows = i + 1)
        }
    }
    ##### herbdat_comments$LON_SECOND  col_no_LON_SECOND
    if(!is.na(herbdat_comments$LON_SECOND[i])){
        if(herbdat_comments$LON_SECOND[i] < 0 | herbdat_comments$LON_SECOND[i] > 60){
            writeComment(filled_temp, 1, col = col_no_LON_SECOND, row = i + 1, comment = createComment(comment = "Seconds must be between 0 and 90"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LON_SECOND, rows = i + 1)
        }
    }
    ##### herbdat_comments$LON_FLAG col_no_LON_FLAG 
    if(!is.na(herbdat_comments$LON_FLAG[i])){
        if(!herbdat_comments$LON_FLAG[i] %in% c("E", "W")){
            writeComment(filled_temp, 1, col = col_no_LON_FLAG, row = i + 1, comment = createComment(comment = "Longitudinal Flag must be either E or W"))
            addStyle(filled_temp, 1, bodyStyle, cols = col_no_LON_FLAG, rows = i + 1)
        }
    }
}

col_no_DATE_COLLECTED <- which(colnames(herbdat_comments) == "DATE_COLLECTED")
col_no_DATE_IDENTIFIED <- which(colnames(herbdat_comments) == "DATE_IDENTIFIED")

for (i in 1:nrow(herbdat_comments)){
#### DATE_COLLECTED
    if(!is.na(suppressWarnings(as.integer(herbdat_comments$DATE_COLLECTED[i])))){
        if(!grepl("^darwin", R.version$os)){
            x <- as.Date(as.integer(herbdat_comments$DATE_COLLECTED[i]), origin="1899-12-30")
        } else {
            x <- as.Date(as.integer(herbdat_comments$DATE_COLLECTED[i]), origin = "1904-01-01")
        } 
    } else {
            x <- herbdat_comments$DATE_COLLECTED[i]
    }
     
    if(x < as.Date("2000-1-1")){
        writeComment(filled_temp, 1, col = col_no_DATE_COLLECTED, row = i + 1, comment = createComment(comment = "Please check the date"))
        addStyle(filled_temp, 1, bodyStyle, cols = col_no_DATE_COLLECTED, rows = i + 1)
    }

    #### DATE_IDENTIFIED
    if(!is.na(suppressWarnings(as.integer(herbdat_comments$DATE_IDENTIFIED[i])))){
        if(!grepl("^darwin", R.version$os)){
            y <- as.Date(as.integer(herbdat_comments$DATE_IDENTIFIED[i]), origin="1899-12-30")
        } else {
            y <- as.Date(as.integer(herbdat_comments$DATE_IDENTIFIED[i]), origin = "1904-01-01")
        }
    } else {
            y <- herbdat_comments$DATE_IDENTIFIED[i]
    } 
    
    if(y < as.Date("2000-1-1")){
        writeComment(filled_temp, 1, col = col_no_DATE_IDENTIFIED, row = i + 1, comment = createComment(comment = "Please check the date"))
        addStyle(filled_temp, 1, bodyStyle, cols = col_no_DATE_IDENTIFIED, rows = i + 1)
    }
}

saveWorkbook(filled_temp, "herbarium_specimens_label_data.xlsx", overwrite = TRUE)

###############################################################################################
###############################################################################################
#### Update the data base
if(!dir.exists("DARWIN_CORE_DB_SAVE")){
    dir.create("DARWIN_CORE_DB_SAVE")
}

if(!file.exists("DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx")){
     temppp <- rep(NA, length(colnames(dwc_filled2)))
     temppp2 <- t(data.frame(temppp))
     colnames(temppp2) <- colnames(dwc_filled2)
     write.xlsx(temppp2, "DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx")
}

if(file.exists("DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx")){
    dat_dc_db <- read.xlsx("DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx")
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

write.xlsx(temp_dat_dc_db, paste("DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx", sep = ""))
invisible(file.copy(from = "DARWIN_CORE_DB_SAVE/darwin_core_database.xlsx", to = paste("DARWIN_CORE_DB_SAVE/", dat_tag, "_darwin_core_database_saved.xlsx", sep = "")))
