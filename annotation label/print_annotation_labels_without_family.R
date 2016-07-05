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

invisible(library(openxlsx))
invisible(library(herblabel))

invisible(Sys.setlocale("LC_TIME", "C"))

if(!file.exists("Annotation_label_data.xlsx")){
   stop("The template file \"Annotation_label_data.xlsx\" can not be found in this directory")   
}

### Save a copy to the history folder
### The time marker
dat_tag <- gsub(":", "", gsub(" ", "-", paste(Sys.time())))

if(!dir.exists("xlsx history")){
    dir.create("xlsx history")
} 

invisible(file.copy(from = "Annotation_label_data.xlsx", to = paste("xlsx history/", dat_tag, "_Annotation_label_data.xlsx", sep = "")))


if(!dir.exists("RTF history")){
    dir.create("RTF history")
} 

invisible(file.copy(from = "annotation_labels_to_print.rtf", to = paste("RTF history/", dat_tag, "_annotation_labels_to_print.rtf", sep = "")))

dat_test <- read.xlsx("Annotation_label_data.xlsx")

dwc_filled <- fill_dwc(dat_test)
## unlink("Annotation_label_data.xlsx")

#### Fill the dataset, edit the herbarium_specimens_label_data file
write.xlsx(x = dwc_filled, file = "Annotation_label_data.xlsx")

dwc_filled2 <- read.xlsx("Annotation_label_data.xlsx")

#### Create the labels for checking or printing
#### 
dwc_filled2$FAMILY <- rep("", nrow(dwc_filled))
annotation_label(dat = dwc_filled2, outfile = paste("annotation_labels_to_print.rtf"))

