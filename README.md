## Welcome to the Homepage of run_herblabel

run_herblabel is a collection of files helping you to generate RTF herbarium labels and annotation labels using the R `herblabel` package (`devtools::install_github("helixcn/herblabel")`) on Windows 7 or above: 

Please download the files from `https://github.com/helixcn/run_herblabel/archive/master.zip` and unzip it to your computer. 

You need to have R as well as the following packages installed. Rtools `https://cran.r-project.org/bin/windows/Rtools/` should also be installed,and correctly configured allowing reading and processing of zip files. 

```R
install.packages("openxlsx") ### select a CRAN mirror close to you. 
install.packages("devtools")
install.packages("Rcpp")     ### Dependency of openxlsx
install.packages("memoise")  ### Dependency of openxlsx 
install.packages("digest")   ### Dependency of openxlsx
install.packages("withr")    ### Dependency of devtools
install.packages("httr")     ### Dependency of devtools
install.packages("R6")       ### Dependency of devtools
install.packages("curl")     ### Dependency of devtools
library(devtools)
install_github("helixcn/herblabel") 
```
## 1. Creating Herbarium Labels (on Windows)

(1). Open and fill different columns of the file "herbarium_specimens_label_data.xlsx". Put Scientific Name or Chinese Name in the column entitled "LOCAL_NAME".

(2). Double click "run.bat", and view the herbarium labels from "herbarium_labels_to_print.rtf". 

(3). The R script "print_label.R" will conduct data checking, including (1) spelling of genus, (2) family,  (3)genus - family relationship, (4) validity of the scientific names, (5) Whether the provinence is complete, (6) Whether the geographical coordinates are within the right range, and (8) the date of collection and identification are correct. Columns with errors will be highlighted, and a comments for the corresponding cells will be provided to help to point out the problems. 

### Important notice: 
Please close "herbarium_labels_to_print.rtf" and "herbarium_specimens_label_data.xlsx" before clicking "run.bat". 

## 2. Creating Annotation Labels (on Windows)

(1). Fill the Annotation_label_data.xlsx. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(2). Double click "print_annotation_labels_normal.bat" or "print_annotation_labels_without_family.bat" if you do not want the Family to be printed. 

(3). Check the annotation labels from "annotation_labels_to_print.rtf".

### Important notice: 
Please close "Annotation_label_data.xlsx" and "annotation_labels_to_print.rtf" before double clicking ".bat" file. 

Please feel free to send an email to the maintainer **Jinlong Zhang** <jinlongzhang01@gmail.com>.
