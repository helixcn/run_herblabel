## Welcome to the Homepage of run_herblabel

run_herblabel is a collection of files helping you to generate RTF herbarium labels and annotation labels using the R `herblabel` package (`devtools::install_github("helixcn/herblabel")`): 

Please download the files as a Master File from `https://github.com/helixcn/run_herblabel/archive/master.zip`
and unzip the `master.zip` to your computer. 

You need to have R installed and also the following packages installed. 

```R
install.packages("openxlsx")
install.packages("devtools")
install.packages("Rcpp")     ### Dependency of openxlsx
install.packages("memoise")  ### Dependency of openxlsx 
install.packages("digest")   ### Dependency of openxlsx
install.packages("withr")    ### Dependency of devtools
install.packages("httr")     ### Dependency of devtools
install.packages("R6")       ### Dependency of devtools
install.packages("curl")     ### Dependency of devtools

### select the a CRAN mirror close to you. 
library(devtools)
install_github("helixcn/herblabel")
```

## 1. Creating Herbarium Labels

(1). Open and Fill the "herbarium_specimens_label_data.xlsx" template. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(2). Double click "run_FRPS.bat", and view the herbarium labels from "herbarium_labels_to_print.rtf".

### Note: 
Please close "herbarium_labels_to_print.rtf" or "herbarium_specimens_label_data.xlsx" before printing. 

## 2. Creating Annotation Labels

(1). Fill the Annotation_label_data.xlsx. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(2). Double click "print_annotation_labels_normal.bat" or "print_annotation_labels_without_family.bat" if you do not want the Family to be printed. 

(3). Check the annotation labels from "annotation_labels_to_print.rtf".

### Note: 
Please close "Annotation_label_data.xlsx" or "annotation_labels_to_print.rtf" before printing. 

Please feel free to send an email to the maintainer **Jinlong Zhang** <jinlongzhang01@gmail.com>.
