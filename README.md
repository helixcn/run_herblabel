## Welcome to the Homepage of run_herblabel

run_herblabel is a collection of files helping you to generate RTF herbarium labels and annotation labels using the R `herblabel` package (`devtools::install_github("helixcn/herblabel")`): 

Please download the files as a Master File from `https://github.com/helixcn/run_herblabel/archive/master.zip`
and unzip the `master.zip` to your computer. 

You need to have R installed and also the following packages installed. 

```R
install.packages("openxlsx")
install.packages("devtools")
### select the a CRAN mirror close to you. 
library(devtools)
install_github("helixcn/herblabel")
```

## 1. Creating Herbarium Labels (on Windows)

(1). Export Specimen Information from BG-BASE specimen table using slist by copying the commands in the file "step 0 export specimen records from BG_BASE using SLIST.txt" from BG-BASE. A CSV file named "BG_BASE_EXPORT.CSV" will be generated in your directory.

(2). Double Click "step 1 convert_BG-BASE.bat" to convert this CSV to darwin format template. After converting to DARWIN Core Format, the csv file "BG_BASE_EXPORT.CSV" will be renamed to "CONVERTED_BG_BASE_EXPORT.CSV".

(3). Check and Fill the "herbarium_specimens_label_data.xlsx" template. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(4). Double click "step 2 fill template and print_labels.bat", and view the herbarium labels from "herbarium_labels_to_print.rtf".

### Note: 
(1) Please close "herbarium_labels_to_print.rtf" or "herbarium_specimens_label_data.xlsx" before printing. 

(2) Please fill in the herbarium_specimens_label_data.xlsx template directly if your collection record is not from BG-BASE, and then click "step 2 fill template and print_labels.bat".

(3) Please change the property of the cells containing dates as "text", so Excel will not change the date to numbers.


## 2. Creating Annotation Labels (on Windows)

(1). Fill the Annotation_label_data.xlsx. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(2). Double click "print_annotation_labels_normal.bat" or "print_annotation_labels_without_family.bat" if you do not want the Family to be printed. 

(3). Check the annotation labels from "annotation_labels_to_print.rtf".

### Note: 

(1). Please close "Annotation_label_data.xlsx" or "annotation_labels_to_print.rtf" before printing. 
    
(2). Please change the property of the cells containing dates as "text", so excel will not change the text to numbers.

Please feel free to send an email to the maintainer **Jinlong Zhang** <jinlongzhang01@gmail.com>.
