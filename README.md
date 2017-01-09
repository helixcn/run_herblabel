## Welcome to the Homepage of run_herblabel

run_herblabel is a collection of files helping you to generate RTF herbarium labels and annotation labels using the R `herblabel` package: 

Please download the files from this page or [http://blog.sciencenet.cn/blog-255662-849868.html](http://blog.sciencenet.cn/blog-255662-849868.html) and unzip it to your computer. 

You need to have R as well as the following packages installed. Rtools `[https://cran.r-project.org/bin/windows/Rtools/](https://cran.r-project.org/bin/windows/Rtools/)` should also be installed,and correctly configured allowing reading and processing of zip files. 

```R
install.packages("herblabel", repos="http://R-Forge.R-project.org")
install.packages("openxlsx") ### select a CRAN mirror close to you. 
install.packages("Rcpp")     ### Dependency of openxlsx
```

## Configuration of Rtools
Please allow the Rtools to change your system path. 
Please also add the following path to your system path, so the Rscript.exe could be invoked, remember to change the version of R if neccessary: 
";C:\Program Files\R\R-3.3.1\bin;C:\Program Files\R\R-3.3.1\bin\i386;" 

## 1. Creating Herbarium Labels

(1). Open and fill different columns of the file "herbarium_specimens_label_data.xlsx". Put Scientific Name or Chinese Name in the column entitled "LOCAL_NAME".

(2). Double click "run.bat", and view the herbarium labels from "herbarium_labels_to_print.rtf". 

(3). The R script "print_label.R" will conduct data checking, including (1) spelling of genus, (2) family,  (3)genus - family relationship, (4) validity of the scientific names, (5) Whether the provinence is complete, (6) Whether the geographical coordinates are within the right range, and (8) the date of collection and identification are correct. Columns with errors will be highlighted, and a comments for the corresponding cells will be provided to help to point out the problems. 

## 2. Creating Annotation Labels

(1). Fill the Annotation_label_data.xlsx. Put Scientific Name or Chinese Name at the column "LOCAL_NAME".

(2). Double click "print_annotation_labels_normal.bat" or "print_annotation_labels_without_family.bat" if you do not want the Family to be printed. 

(3). Check the annotation labels from "annotation_labels_to_print.rtf".

### Important notice: 
Please close ".xlsx" and ".rtf" files before double clicking ".bat" file. 

Please feel free to send an email to the maintainer **Jinlong Zhang** <jinlongzhang01@gmail.com>.
