### Mycoplasma Project, Pre-Google Form

#Load packages
check.packages <- function(pkg){
    new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
    if (length(new.pkg)) 
        install.packages(new.pkg, dependencies = TRUE)
    sapply(pkg, require, character.only = TRUE)
}
packages <- c("openxlsx", "data.table")
check.packages(packages)

##Format #1: 6/17/2016 to 6/29/2018

#Read files

ls <- list()
for (filename in list.files()) {
    file <- openxlsx::read.xlsx(filename, startRow=3, colNames=FALSE, skipEmptyRows=FALSE)
    
    names <- file[1:8, c(1:12)]
    results <- file[12:19, 1:12]
    results <- sapply(results, as.numeric)
    
    nlist <- unlist(names)
    nlist <- nlist[!is.na(nlist)]
    
    #Match values with names
    combined <- data.frame(Name=nlist, Ratio=NA, Result=NA)
    avg <- vector()
    for (i in c(1,4,7,10)) {
        for (row in 1:nrow(names)) {
            if (!is.na(names[row, i])) {
                ratio <- (mean(results[row, i], results[row, i+1], results[row, i+2], na.rm=TRUE, trim=0))
                avg <- c(avg, ratio)
            }
        }
    }
    combined$Ratio <- avg
    
    for (i in 1:nrow(combined)) {
        if (combined[i,2] > 1.2) {
            combined[i,3] <- "+"
        } else if (combined[i,2] <= 1.9 & combined[i,2] >= 0.9) {
            combined[i,3] <- "?"
        } else if (combined[i,2] < 0.9) {
            combined[i,3] <- "-"
        }
    }
    combined <- combined[!grepl("control|^x$", combined[,1], ignore.case=TRUE),]
    index <- match(filename, list.files())
    ls[[index]] <- combined
    names(ls)[index] <- filename
}

##Format #2: 9/11/2015 to 6/3/2016

ls2 <- list()
for (filename in list.files()) {
    file <- openxlsx::read.xlsx(filename, colNames=FALSE)
    
    names <- file[29:36, c(1:12)]
    results <- file[20:27, 1:12]
    results <- sapply(results, as.numeric)
    
    nlist <- unlist(names)
    nlist <- nlist[!is.na(nlist)]
    
    #Match values with names
    combined <- data.frame(Name=nlist, Ratio=NA, Result=NA)
    combined <- combined[!grepl("control|^x$|empty", combined[,1], ignore.case=TRUE),]
    avg <- vector()
    for (i in c(1,4,7,10)) {
        for (row in 1:nrow(names)) {
            if (!is.na(names[row, i]) & !grepl("control|^x$|empty", names[row, i], ignore.case=TRUE)) {
                ratio <- (mean(results[row, i], results[row, i+1], results[row, i+2], na.rm=TRUE, trim=0))
                avg <- c(avg, ratio)
            }
        }
    }
    combined$Ratio <- avg
    
    for (i in 1:nrow(combined)) {
        if (combined[i,2] > 1.2) {
            combined[i,3] <- "+"
        } else if (combined[i,2] <= 1.9 & combined[i,2] >= 0.9) {
            combined[i,3] <- "?"
        } else if (combined[i,2] < 0.9) {
            combined[i,3] <- "-"
        }
    }
    index <- match(filename, list.files())
    ls2[[index]] <- combined
    names(ls2)[index] <- filename
}

##Export file
df <- do.call("rbind", ls)
df <- rbind(df, do.call("rbind", ls2))
df <- data.table::setDT(df, keep.rownames = TRUE)[]
colnames(df)[1] <- "File name"

wb <- openxlsx::createWorkbook("myco.xlsx")
openxlsx::addWorksheet(wb, sheetName=1)
openxlsx::writeData(wb, sheet=1, df)
openxlsx::saveWorkbook(wb, "myco.xlsx", overwrite=TRUE)
