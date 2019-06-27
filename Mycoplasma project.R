#Load required packages
check.packages <- function(pkg){
    new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
    if (length(new.pkg)) 
        install.packages(new.pkg, dependencies = TRUE)
    sapply(pkg, require, character.only = TRUE)
}
packages <- c("googledrive", "dplyr", "stringr", "openxlsx")
check.packages(packages)

#Access the file from Google Drive and reformat data
setwd("~")
filename <- "NCATS Mycoplasma Testing Responses"
googledrive::drive_download(filename, type="csv", overwrite=TRUE)
responses <- read.csv(paste0(filename, ".csv"), na.strings=c("", NA), stringsAsFactors=FALSE)
responses <- responses[-which(is.na(responses[,1])),]
colnames(responses[,5:14]) <- rep("Sample", 10)
responses[,1] <- as.POSIXct(responses[,1], format="%d/%m/%Y %H:%M:%S")
responses <- responses[-which(responses[,1] <= as.POSIXct(Sys.Date()-7)),]  #remove rows >1 week old
responses[,1] <- as.character(responses[,1])

#Place each sample on its own row
res_list <- split(responses, seq(nrow(responses)))
res_list2 <- lapply(res_list, function(x) {
    sep_rows <- data.frame("Timestamp"=character(), "Email"=character(), "Contact"=character(),
                           "SampleNumber"=character(), "Sample"=character(), "Collab"=character(),
                            "CollabID"=character())
    for (i in 5:14) {
        if (!is.na(x[1, i])) {
            id <- colnames(x)[i]
            newrow <- cbind(x[1, 1:3], id, x[1, c(i, 15:16)])
            colnames(newrow) <- c("Timestamp", "Email", "Contact", "SampleNumber", "Sample", "Collab",
                                  "CollabID")
            sep_rows <- dplyr::bind_rows(sep_rows, newrow)
        }
    }
    return(sep_rows)
})

#Label the collaborator samples
res_list3 <- lapply(res_list2, function(x) {
    for (i in (1:nrow(x))) {
        rowtest <- x[i,4]
        if (!is.na(rowtest)) {
            if (grepl(gsub("\\D+", "", get("rowtest")),  x[i,7])) {
                x[i,6] <- TRUE
            } else {
                x[i,6] <- FALSE
            }
        } else {
            x[i,6] <- FALSE
        }
    }
    x <- x[,-c(4,7)]
    return(x)
})

#Append data to Mycoplasma Excel sheet
res_df <- do.call(rbind.data.frame, res_list3)
path <- "S:/Mycoplasma Test Results/MASTER MYCOPLASMA RESULTS.xlsx"
file_read <- openxlsx::read.xlsx(path)
file_load <- openxlsx::loadWorkbook(path)
sheet <- openxlsx::sheets(file_load)[[1]]
openxlsx::writeData(wb=file_load, sheet=sheet, x=res_df, startRow=nrow(file_read), colNames=FALSE)
openxlsx::saveWorkbook(file_load, file=path, overwrite=TRUE)
