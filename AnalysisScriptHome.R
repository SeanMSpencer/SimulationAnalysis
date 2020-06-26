library(readxl)
library(utils)
library(data.table)
library(writexl)
library(tictoc)

tic()

#choose directory for files to read in.
file.list <- list.files(choose.dir(), full.names = TRUE)


#creates list of dataframes to operate on.
#NOTE: function(x) is something is called a lambda or anonymous functions
df.list <- lapply(file.list, function(x) read_excel(x, col_types = "numeric"))

#specify colNames to applied to all data frames
col_Names <- c("Time UpCount",
               "UpCount",
               "Mission Ready",
               "MOS Timer",
               "MOS 1",
               "MOS 2",
               "MOS 3",
               "MOS 4",
               "MOS 5",
               "MOS 6",
               "MOS 7",
               "MOS 8",
               "MOS 9",
               "MOS 10",
               "Time Parts",
               "Parts A",
               "Parts B",
               "Parts C",
               "Parts D",
               "Parts E",
               "Parts F",
               "Parts G",
               "Parts H", 
               "Parts I",
               "Parts J")

#apply column names and transform "tibbles" to data frames to make them easier to work with
#I kept getting errors trying to pass tibbles to apply functions.
df.list <- lapply(df.list, setNames, nm = col_Names)
df.list <- lapply(df.list, as.data.frame)

#returns all statistics stated as a vector
multi.fun <- function(x) {
  
  c(mean = mean(x), median = median(x), min = min(x), max= max(x), var = var(x))
  
}

#apply the multi.fun function on the UpCount, MOS1, PartsA/B/C column and create a new dataframe with those colNames
upCount.df <- as.data.frame(t(sapply(df.list, function(x) multi.fun(x[["UpCount"]]))))
colnames(upCount.df) <- c("Mean UpCount", "Median UpCount", "Min UpCount", "Max UpCount", "Var UpCount")


MOS1.df <- as.data.frame(t(sapply(df.list, function(x) multi.fun(x[["MOS 1"]]))))
colnames(MOS1.df) <- c("Mean MOS1", "Median MOS1", "Min MOS1", "Max MOS1", "Var MOS1")

partsA.df <- as.data.frame(t(sapply(df.list, function(x) multi.fun(x[["Parts A"]]))))
colnames(partsA.df) <- c("Mean partsA", "Median partsA", "Min partsA", "Max partsA", "Var partsA")

partsB.df <- as.data.frame(t(sapply(df.list, function(x) multi.fun(x[["Parts B"]]))))
colnames(partsB.df) <- c("Mean partsB", "Median partsB", "Min partsB", "Max partsB", "Var partsB")

partsC.df <- as.data.frame(t(sapply(df.list, function(x) multi.fun(x[["Parts C"]]))))
colnames(partsC.df) <- c("Mean partsC", "Median partsC", "Min partsC", "Max partsC", "Var partsC")

#combind the resultant dataframes into one and use this to write out to an excel sheet
combined.df <- cbind(upCount.df, MOS1.df, partsA.df, partsB.df, partsC.df)

#write the file out too a blank excel file
write_xlsx(combined.df, choose.files())

toc()
