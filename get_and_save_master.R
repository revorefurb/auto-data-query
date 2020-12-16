######
#install.packages("openxlsx")
#install.packages("ROracle")
#install.packages("RMySQL")
######

setwd("C:/Users/J11/Dropbox/R/Project/auto-data-query/")
day <- format(Sys.time(), "%A")
today <- format(Sys.time(), "%Y-%m-%d")

# open the conf file which must named conf.xlsx and be stored in the script directory
library(openxlsx)
conf <- read.xlsx("conf.xlsx", sheet = 1, startRow = 1, colNames = TRUE)

i <- 1
sched <- conf[i,3]

# cycle through the list in the conf file and stop at the end
condition1<-conf[which(conf[,1]=="TRUE"),]
# only run the query if it is scheduled to be run today
condition2<-condition1[which(grepl(day,condition1[,3])),]

for (i in 1:nrow(condition2)){

    # read conditions in conf file
    index <- conf[i,2]
    queryFileName <- conf[i,5]
    preFileName <- conf[i,6]
    dataFileName <- conf[i,7]
    saveFileName <- paste(preFileName, dataFileName, today, sep="_")
    saveLoc <- conf[i,8]
    dbType <- conf[i,4]
    dbName <- conf[i,9]
    dbUser <- conf[i,10]
    dbPass <- conf[i,11]
    
    print(paste("Attempting to run query --> ", queryFileName, sep=""))  
    queryText <- readChar(queryFileName, file.info(queryFileName)$size)
    
    if (dbType == "Oracle"){
      library('ROracle')
      drv <- dbDriver("Oracle")
    }
    if (dbType == "MySQL"){
      library('RMySQL')
      drv <- dbDriver("MySQL")
    }
      con <- dbConnect(drv, dbUser,dbPass, dbname=dbName)
      query <- dbSendQuery(con, queryText)
      data <- fetch(query)
  
  print(paste("Attempting to save to loc --> ", saveLoc, sep=""))  
  write.csv(data, file = paste(saveLoc, saveFileName,".csv", sep=""), row.names = FALSE)
  dbClearResult(query)
  dbUnloadDriver(drv)

  
  rm(list = c('index','queryFileName','preFileName','dataFileName','saveFileName','saveLoc','dbType','dbName','dbUser','dbPass','queryText','drv','con','query','data'))
  i <<- i+1
  print(i)
  sched <<- conf[i,3]
}

rm(list=ls())
