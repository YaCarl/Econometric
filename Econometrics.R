library("xlsx")

###ALTMAN Z-SCORE MODEL SCORES & TAFFLER'S MODEL SCORES for bankrupt companies:

setwd("/Users/irina/Desktop/companies_sample_2/bankrupt")
bankrupt_сomp <- list.files(pattern = "\\.xlsx$")
Z_bankrupt_2y <- c()
Z_bankrupt_1y <- c()
T_bankrupt_2y <- c()
T_bankrupt_1y <- c()
for (i in bankrupt_сomp) {
  
#Raw data on balance sheet and income statement
  data_BS<-read.xlsx(file.path(i),
                               "Balance sheet", 
                               rowIndex = 15:59, 
                               colIndex = 3:7, 
                               colNames=TRUE,rowNames=TRUE,stringsAsFactors=FALSE)
  data_IS<-read.xlsx(file.path(i),
                   "Income Statement", 
                   rowIndex = 15:33, 
                   colIndex = 3:7, 
                   colNames=TRUE,rowNames=TRUE,stringsAsFactors=FALSE)
  
#Preparation of the data
  data_BS <- data_BS[!data_BS$Code=="", ]
  data_BS <- data_BS[!is.na(data_BS$Code), ]
  data_BS_cl <- data_BS[-c(1,3,5,6,7)]
  data_BS_cl <- data.frame(lapply(data_BS_cl, function(x)gsub('\\s+','',x)),stringsAsFactors=FALSE)
  data_BS_cl <- data.frame(lapply(data_BS_cl, as.numeric))
  rownames(data_BS_cl) <- data_BS[,1]
  
  
  if (ncol(data_BS_cl)==2) {colnames(data_BS_cl) <- c(1,2)} else {colnames(data_BS_cl <- c(1))}

  data_IS <- data_IS[!data_IS$Code=="", ]
  data_IS_cl <- data_IS[-c(1,3,5,6,7)]
  data_IS_cl <- data.frame(lapply(data_IS_cl, function(x)gsub('\\s+','',x)),stringsAsFactors=FALSE)
  data_IS_cl <- data.frame(lapply(data_IS_cl, as.numeric))
  rownames(data_IS_cl) <- data_IS[,1]
  if (ncol(data_IS_cl)==2) {colnames(data_IS_cl) <- c(1,2)} else {colnames(data_IS_cl <- c(1))}

#Calculation of Z-scores for years 1 and 2:                                 
  Z_score_1y <- 3.25+6.56*(data_BS_cl["1200",1]-data_BS_cl["1500",1])/data_BS_cl["1600",1]+3.26*data_BS_cl["1370",1]/data_BS_cl["1600",1]+6.72*data_IS_cl["2300",1]/data_BS_cl["1600",1]+1.05*data_BS_cl["1300",1]/(data_BS_cl["1400",1]+data_BS_cl["1500",1])
  Z_score_2y <- 3.25+6.56*(data_BS_cl["1200",2]-data_BS_cl["1500",2])/data_BS_cl["1600",2]+3.26*data_BS_cl["1370",2]/data_BS_cl["1600",2]+6.72*data_IS_cl["2300",2]/data_BS_cl["1600",2]+1.05*data_BS_cl["1300",2]/(data_BS_cl["1400",2]+data_BS_cl["1500",2])
  T_score_1y <- 0.53*data_IS_cl["2300",1]/data_BS_cl["1500",1]+0.13*data_BS_cl["1200",1]/data_BS_cl["1700",1]+0.18*data_BS_cl["1500",1]/data_BS_cl["1600",1]+0.16*data_IS_cl["2110",1]/data_BS_cl["1600",1]
  T_score_2y <- 0.53*data_IS_cl["2300",2]/data_BS_cl["1500",2]+0.13*data_BS_cl["1200",2]/data_BS_cl["1700",2]+0.18*data_BS_cl["1500",2]/data_BS_cl["1600",2]+0.16*data_IS_cl["2110",2]/data_BS_cl["1600",2]
  
  Z_bankrupt_2y <- c(Z_bankrupt_2y,Z_score_2y)
  Z_bankrupt_1y <- c(Z_bankrupt_1y,Z_score_1y)
  T_bankrupt_2y <- c(T_bankrupt_2y,T_score_2y)
  T_bankrupt_1y <- c(T_bankrupt_1y,T_score_1y)}
                                  
#Extra cleaning specifically for this sample of companies since some of the files had different formatting
T_bankrupt_1y <- na.omit(T_bankrupt_1y)


###ALTMAN Z-SCORE MODEL SCORES & TAFFLER'S MODEL SCORES for non-bankrupt companies:
#The structure of the following code is the same as of the previous one:
setwd("/Users/irina/Desktop/companies_sample_2/nonbankrupt")
nonbankrupt_сomp <- list.files(pattern = "\\.xlsx$")
Z_nonbankrupt_2y <- c()
Z_nonbankrupt_1y <- c()
T_nonbankrupt_2y <- c()
T_nonbankrupt_1y <- c()
for (i in nonbankrupt_сomp) {
  data_BS<-read.xlsx(file.path(i),
                     "Balance sheet", 
                     rowIndex = 15:59, 
                     colIndex = 3:7, 
                     colNames=TRUE,rowNames=TRUE,stringsAsFactors=FALSE)
  data_IS<-read.xlsx(file.path(i),
                     "Income Statement", 
                     rowIndex = 15:33, 
                     colIndex = 3:7, 
                     colNames=TRUE,rowNames=TRUE,stringsAsFactors=FALSE)
  
  data_BS <- data_BS[!data_BS$Code=="", ]
  data_BS <- data_BS[!is.na(data_BS$Code), ]
  data_BS_cl <- data_BS[-c(1,3,5,6,7)]
  data_BS_cl <- data.frame(lapply(data_BS_cl, function(x)gsub('\\s+','',x)),stringsAsFactors=FALSE)
  data_BS_cl <- data.frame(lapply(data_BS_cl, as.numeric))
  rownames(data_BS_cl) <- data_BS[,1]
  
  
  if (ncol(data_BS_cl)==2) {colnames(data_BS_cl) <- c(1,2)} else {colnames(data_BS_cl <- c(1))}
  
  data_IS <- data_IS[!data_IS$Code=="", ]
  data_IS_cl <- data_IS[-c(1,3,5,6,7)]
  data_IS_cl <- data.frame(lapply(data_IS_cl, function(x)gsub('\\s+','',x)),stringsAsFactors=FALSE)
  data_IS_cl <- data.frame(lapply(data_IS_cl, as.numeric))
  rownames(data_IS_cl) <- data_IS[,1]
  if (ncol(data_IS_cl)==2) {colnames(data_IS_cl) <- c(1,2)} else {colnames(data_IS_cl <- c(1))}
  
  Z_score_1y <- 3.25+6.56*(data_BS_cl["1200",1]-data_BS_cl["1500",1])/data_BS_cl["1600",1]+3.26*data_BS_cl["1370",1]/data_BS_cl["1600",1]+6.72*data_IS_cl["2300",1]/data_BS_cl["1600",1]+1.05*data_BS_cl["1300",1]/(data_BS_cl["1400",1]+data_BS_cl["1500",1])
  Z_score_2y <- 3.25+6.56*(data_BS_cl["1200",2]-data_BS_cl["1500",2])/data_BS_cl["1600",2]+3.26*data_BS_cl["1370",2]/data_BS_cl["1600",2]+6.72*data_IS_cl["2300",2]/data_BS_cl["1600",2]+1.05*data_BS_cl["1300",2]/(data_BS_cl["1400",2]+data_BS_cl["1500",2])
  T_score_1y <- 0.53*data_IS_cl["2300",1]/data_BS_cl["1500",1]+0.13*data_BS_cl["1200",1]/data_BS_cl["1700",1]+0.18*data_BS_cl["1500",1]/data_BS_cl["1600",1]+0.16*data_IS_cl["2110",1]/data_BS_cl["1600",1]
  T_score_2y <- 0.53*data_IS_cl["2300",2]/data_BS_cl["1500",2]+0.13*data_BS_cl["1200",2]/data_BS_cl["1700",2]+0.18*data_BS_cl["1500",2]/data_BS_cl["1600",2]+0.16*data_IS_cl["2110",2]/data_BS_cl["1600",2]
  Z_nonbankrupt_2y <- c(Z_nonbankrupt_2y,Z_score_2y)
  Z_nonbankrupt_1y <- c(Z_nonbankrupt_1y,Z_score_1y)
  T_nonbankrupt_2y <- c(T_nonbankrupt_2y,T_score_2y)
  T_nonbankrupt_1y <- c(T_nonbankrupt_1y,T_score_1y)}
  Z_bankrupt_1y <- na.omit(Z_bankrupt_1y)
