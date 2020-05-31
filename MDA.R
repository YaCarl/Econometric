library(tidyverse)

Z_model<- data.frame("status"=integer(),"X1"=double(),"X2"=double(),"X3"=double(),"X4"=double())
observations_vector <- c()
setwd("/Users/irina/Desktop/companies_sample_2/bankrupt")
bankrupt_сomp <- list.files(pattern = "\\.xlsx$")
for (i in bankrupt_сomp) {
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
  observations_vector <- c(1,
                           (data_BS_cl["1200",1]-data_BS_cl["1500",1])/data_BS_cl["1600",1],
                           data_BS_cl["1370",1]/data_BS_cl["1600",1],
                           data_IS_cl["2300",1]/data_BS_cl["1600",1],
                           data_BS_cl["1300",1]/(data_BS_cl["1400",1]+data_BS_cl["1500",1]))
  Z_model <- rbind(Z_model,observations_vector)}

setwd("/Users/irina/Desktop/companies_sample_2/nonbankrupt")
bankrupt_сomp <- list.files(pattern = "\\.xlsx$")
observations_vector<-c()
for (i in bankrupt_сomp) {
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
  observations_vector <- c(0,
                           (data_BS_cl["1200",1]-data_BS_cl["1500",1])/data_BS_cl["1600",1],
                           data_BS_cl["1370",1]/data_BS_cl["1600",1],
                           data_IS_cl["2300",1]/data_BS_cl["1600",1],
                           data_BS_cl["1300",1]/(data_BS_cl["1400",1]+data_BS_cl["1500",1]))
  Z_model <- rbind(Z_model,observations_vector)}
Z_model <- na.omit(Z_model)
colnames(Z_model) <- c("status","X1","X2","X3","X4")

#The following code was borrowed from https://www.rpubs.com/vijetk/azscore:
reg <- lm(status~X1+X2+X3+X4, Z_model)
summary(reg)
groupcount=Z_model %>% 
  group_by(status) %>% tally()

groupcount$percentage=groupcount$n/sum(groupcount$n)
Z_model$predicted=predict(reg)

avgprdtd=Z_model %>% 
  group_by(status) %>% 
  summarise(mean(predicted))

avgprdtd
cutoff=(groupcount[1,2]*avgprdtd[1,2]+groupcount[2,2]*avgprdtd[2,2])/sum(groupcount[,2])
cutoff=c(cutoff)
cutoff
Z_model$grp_prd=ifelse(Z_model$predicted<cutoff,1,0)

Z_model$accu=ifelse(Z_model$status==Z_model$grp_prd,1,0)

accu_cnt=Z_model %>% 
  group_by(accu) %>% 
  tally()

accu_cnt$percentage=accu_cnt$n/sum(accu_cnt$n)

accu_cnt
accuracy=accu_cnt[2,2]/sum(accu_cnt[,2])
accuracy
