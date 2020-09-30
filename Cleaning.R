library(xlsx)
library(dplyr)
library(tidyverse)
filename <-
    "Coping Mechanism of Livestock Farmers during COVID-19 (Responses) (1).xlsx"
data <- read.xlsx(filename, sheetIndex = 1)
# subData <- data[22:]
group1raw <- data[22:26]
group2raw <- data[c(28:29, 33:34, 36)]
group3raw <- data[30:31]
group4raw <- data[37:43]
group5raw <- data[45:51]
group6raw <- data[52:59]

findPercentofGroups <- function(x) {
    #Convert to factor
    groupraw <- data.frame(lapply(x, factor,
                                  unique(c(t(
                                      x
                                  )))))
    
    groupsummary <- sapply(groupraw, table)
    groupsummary <- data.frame(t(groupsummary))
    
    #Calculate the total of each row
    groupsummary$total = rowSums(groupsummary)
    
    #Calculate percentages
    grppercent <- groupsummary %>% mutate_at(vars(!total)
                                             , .funs = ~ (. / total * 100))
    
    #Format percentages
    grppercent <- grppercent %>%
        mutate_at(-nrow(grppercent), list(~ paste(format(., digits = 3)
                                                  , "%", sep = "")))
    
    
    #To combine frequency and percentage for grp
    groupcomb <- groupsummary
    for (i in 1:nrow(groupcomb)) {
        for (j in 1:(ncol(groupcomb) - 1)) {
            groupcomb[i, j] <- paste(groupcomb[i, j], " (",
                                     grppercent[i, j], ")",
                                     sep = "")
        }
    }
    
    groupcomb
}

group1comb = findPercentofGroups(group1raw)
group2comb = findPercentofGroups(group2raw)
group3comb = findPercentofGroups(group3raw)
group4comb = findPercentofGroups(group4raw)
group5comb = findPercentofGroups(group5raw)
group6comb = findPercentofGroups(group6raw)


personalDetailsSummary <- data.frame(
    Variable = character(),
    Group = character(),
    frequency = character(),
    Percentage = character()
)

listFactors <- list(
    c("<20", "20-29", "30-39", "40-49", "50-59", "60-69", ">69"),
    c("Female", "Male"),
    c("< 50,000", "50,000 - 100,000", "> 100,000"),
    c("Primary", "Secondary", "Tertiary"),
    c("Islam", "Christianity", 'Traditional religion')
)

for (i in 5:9) {
    raw <- factor(data[[i]], listFactors[[i - 4]])
    
    summary <- table(raw)
    summary = data.frame(summary)
    #Add rows in level with frequency of 0
    temp = data.frame(lev = levels(raw)[!(levels(raw) %in% summary$raw)],
                      freq = numeric())
    
    summary$Variable = names(data[i])
    summary$Percent = signif(summary$Freq / sum(summary$Freq) * 100,
                             digits = 4)
    personalDetailsSummary <-
        rbind(personalDetailsSummary, summary[c(3, 1, 2, 4)])
}

write.xlsx(group1comb, "freqPercent.xlsx")
write.xlsx(group2comb,
           "freqPercent.xlsx",
           append = T,
           sheetName = "group2")
write.xlsx(group3comb,
           "freqPercent.xlsx",
           append = T,
           sheetName = "group3")
write.xlsx(group4comb,
           "freqPercent.xlsx",
           append = T,
           sheetName = "group4")
write.xlsx(group5comb,
           "freqPercent.xlsx",
           append = T,
           sheetName = "group5")
write.xlsx(group6comb,
           "freqPercent.xlsx",
           append = T,
           sheetName = "group6")
write.xlsx(
    personalDetailsSummary,
    "freqPercent.xlsx",
    append = T,
    sheetName = "personal"
)




#----------------------------
#Cleaning 2
#adjust livestock type to long yes/no

otherLivestock = gsub("Poultry", "", data[[10]])
otherLivestock = gsub("Piggery", "", otherLivestock)
#adjust leftover list for commas
otherLivestock = trimws(otherLivestock)
otherLivestock = gsub("(^,+( ,+)*( )*)", "", otherLivestock)

LivestockTypeLong <- data.frame(
    Livestock = data[[10]],
    Poultry = grepl("Poultry", data[[10]]),
    Piggery = grepl("Piggery", data[[10]]),
    Other = unname(sapply(otherLivestock, nchar) >
                       0 & !is.na(otherLivestock))
)

temp <- data.frame(lapply(LivestockTypeLong[2:4], sum))


#To convert farm enterprise details
listFactors <- list(
    c("Personally owned", "Employee", "Leased/rent"),
    c("<1yr", "1-5yrs", "5-10yrs", "above 10yrs"),
    c("Wholesales", "Individuals", "Government")
)
FarmDetails <- data.frame(
    Variable = character(),
    Group = character(),
    frequency = character(),
    Percentage = character()
)
index = c(16, 15, 17)
for (i  in 1:length(index)) {
    raw <- factor(data[[index[i]]], listFactors[[i]])
    
    summary <- table(raw)
    summary = data.frame(summary)
    #Add rows in level with frequency of 0
    temp = data.frame(lev = levels(raw)[!(levels(raw) %in% summary$raw)],
                      freq = numeric())
    
    summary$Variable = names(data[index[i]])
    summary$Percent = signif(summary$Freq / sum(summary$Freq) * 100,
                             digits = 4)
    
    FarmDetails <-
        rbind(FarmDetails, summary[c(3, 1, 2, 4)])
}


#To summarise a check box field
summaryCheckBoxes <- function(column, levels, has.other = F) {
    #Column is a list of the data frame column
    #Levels is the levels in the column
    TrueFalsedf <- data.frame()
    for (x in levels) {
        TrueFalsedf[x] = grepl(x, column)
    }
    #If there is an "other' field
    if (has.other == T) {
        other = column
        for(x in levels) other = gsub(x, "",other)
        
        #adjust leftover list for commas
        other = trimws(other)
        other = gsub("(^,+( ,+)*( )*)", "", other)
        TrueFalsedf["other"] = nchar(other) > 0 &
            !is.na(other)
    }
    
    summaryDF = lapply(TrueFalsedf, sum)
    summaryDF
}
