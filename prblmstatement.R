#reading xlsx type file
install.packages("readxl") # CRAN version

#or

devtools::install_github("hadley/readxl") # development version


library("readxl", lib.loc="~/R/win-library/3.3")
all_names <- read_excel("C:/Users/chakradhar/Desktop/social cops/Book1.xlsx")
missing_data <- read_excel("C:/Users/chakradhar/Desktop/social cops/Book2.xlsx")
names(all_names) <- all_names[1,]
# Using dply for data frame operations
install.packages("dplyr")
library(dplyr)
# Considering only relevant columns for calculation
tab_all_data <- all_names[,c(4:6)]
# Sorting the dataset
tab_all_data <- tab_all_data %>% arrange(`Tab No`)
#Filtering for only those tabs where the data is missing
tab_all_data <- tab_all_data[tab_all_data$`Tab No` %in% missing_data$`Tab No`,]
tab_all_data <- tab_all_data[,c(1:5)]
# Creating a flag where the date is in the range of start and end date
tab_all_data$flag <- 0
tab_all_data$flag[(tab_all_data$`Survey Date`>=tab_all_data$`Survey Start Date`) & (tab_all_data$`Survey Date`<=tab_all_data$`Survey End Date`)]<-1
# Considering only those rows where the flag is 1 i.e. we found a match for a particular Tab
data_found <- tab_all_data[tab_all_data$flag ==1,c(1,4,2,3)]
names(data_found)[1] <- 'Tab No'
# Joining with all_names to get the village level details
data_found <- left_join(data_found,all_names,by=NULL)
# Joining the missing data with the final created dataset
missing_data <- left_join(missing_data,data_found,by=NULL)
missing_data <- missing_data[,c(1:3,7:11)]
# Hygeine changes in the final created dataset - For Tab's where the mapping is not done
# the values for columns are taken as 'Not Mapped'
missing_data$`AC Name`[is.na(missing_data$`AC Name`)] <- 'Not mapped'
missing_data$`Mandal Name`[is.na(missing_data$`Mandal Name`)] <- 'Not mapped'
missing_data$`Village Name`[is.na(missing_data$`Village Name`)] <- 'Not mapped'
missing_data$`Survey Start Date` <- as.character(missing_data$`Survey Start Date`)
missing_data$`Survey End Date` <- as.character(missing_data$`Survey End Date`)
missing_data$`Survey Start Date`[(is.na(missing_data$`Survey Start Date`))] <- 'Not mapped'
missing_data$`Survey End Date`[(is.na(missing_data$`Survey End Date`))] <- 'Not mapped'

missing_data$no_data_flag <- 0
missing_data$no_data_flag[missing_data$`AC Name` == 'Not mapped'] <- 1

#Checking the data if for same Tab No and survey data, we have more than one values
data_chk <- missing_data %>% filter(no_data_flag == 0) %>% group_by(Tab.No,Response.No) %>% summarise(count = length(Tab.No))

missing_data <- missing_data %>% arrange(Tab.No)

# Writing the final dataset created to the local
write_excel("C:/Users/chakradhar/Desktop/social cops/Missing_details.xlsx")

