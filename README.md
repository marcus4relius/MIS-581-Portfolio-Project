##The tidyverse is a collection of packages in R Studio that share the same syntax and are designed for data analysis.  
##Core packages in the tidyverse include ggplot2, dplyr, tidyr, and stringr.  Analysis performed in this study will use
##several functions that are included in the tidyverse collection.
library(tidyverse)

##The data set that is being analyzed is located on remotely and on the local disk.  The data set is in an Excel file format.
##To import this file into R Studio the readx1 package will be installed.  This will enable the function read_excel to be used
##to load data set into the workspace.
library(readxl)
##The loading of the data set into R Studio.
Std.OM <- read_excel("MIS 581/emdi 1.1.9.xlsx")

##Determining the dimensions of the data set.  This shows the number of columns and rows and provides a sense of how large the data set is
##and how much work may be required to prepare the data for analysis.  
dim(Schlage.df)

##Identifying what the different variable names are in the Std.OM data set.  This will better understand what variables exist
##in the data set and how they might be grouped together.  It will also help to identify those variable names that may need to be
##changed to ensure the anonymization of the data to protect internal organization information.
names(Std.OM)

##Printing column names in list format.  This view may be easier to view the variables in the data set.  
t(t(names(Std.OM)))

##Running the summary statistics for the Order Management data set.  Using the pastecs package the stat.desc function provides the median,
##mean, standard error, confidence interval of the mean, variance, standard deviation, and the variation coefficient.  Values were set to 
##round to the second decimal place.
install.packages("pastecs")
library(pastecs)
res <- stat.desc(Std.OM[, -5])
round(res, 2)

#As shown by the summary statistics, some of the variables define production goals and do  not change over time.  As they are
##not linked to any work activity they will be removed to clean up the data set.
Std.OM1 <- Std.OM %>% select(-one_of('OT Goal', 'OT Under Goal', 'Backlog Goal', 'Under Backlog Goal', 'KC'))

##The names function will be run again to confirm that desired columns have been removed.
names(Std.OM1)

##As the Date Reported field is not a numerical value it will be removed from this data set for this current series of analysis.  Date reported will 
##be used in future analysis in identifying any potential seasonality to production or other variables in data set.
Std.OM2 <- Std.OM1 %>% select(-one_of('Date Reported'))

##The summary statistics also showed that some of the variables were missing information beginning at different points in the data collection.  To 
##ensure that all of the variables being analyzed have the same number of data points all rows missing information will be removed for the purpose of
##this study.  Future analysis will look at what data has been omitted and determine how they can be used to supplement the results of this project.
Std.OM3 <- Std.OM2[complete.cases(Std.OM2), ]

##The Order Management Department data set includes several teams that work together yet their processes and the orders they manage remain distinct.  
##As the standard team production accounts for 80% of the total department production the focus will be on this team's KPIs.  This will mean the 
##removal of columns from the other teams with only the standard team's remaining for analysis.
Std.OM4 <- Std.OM3 %>% select(-contains('KC', 'MK', 'Master', 'POP', 'International'))

##Confirming that only the Standard OE Team's variables remain
names(Std.OM4)

##As a part of the process of ensuring data privacy for the organization's internal production data the specific names in the data set will be
##anonymized using single letters.  
names(Std.OM4)[names(Std.OM4) == "Schlage Std"] <- "S Std"
names(Std.OM4)[names(Std.OM4) == "Falcon / Electronics Std"] <- "F/E S"
names(Std.OM4)[names(Std.OM4) == "Falcon C2B SLA"] <- "F C2B SLA"
names(Std.OM4)[names(Std.OM4) == "Standard C2B SLA"] <- "S C2B SLA"
names(Std.OM4)[names(Std.OM4) == "Standard C2B Days to 95%"] <- "S C2B Days to 95%"
names(Std.OM4)[names(Std.OM4) == "11i Orders Booked"] <- "EI Orders Booked"
names(Std.OM4)[names(Std.OM4) == "11i Lines Booked"] <- "EI Lines Booked"
names(Std.OM4)[names(Std.OM4) == "Std Orders Booked"] <- "S Orders Booked"
names(Std.OM4)[names(Std.OM4) == "Std Lines Booked"] <- "S Lines Booked"

##Cluster analysis will be performed using the heat map function that will color code how strong the correlation is between the production variables
##and the dependent variables in the data set. The first group of KPI's tested for a correlational analysis are the number of the individual team's unbooked orders.
corrplot::corrplot(cor(select(Std.OM4, "S Std", "F/E S", "Total EI orders", Total, "Orders Booked", "Lines Booked")))
##
corrplot::corrplot(cor(select(Std.OM4, "F C2B SLA", "S C2B SLA", "S C2B Days to 95%", "Total EI orders", "Orders Booked", "Lines Booked")))
##
corrplot::corrplot(cor(select(Std.OM4, "EI Orders Booked", "EI Lines Booked", "S Orders Booked", "S Lines Booked", "Std OT", "Orders Booked", "Lines Booked")))

##From the heat map analysis there are two variables that have a strong positive correlation with the number of Orders and Lines Booked.  To help determine if the
##two variables have a linear relationship to predict future production regression analysis was performed for these two variables versus number of orders booked.
lm(formula = Std.OM4$'Orders Booked' ~ Std.OM4$'Std OT' + Std.OM4$'Total Incoming Order #')

##From the linear model analysis only Overtime had a statistically significant linear relationship with the number of orders booked.  To show this relationship
##to help leadership better visualize the expected production returns for each hour of overtime worked a scatter plot with a best fit line was plotted.
plot(Std.OM4$'Std OT', Std.OM4$'Orders Booked', main = "Standard Overtime versus Orders Booked", xlab = "Standard Overtime", ylab = "Orders Booked", pch = 19)
abline(lm(Std.OM4$'Std Ot' ~ Std.OM4$'Orders Booked'), col = "blue")

##As overtime can have a positive impact on the number of orders produced, understanding if there is a statistically significant negative linear relationship with
##those variables that have a strong negative correlation to product can help predict the deleterious impact of the order entry processes.
lm(formula = Std.OM4$'Orders Booked' ~ Std.OM4$Diff + Std.OM4$'Lines / order avg' + Std.OM4'S C2B SLA' + Std.OM4$'Total Std SLA Orders')

##To further investigate the relationship between the amount of overtime worked and production a two sample t-test was performed, comparing the means for the 
##number of orders booked at or below 20 hours of overtime worked compared to the average number of orders booked with above 20 hours of overtime worked.
G1 <- subset(Schlage.df, Schlage.df$`Std OT` <= 20)
G2 <- subset(Schlage.df, Schlage.df$`Std OT` > 20)

##A variance test will show whether the var.equal statement of the t-test should be set as TRUE or FALSE.  
var.test(c1$`Std Orders Bkd`, c2$`Std Orders Bkd`)

##Since the variance showed as equal the t-test will be performed with the var.equal statement set as TRUE.
t.test(c1$`Std Orders Bkd`, c2$`Std Orders Bkd`, var.equal = TRUE)

##The results of the t-test showed that the two means were statistically significantly different.  To see if a different threshold will show two means that 
##are not significantly different the cutoff for overtime hours was set to greater or lesser than 25 hours.  
G3 <- subset(Schlage.df, Schlage.df$`Std OT` <= 25)
G4 <- subset(Schlage.df, Schlage.df$`Std OT` > 25)

##A variance test will be performed again to determine if the variances of the two samples is equivalent.
var.test(c3$`Std Orders Bkd`, c4$`Std Orders Bkd`)

##As the variances are not equal the var.equal statement of the t-test will be set as FALSE.  This will run a Welch's t-test which is used when the variances
##of the two samples are not equivalent.
t.test(c3$`Std Orders Bkd`, c4$`Std Orders Bkd`, var.equal = FALSE)
