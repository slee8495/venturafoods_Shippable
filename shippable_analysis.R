# https://damio.tistory.com/55?category=1172750
# https://rpubs.com/jmhome/R_data_wrangling


library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)
library(reshape2)



options(scipen = 100000000)


# Read model_1 ----

model_1 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Desktop/model_1.xlsx", 
                      col_types = c("text", "text", "numeric", 
                                    "numeric", "numeric", "numeric", 
                                    "numeric"))

model_1[-1:-3, ] -> model_1
model_1[is.na(model_1)] <- 0

colnames(model_1)[1] <- "Location"
colnames(model_1)[2] <- "Sku"
colnames(model_1)[3] <- "PastDate"
colnames(model_1)[4] <- "Next15Days"
colnames(model_1)[5] <- "Next16_30Days"
colnames(model_1)[6] <- "Next30_60Days"
colnames(model_1)[7] <- "ABOVE60"

model_1 %>% 
  dplyr::mutate(ref = paste(Location, "_", Sku)) %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) %>% 
  dplyr::relocate(ref) %>% 
  dplyr::mutate(nchar = nchar(ref)) %>% 
  dplyr::filter(nchar != 9) %>% 
  dplyr::mutate(ref = gsub("-", "", ref)) %>% 
  dplyr::mutate(Sku = gsub("-", "", Sku)) %>% 
  dplyr::select(-nchar) -> model_1

model_1 %>% 
  dplyr::mutate(SumOfNext = Next15Days + Next16_30Days + Next30_60Days + ABOVE60) %>% 
  dplyr::mutate(PercentOfPastDate = PastDate / SumOfNext) -> model_1

model_1[is.na(model_1)] <- 0

model_1


# Read model_2 ----

model_2 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Desktop/model_2.xlsx", 
                      col_types = c("text", "text", "text", 
                                    "text", "numeric"))

model_2[-1:-2,] -> model_2


colnames(model_2)[1] <- "Location"
colnames(model_2)[2] <- "CustomerShipTo"
colnames(model_2)[3] <- "CustomerShipToName"
colnames(model_2)[4] <- "Sku"
colnames(model_2)[5] <- "CurrentOpenOrderCases"


model_2 %>% 
  dplyr::mutate(ref = paste(Location, "_", Sku)) %>% 
  dplyr::mutate(ref = gsub(" ", "", ref)) %>% 
  dplyr::relocate(ref) %>%  
  dplyr::mutate(ref = gsub("-", "", ref)) %>% 
  dplyr::mutate(Sku = gsub("-", "", Sku)) %>% 
  dplyr::mutate(CustomerShipTo = gsub("-", "_", CustomerShipTo)) -> model_2


# Read model_3 ----

model_3 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Desktop/model_3.xlsx")

model_3[-1:-2,] -> model_3

colnames(model_3)[1] <- "CustomerShipTo"
colnames(model_3)[2] <- "Location"
colnames(model_3)[3] <- "PercentOfShelfLife"

model_3 %>% 
  dplyr::mutate(CustomerShipTo = gsub("-", "_", CustomerShipTo)) -> model_3


# Vlookup ----
merge(model_2, model_1[, c("ref", "PastDate")], by = "ref", all.x = TRUE) -> model_2
merge(model_2, model_1[, c("ref", "SumOfNext")], by = "ref", all.x = TRUE) -> model_2
merge(model_2, model_1[, c("ref", "PercentOfPastDate")], by = "ref", all.x = TRUE) -> model_2

colnames(model_2)[7] <- "NotShippableAmount"
colnames(model_2)[8] <- "ShippableAmount"



# customer pivot ---- Customer_Pivot ----

reshape2::dcast(model_2, CustomerShipTo + CustomerShipToName + Location ~  . , value.var = "NotShippableAmount", sum) -> Customer_Not_Shippable
reshape2::dcast(model_2, CustomerShipTo + CustomerShipToName + Location ~  . , value.var = "ShippableAmount", sum) -> Customer_Shippable

cbind(Customer_Not_Shippable, Customer_Shippable) -> Customer_Pivot
Customer_Pivot[, -5:-7] -> Customer_Pivot

colnames(Customer_Pivot)[4] <- "NotShippable"
colnames(Customer_Pivot)[5] <- "Shippable"

Customer_Pivot %>% 
  dplyr::mutate(PercentOfPastDate = NotShippable / Shippable) -> Customer_Pivot

# divide with Inf and the others ---- Customer_Pivot_Inf, Customer_Pivot_all ----

Customer_Pivot %>% 
  dplyr::filter(PercentOfPastDate == "Inf") -> Customer_Pivot_Inf

Customer_Pivot %>% 
  dplyr::filter(PercentOfPastDate != "Inf") -> Customer_Pivot_all


Customer_Pivot_Inf
Customer_Pivot_all

# sort by worst Customer actors (Customer_Pivot_all) ---- Customer_Pivot_all_byPercent,  ----

Customer_Pivot_all$PercentOfPastDate -> test1

replace(test1, is.nan(test1) | is.na(test1), 0) -> test1


cbind(Customer_Pivot_all, test1) -> Customer_Pivot_all


Customer_Pivot_all$PercentOfPastDate[is.na(Customer_Pivot_all$PercentOfPastDate)] <- 0

Customer_Pivot_all[, -(ncol(Customer_Pivot_all)-1)] -> Customer_Pivot_all
colnames(Customer_Pivot_all)[ncol(Customer_Pivot_all)] <- "PercentOfPastDate"


Customer_Pivot_all %>% 
  dplyr::arrange(desc(PercentOfPastDate), desc(NotShippable)) -> Customer_Pivot_all_byPercent

Customer_Pivot_all %>% 
  dplyr::arrange(desc(NotShippable), desc(PercentOfPastDate)) -> Customer_Pivot_all_NotShippable


# sort by worst Customer actors (Customer_Pivot_Inf) ---- Customer_Pivot_Inf----

Customer_Pivot_Inf %>% 
  dplyr::arrange(desc(NotShippable)) -> Customer_Pivot_Inf


# sort by worst SKU actors (by SKU amount) ---- Sku_analysis_NotShippable ----

reshape2::dcast(model_2, Sku + ref + Location ~ . , value.var = "NotShippableAmount", sum) -> Sku_NotShippable
reshape2::dcast(model_2, Sku + ref + Location ~ . , value.var = "ShippableAmount", sum) -> Sku_Shippable

cbind(Sku_NotShippable, Sku_Shippable) -> Sku_analysis

Sku_analysis[, -5:-7] -> Sku_analysis
colnames(Sku_analysis)[4] <- "NotShippable"
colnames(Sku_analysis)[5] <- "Shippable"

Sku_analysis %>% 
  dplyr::mutate(PercentOfPastDate = NotShippable / Shippable) -> Sku_analysis


Sku_analysis %>% 
  dplyr::arrange(desc(NotShippable)) -> Sku_analysis_NotShippable

# sort by worst SKU actors (by SKU amount - Inf) ---- Sku_analysis_byPercent_Inf ----

Sku_analysis %>% 
  dplyr::arrange(desc(PercentOfPastDate), desc(NotShippable)) -> Sku_analysis_byPercent

Sku_analysis %>% 
  dplyr::filter(PercentOfPastDate == Inf) -> Sku_analysis_byPercent_Inf

Sku_analysis_byPercent_Inf %>% 
  dplyr::arrange(desc(PercentOfPastDate), desc(NotShippable)) -> Sku_analysis_byPercent_Inf


# sort by worst SKU actors (by SKU amount - all) ---- Sku_analysis_byPercent_all ----

Sku_analysis %>% 
  dplyr::filter(PercentOfPastDate != Inf) -> Sku_analysis_byPercent_all

Sku_analysis_byPercent_all %>% 
  dplyr::arrange(desc(PercentOfPastDate), desc(NotShippable)) -> Sku_analysis_byPercent_all


# Location Analysis sort by worst by Location ---- Location_analysis_byNotShippable, Location_analysis_byPercent ----

reshape2::dcast(Customer_Pivot_all_NotShippable, Location ~ . , value.var = "NotShippable", sum) -> Location_analysis_1

colnames(Location_analysis_1)[2] <- "NotShippable"

Location_analysis_1 %>% 
  dplyr::arrange(desc(NotShippable)) -> Location_analysis_1




reshape2::dcast(Customer_Pivot_all_NotShippable, Location ~ . , value.var = "PercentOfPastDate", sum) -> Location_analysis_2

colnames(Location_analysis_2)[2] <- "PercentOfPastDate"

Location_analysis_2 %>% 
  dplyr::arrange(desc(PercentOfPastDate)) -> Location_analysis_2




reshape2::dcast(Customer_Pivot_all_NotShippable, Location ~ . , value.var = "Shippable", sum) -> Location_analysis_3

colnames(Location_analysis_3)[2] <- "Shippable"

Location_analysis_3 %>% 
  dplyr::arrange(desc(Shippable)) -> Location_analysis_3


merge(Location_analysis_1, Location_analysis_2[, c("Location", "PercentOfPastDate")], by = "Location", all.x = TRUE) -> Location_analysis
merge(Location_analysis, Location_analysis_3[, c("Location", "Shippable")], by = "Location", all.x = TRUE) -> Location_analysis


Location_analysis %>% 
  dplyr::arrange(desc(NotShippable)) %>% 
  dplyr::relocate(Shippable, .after = NotShippable) -> Location_analysis_byNotShippable


Location_analysis %>% 
  dplyr::arrange(desc(PercentOfPastDate)) %>% 
  dplyr::relocate(Shippable, .after = NotShippable) -> Location_analysis_byPercent


Location_analysis_byNotShippable
Location_analysis_byPercent



#  Shelf Life Analysis
model_1
model_2
model_3

model_2 %>% 
  dplyr::select(ref, Location, CustomerShipTo, CustomerShipToName, Sku) -> Shippable_1

merge(model_2, model_3[, c("CustomerShipTo", "PercentOfShelfLife")], by = "CustomerShipTo", all.x = TRUE) -> Shippable_2


Shippable_2[is.na(Shippable_2)] <- 0

Shippable_2 %>% 
  dplyr::select(CustomerShipTo, CustomerShipToName, ref, Location, Sku, PercentOfShelfLife) %>% 
  dplyr::filter(PercentOfShelfLife != 0) %>% 
  dplyr::arrange(Sku, desc(PercentOfShelfLife), CustomerShipTo) %>% 
  dplyr::relocate(Sku, CustomerShipTo, CustomerShipToName, Location, PercentOfShelfLife) -> ShelfLife_Analysis




Shippable_2 %>% 
  dplyr::select(CustomerShipTo, CustomerShipToName, ref, Location, Sku, PercentOfShelfLife) %>% 
  dplyr::filter(PercentOfShelfLife == 0) %>% 
  dplyr::arrange(Sku, desc(PercentOfShelfLife), CustomerShipTo) -> ShelfLife_NoInfo





#################################################################  wrap up  ###############################################################



# Customer Analysis
Customer_Pivot_all_byPercent
Customer_Pivot_all_NotShippable
Customer_Pivot_Inf


# Sku Analysis
Sku_analysis_byPercent_all
Sku_analysis_NotShippable
Sku_analysis_byPercent_Inf

# Location Analysis
Location_analysis_byPercent
Location_analysis_byNotShippable

# Shelf Life Analysis
ShelfLife_Analysis
ShelfLife_NoInfo

###########################################################################################################################################
###########################################################################################################################################
#################################################################  export  ################################################################
###########################################################################################################################################
###########################################################################################################################################


openxlsx::createWorkbook("example") -> example
openxlsx::addWorksheet(example, "Cust List by PastDate Ratio")
openxlsx::addWorksheet(example, "Cust List by Not Shippable")
openxlsx::addWorksheet(example, "Customer List <Inf>")
openxlsx::addWorksheet(example, "SKU List by PastDate Ratio")
openxlsx::addWorksheet(example, "SKU List by Not Shippable")
openxlsx::addWorksheet(example, "SKU List <INf>")
openxlsx::addWorksheet(example, "Loc List by PastDate Ratio")
openxlsx::addWorksheet(example, "Loc List by Not Shippable")

openxlsx::writeDataTable(example, "Cust List by PastDate Ratio", Customer_Pivot_all_byPercent)
openxlsx::writeDataTable(example, "Cust List by Not Shippable", Customer_Pivot_all_NotShippable)
openxlsx::writeDataTable(example, "Customer List <Inf>", Customer_Pivot_Inf)
openxlsx::writeDataTable(example, "SKU List by PastDate Ratio", Sku_analysis_byPercent_all)
openxlsx::writeDataTable(example, "SKU List by Not Shippable", Sku_analysis_NotShippable)
openxlsx::writeDataTable(example, "SKU List <INf>", Sku_analysis_byPercent_Inf)
openxlsx::writeDataTable(example, "Loc List by PastDate Ratio", Location_analysis_byPercent)
openxlsx::writeDataTable(example, "Loc List by Not Shippable", Location_analysis_byNotShippable)

openxlsx::saveWorkbook(example, file = "Shippable Analysis.xlsx")

###########################################################################################################################################
###########################################################################################################################################
#################################################################  export  ################################################################
###########################################################################################################################################
###########################################################################################################################################

openxlsx::createWorkbook("example_1") -> example_1
openxlsx::addWorksheet(example_1, "Shelf Life Analysis")
openxlsx::addWorksheet(example_1, "Shelf Life Analysis _No Info")

openxlsx::writeDataTable(example_1, "Shelf Life Analysis", ShelfLife_Analysis)
openxlsx::writeDataTable(example_1, "Shelf Life Analysis _No Info", ShelfLife_NoInfo)

openxlsx::saveWorkbook(example_1, file = "Shelf Life Analysis.xlsx")



