library(tidyverse)
library(readxl)
library(writexl)
library(openxlsx)
library(reshape2)


Shelf_Life_Analysis <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shelf Life Analysis.xlsx", 
                                  sheet = "Shelf Life Analysis")


colnames(Shelf_Life_Analysis)[1] <- "Sku"
colnames(Shelf_Life_Analysis)[2] <- "Super_Customer_No"
colnames(Shelf_Life_Analysis)[3] <- "Super_Customer_Name"
colnames(Shelf_Life_Analysis)[4] <- "Customer_Ship_To_No"
colnames(Shelf_Life_Analysis)[5] <- "Customer_Ship_To_Name"
colnames(Shelf_Life_Analysis)[6] <- "Location"
colnames(Shelf_Life_Analysis)[7] <- "Percent_Of_Shelf_Life"
colnames(Shelf_Life_Analysis)[8] <- "ref"
colnames(Shelf_Life_Analysis)[9] <- "Product_Sku_Name"
colnames(Shelf_Life_Analysis)[10] <- "Label"
colnames(Shelf_Life_Analysis)[11] <- "Product_Category_Code"
colnames(Shelf_Life_Analysis)[12] <- "Product_Cateogry_Name"
colnames(Shelf_Life_Analysis)[13] <- "Product_Sub_Category_Code"
colnames(Shelf_Life_Analysis)[14] <- "Product_Sub_Category_Name"
colnames(Shelf_Life_Analysis)[15] <- "Product_Platform_Code"
colnames(Shelf_Life_Analysis)[16] <- "Product_Platform_Name"
colnames(Shelf_Life_Analysis)[17] <- "Sales_Channel_No"
colnames(Shelf_Life_Analysis)[18] <- "Sales_Channel_Desc"
colnames(Shelf_Life_Analysis)[19] <- "Sales_Manager_No"
colnames(Shelf_Life_Analysis)[20] <- "Sales_Manager_Name"
colnames(Shelf_Life_Analysis)[21] <- "Net_Pounds(lbs.)"

Shelf_Life_Analysis %>% 
  dplyr::relocate(Sku, Product_Sku_Name) -> Shelf_Life_Analysis


# If only one customer falls into SKU -> eliminate

data.frame(table(Shelf_Life_Analysis$Sku)) -> counta

colnames(counta)[1] <- "Sku"
colnames(counta)[2] <- "counta"

merge(Shelf_Life_Analysis, counta[, c("Sku", "counta")], by = "Sku", all.x = TRUE) -> Shelf_Life_Analysis

Shelf_Life_Analysis %>% 
  dplyr::relocate(counta) %>% 
  dplyr::filter(counta != 1) -> Shelf_Life_Analysis




# If every customer has the same shelf life -> eliminate

Shelf_Life_Analysis %>% 
  dplyr::mutate(num_Percent_Of_Shelf_Life = Percent_Of_Shelf_Life) %>% 
  dplyr::mutate(num_Percent_Of_Shelf_Life = as.numeric(num_Percent_Of_Shelf_Life)) -> Shelf_Life_Analysis

reshape2::dcast(Shelf_Life_Analysis, Sku ~  Percent_Of_Shelf_Life , value.var = "num_Percent_Of_Shelf_Life", length) -> counta_2





# Here converted to excel, and imported again later
openxlsx::createWorkbook("example_1") -> example_1
openxlsx::addWorksheet(example_1, "Shelf Life Analysis")
openxlsx::addWorksheet(example_1, "counta_2")

openxlsx::writeDataTable(example_1, "Shelf Life Analysis", Shelf_Life_Analysis)
openxlsx::writeDataTable(example_1, "counta_2", counta_2)


openxlsx::saveWorkbook(example_1, file = "work_2.xlsx")

