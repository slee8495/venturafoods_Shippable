library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)



# Read attributes file from MS ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/1ADEF816D3436E56B40245AE19C3CF5B/K53--K46
## Make sure to go with past 6 months (Edit dataset -> Edit Filter)
attributes <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/2023/11.01.2023/attributes.xlsx")

attributes %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(ref = paste0(location_no, "_", product_label_sku_code)) %>% 
  dplyr::mutate(ref = gsub("-", "", ref)) -> attributes


# Read model_2 file from MS ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/899AA319964180E4D2C3E2AF594BB2F5
model_2 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/2023/11.01.2023/model_2.xlsx")

model_2[-1, ] -> model_2
colnames(model_2) <- model_2[1, ]
model_2[-1, ] -> model_2

model_2 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(product_label_sku_code = gsub("-", "", product_label_sku_code)) %>% 
  dplyr::mutate(ref = paste0(location_no, "_", product_label_sku_code)) -> model_2


# Read model_3 file from MS ----
# https://edgeanalytics.venturafoods.com/MicroStrategyLibrary/app/DF007F1C11E9B3099BB30080EF7513D2/7A11AE535B42A4DBD97073A3856B968F/K53--K46

# temporary using AS400 file from Dominee
model_3 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/2023/11.01.2023/model_3.xlsx")

model_3[-1, ] -> model_3
colnames(model_3) <- model_3[1, ]

model_3 %>% 
  dplyr::slice(-1) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(customer_number = customer_ship_to,
                shelf_life_percentage = product_ship_shelf_life_percent) %>% 
  dplyr::select(customer_number, location, shelf_life_percentage) %>% 
  dplyr::rename(customer_ship_to = customer_number,
                product_ship_shelf_life_percent = shelf_life_percentage) %>% 
  dplyr::mutate(ref_2 = paste0(location, "_", customer_ship_to)) -> model_3


# Ship to customer into attributes
model_2 %>% 
  dplyr::select(ref, customer_ship_to_ship_to, customer_ship_to_name_1) -> model_2_cust_ship_to

attributes %>% 
  dplyr::left_join(model_2_cust_ship_to, by = "ref") -> attributes


# attributes ref_2 
attributes %>% 
  dplyr::mutate(ref_2 = paste0(location_no, "_", customer_ship_to_ship_to)) -> attributes


# percent of shelf life to attributes
model_3 %>% 
  dplyr::select(ref_2, product_ship_shelf_life_percent) -> model_3_perct_shelf_life

attributes %>% 
  dplyr::left_join(model_3_perct_shelf_life, by = "ref_2") %>% 
  dplyr::mutate(product_ship_shelf_life_percent = replace(product_ship_shelf_life_percent, is.na(product_ship_shelf_life_percent), 50)) -> attributes


attributes %>% 
  dplyr::select(customer_ship_to_name_1, customer_ship_to_ship_to, label_code, location_no, product_ship_shelf_life_percent, product_category_code,
                product_category_name, product_platform_code, product_platform_name, product_label_sku_name, product_sub_category_code,
                product_sub_category_name, sales_channel_desc, sales_channel_no, sales_manager_name, sales_manager_no, product_label_sku_code,
                super_customer_name, super_customer_no, net_pounds_lbs, selling_region_no, selling_region_name, ref) %>% 
  dplyr::rename("Customer Ship To Name" = customer_ship_to_name_1,
                "Customer Ship To No" = customer_ship_to_ship_to,
                Label = label_code,
                Location = location_no,
                "Percent Of Shelf Life" = product_ship_shelf_life_percent,
                "Product Category Code" = product_category_code,
                "Product Category Name" = product_category_name,
                "Product Platform Code" = product_platform_code,
                "Product Platform Name" = product_platform_name,
                "Product Sku Name" = product_label_sku_name,
                "Product Sub Category Code" = product_sub_category_code,
                "Product Sub Category Name" = product_sub_category_name,
                "Sales Channel Desc" = sales_channel_desc,
                "Sales Channel No" = sales_channel_no,
                "Sales Manager Name" = sales_manager_name,
                "Sales Manager No" = sales_manager_no,
                Sku = product_label_sku_code,
                "Super Customer Name" = super_customer_name,
                "Super Customer No" = super_customer_no,
                "Net Pounds(lbs.)" = net_pounds_lbs,
                "Selling Region No" = selling_region_no,
                "Selling Region Name" = selling_region_name) -> attributes


### Export to Excel 
writexl::write_xlsx(attributes, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/2023/11.01.2023/Shelf Life Analysis.xlsx")
