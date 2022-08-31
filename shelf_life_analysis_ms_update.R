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
attributes <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/8.31.22/attributes.xlsx")

attributes %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(ref = paste0(location_no, "_", product_label_sku_code)) %>% 
  dplyr::mutate(ref = gsub("-", "", ref)) -> attributes


# Read model_2 file from MS ----
model_2 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/8.31.22/model_2.xlsx")

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
model_3 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Shipabble/Shippable Tool MS Dossier Automation/8.31.22/model_3.xlsx")

model_3[-1, ] -> model_3
colnames(model_3) <- model_3[1, ]
model_3[-1, ] -> model_3

model_3 %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  data.frame() %>% 
  dplyr::mutate(ref_2 = paste0(location, "_", ship_to_customer)) -> model_3


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
  dplyr::select(ref_2, percent_of_shelf_life) -> model_3_perct_shelf_life

attributes %>% 
  dplyr::left_join(model_3_perct_shelf_life, by = "ref_2") %>% 
  dplyr::mutate(percent_of_shelf_life = replace(percent_of_shelf_life, is.na(percent_of_shelf_life), 50)) -> attributes


attributes %>% 
  dplyr::select(customer_ship_to_name_1, customer_ship_to_ship_to, label_code, location_no, percent_of_shelf_life, product_category_code,
                product_category_name, product_platform_code, product_platform_name, product_label_sku_name, product_sub_category_code,
                product_sub_category_name, sales_channel_desc, sales_channel_no, sales_manager_name, sales_manager_no, product_label_sku_code,
                super_customer_name, super_customer_no, net_pounds_lbs, selling_region_no, selling_region_name, ref) %>% 
  dplyr::rename("Customer Ship To Name" = customer_ship_to_name_1,
                "Customer Ship To No" = customer_ship_to_ship_to,
                Label = label_code,
                Location = location_no,
                "Percent Of Shelf Life" = percent_of_shelf_life,
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
writexl::write_xlsx(attributes, "Shelf Life Analysis.xlsx")
