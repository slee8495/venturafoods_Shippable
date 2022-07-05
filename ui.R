library(shiny)
library(shinydashboard)


shinyUI(fluidPage(
  headerPanel(title = "Shelf Life Analysis"),
  sidebarLayout(
    sidebarPanel(
      selectInput(inputId = "Sku", label = "Select the Sku", choices = c("all", ShelfLife_Analysis$Sku))
      
          
      ),
    
    mainPanel(
      dataTableOutput("ShelfLife_Analysis"),
      downloadButton("downloadData", "Download Data", style = "display: black; margin: 0 auto; width: 230px;color: black;")
      
    )
  )
))



