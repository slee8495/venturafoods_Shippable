library(shiny)
library(DT)


shinyServer(function(input, output, session){
  output$ShelfLife_Analysis <- renderDataTable({
    names <- NULL
    if (input$Sku == "all"){
      names <- ShelfLife_Analysis$Sku
    } else {
      names <- input$Sku
    }
    filter(ShelfLife_Analysis, Sku %in% names)
  })

})
