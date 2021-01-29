
library(shiny)
library(readxl)
library(dplyr)
library(janitor)
library(XLConnect)
library(xlsx)


ui <- fluidPage(

    titlePanel("Auto-Mailout"),

    sidebarLayout(
        sidebarPanel(
            fileInput("file1", "Upload RDB Extract", accept = c("xlsx")),
            tags$hr(),
            downloadButton("download", "Download Mailout")
        ),

        # Show a plot of the generated distribution
        mainPanel(
           tableOutput("extract")
        )
    )
)


server <- function(input, output) {
      
        getData <- reactive({
            
            inFile <- input$file1
            
            if(is.null(input$file1))
                return(NULL)
            
            read_xlsx(inFile$datapath) %>%
                clean_names()
        })
        
        
        output$extract <- renderTable(
            getData()
        )
        
        output$download <- downloadHandler(
            filename = function(){
                paste0("rdbtest.xlsx")
            },
            content = function(file){
                write.xlsx(getData(), file)
            }
        )

}


shinyApp(ui = ui, server = server)
