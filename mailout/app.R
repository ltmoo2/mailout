
library(shiny)
library(readxl)


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
      
        output$extract <- renderTable({
            
            req(input$file1)
            
            df <- read_xlsx(input$file1$datapath)
            
            return(df)
        })
}


shinyApp(ui = ui, server = server)
