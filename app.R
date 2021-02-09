library(shiny)
library(readxl)
library(dplyr)
library(janitor)
library(XLConnect)
library(tidyr)
library(DT)


wb <- loadWorkbook("List Template.xlsx")
setStyleAction(wb, XLC$"STYLE_ACTION.NONE")




ui <- fluidPage(

    titlePanel(title = div (img(src = "logo.png", width = 200, height = 80),
               "Auto-Mailout")),

    sidebarLayout(
        sidebarPanel(
            fileInput("file1", "Upload RDB Extract", accept = c("xlsx")),
            "Please wait for table to appear to the right before downloading files",
            tags$hr(),
            downloadButton("download", "Download Mailout")
        ),

        # Show a plot of the generated distribution
        mainPanel(
           DTOutput("extract")
        )
    )
)


server <- function(input, output) {
  
        maildate <- reactive({
          
          req(input$file1)
          
          as.character(Sys.Date())
          
        })
        
        observe(dir.create(paste0(maildate(), "-mailout")))
    
      
        getData <- reactive({
            
            inFile <- input$file1
            
            if(is.null(input$file1))
                return(NULL)
            
            data <- read_xlsx(inFile$datapath) %>%
                clean_names() %>%
                filter(!is.na(party_managing_agent_name)) %>%
                filter(!grepl("/", party_managing_agent_name))
            
            return(data)
        })
        
        listData <- reactive({
          inFile <- input$file1
          
          if(is.null(input$file1))
            return(NULL)
          
          list_data <- getData() %>%
            select(municipality, party_managing_agent, party_managing_agent_name) %>%
            group_by(party_managing_agent,party_managing_agent_name) %>%
            summarise(Melbourne = as.numeric(sum(municipality == "Melbourne")),
                      Yarra = as.numeric(sum(municipality == "Yarra")),
                      Darebin = as.numeric(sum(municipality == "Darebin")),
                      Maribyrnong = as.numeric(sum(municipality == "Maribyrnong")),
                      Knox = as.numeric(sum (municipality == "Knox")),
                      Monash = as.numeric(sum (municipality == "Monash"))) %>%
            separate(party_managing_agent_name, c("first", "last"), remove = FALSE, sep = "\\s") %>%
            mutate(Total = Melbourne + Yarra + Darebin + Maribyrnong + Knox + Monash) %>%
            select(party_managing_agent:last, Total, Melbourne:Monash)
          
          return(list_data)
        })
        
        output$extract <- renderDT(
            listData(),
            options = list(pageLength = 25, info = FALSE,
                           lengthMenu = list(c(25, -1), c("25", "All")))
        )
        
        observeEvent(input$file1, writeWorksheet(wb, listData(), "Sheet1", startRow = 2, startCol = 1, header = FALSE))
        observeEvent(input$file1, saveWorkbook(wb, paste0(maildate(), "-mailout/", maildate(),"-Mailing List.xlsx")))
        
        observeEvent(input$file1, 
                     for(i in (unique(getData()$party_managing_agent_name))){
                       data <- subset(getData(), party_managing_agent_name == i)
                       
                       municipality <- data %>%
                         select(municipality)
                       
                       adresses <- data %>%
                         select(address)
                       
                       agency <- data %>%
                         distinct(party_managing_agent) %>%
                         top_n(1)
                       
                       agent <- data %>%
                         distinct(party_managing_agent_name)
                       
                       wb1 <- loadWorkbook("Information Template V2.xlsx", create = TRUE)
                       setStyleAction(wb1, XLC$"STYLE_ACTION.NONE")
                       
                       
                       writeWorksheet(wb1, municipality, "City Council Questionnaire", startRow = 5, startCol = 2, header = FALSE)
                       
                       writeWorksheet(wb1, adresses, "City Council Questionnaire", startRow = 5, startCol = 3, header = FALSE)
                       
                       writeWorksheet(wb1, agency, "City Council Questionnaire", startRow = 2, startCol = 2, header = FALSE)
                       
                       writeWorksheet(wb1, agent, "City Council Questionnaire", startRow = 3, startCol = 2, header = FALSE)
                       
                       saveWorkbook(wb1, paste0(maildate(), "-mailout/",  unique(data$party_managing_agent_name), ".xlsx"))
                     })
        
        
        output$download <- downloadHandler(
          filename <- function() {
            paste("output", "zip", sep = ".")
          },
          
          content <- function(file){
            zip(file, files = paste0(maildate(), "-mailout"))
          }
        )


}


shinyApp(ui = ui, server = server)
