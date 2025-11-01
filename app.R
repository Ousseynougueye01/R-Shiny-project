library(shiny)
library(dplyr)
library(bslib)
library(openxlsx)
library(readxl)
library(leaflet)

# Fonction pour déterminer la feuille à utiliser selon l'effectif
determiner_feuille <- function(effectif) {
  if (effectif >= 0 && effectif <= 25) {
    return("Feuil2")
  } else if (effectif >= 26 && effectif <= 50) {
    return("Feuil3")
  } else if (effectif > 50) {
    return("Feuil1")
  } else {
    return("Feuil1")  # Par défaut
  }
}

# Fonction pour charger le barème selon l'effectif
charger_bareme <- function(effectif) {
  feuille <- determiner_feuille(effectif)
  tryCatch({
    read_excel("Tarif.xlsx", sheet = feuille)
  }, error = function(e) {
    # En cas d'erreur, créer un data frame vide avec les colonnes attendues
    data.frame(
      TERRITORIALITE = character(0),
      TAUX = character(0),
      Prime_enfant = numeric(0),
      Prime_adulte = numeric(0),
      Prime_senior_61_65 = numeric(0),
      Prime_senior_66_70 = numeric(0),
      Prime_senior_71_80 = numeric(0)
    )
  })
}

# Fonction pour calculer la réduction selon l'effectif
calculer_reduction <- function(effectif_total) {
  if (effectif_total <= 100) return(0)
  else if (effectif_total <= 200) return(7)
  else if (effectif_total <= 300) return(10)
  else if (effectif_total <= 400) return(12)
  else if (effectif_total <= 500) return(15)
  else if (effectif_total <= 750) return(18)
  else if (effectif_total <= 1000) return(20)
  else if (effectif_total <= 2500) return(25)
  else return(30)
}

ui <- fluidPage(
  theme = bs_theme(
    bootswatch = "flatly",
    primary = "#3498db",
    secondary = "#2c3e50",
    success = "#27ae60"
  ),
  
  tags$head(
    tags$style(HTML("
      .shiny-output-error { visibility: hidden; }
      .shiny-output-error:before { visibility: hidden; }
      
      body {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Arial', sans-serif;
        min-height: 100vh;
      }
      
      .main-container {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        margin: 20px;
        padding: 30px;
      }
      
      .card {
        background: white;
        border-radius: 15px;
        box-shadow: 0 8px 25px rgba(0,0,0,0.08);
        margin-bottom: 20px;
        padding: 20px;
      }
      
      .card-header {
        background: linear-gradient(135deg, #3498db, #2980b9);
        color: white;
        border-radius: 10px;
        padding: 15px 20px;
        margin: -20px -20px 20px -20px;
        font-weight: 600;
      }
      
      .form-control {
        border-radius: 10px;
        border: 2px solid #e9ecef;
        padding: 12px 15px;
        margin-bottom: 15px;
      }
      
      .form-control:focus {
        border-color: #3498db;
        box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
      }
      
      .btn {
        border-radius: 10px;
        padding: 12px 25px;
        font-weight: 600;
        border: none;
      }
      
      .btn-primary {
        background: linear-gradient(135deg, #3498db, #2980b9);
      }
      
      .btn-success {
        background: linear-gradient(135deg, #27ae60, #229954);
      }
      
      .prime-display {
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        border-radius: 15px;
        padding: 20px;
        margin: 15px 0;
        border-left: 5px solid #3498db;
      }
      
      .prime-nette {
        font-size: 18px;
        font-weight: 600;
        color: #2980b9;
        margin-bottom: 8px;
      }
      
      .taxe-display {
        font-size: 16px;
        font-weight: 500;
        color: #8e44ad;
        margin-bottom: 8px;
      }
      
      .commission-display {
        font-size: 16px;
        font-weight: 500;
        color: #d35400;
        margin-bottom: 8px;
      }
      
      .prime-ttc {
        font-size: 24px;
        font-weight: 800;
        color: #27ae60;
        margin-top: 10px;
        padding-top: 10px;
        border-top: 2px solid #ecf0f1;
      }
      
      .effectif-total {
        font-size: 16px;
        font-weight: 600;
        color: #34495e;
        margin-bottom: 10px;
      }
      
      .feuille-info {
        font-size: 14px;
        font-weight: 500;
        color: #e67e22;
        margin-bottom: 10px;
        font-style: italic;
      }
      
      .header-title {
        color: #2c3e50;
        font-weight: 700;
        margin-bottom: 30px;
      }
      
      .stats-row {
        display: flex;
        justify-content: space-around;
        margin: 20px 0;
      }
      
      .stats-item {
        text-align: center;
        background: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        flex: 1;
        margin: 0 10px;
      }
      
      .stats-number {
        font-size: 24px;
        font-weight: 800;
        color: #3498db;
        display: block;
      }
      
      .stats-label {
        font-size: 12px;
        color: #7f8c8d;
        text-transform: uppercase;
        font-weight: 600;
      }
      
      table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
      }
      
      th, td {
        padding: 12px;
        text-align: left;
        border-bottom: 1px solid #ddd;
      }
      
      th {
        background-color: #f8f9fa;
        font-weight: 600;
        color: #2c3e50;
      }
      
      tr:hover {
        background-color: #f5f5f5;
      }
    ")),
    tags$title("Plateforme de Tarification - Assurance Santé Collective")
  ),
  tags$head(tags$link(rel = "shortcut icon", href = "favicon.png")),
  
  div(class = "main-container",
      # En-tête
      fluidRow(
        column(2, 
               tags$img(src = "logos.png", height = "80px", style = "max-width: 100%;")
        ),
        column(10,
               h1("Plateforme de Tarification", class = "header-title", style = "margin-top: 15px; text-align: center;"),
               h4("Assurance Santé Collective", style = "color: #7f8c8d; font-weight: 400; margin-top: -10px; text-align: center;")
        )
      ),
      
      hr(style = "border-top: 2px solid #3498db; margin: 30px 0;"),
      
      fluidRow(
        # Panneau latéral
        column(4,
               div(class = "card",
                   div(class = "card-header",
                       h4("Configuration du Contrat", style = "margin: 0;")
                   ),
                   textInput("prospect", "Nom du prospect :"),
                   
                   selectInput("territoire", "Territorialité :",
                               choices = NULL,
                               selected = NULL),
                   
                   selectInput("taux", "Taux de prise en charge :",
                               choices = NULL,
                               selected = NULL),
                   
                   h5("Composition des Assurés", style = "color: #2c3e50; margin: 20px 0 15px 0;"),
                   
                   numericInput("nb_enfants", "Enfants", 
                                value = 0, min = 0, step = 1),
                   
                   numericInput("nb_adultes", "Adultes", 
                                value = 0, min = 0, step = 1),
                   
                   numericInput("nb_seniors_61_65", "Adultes (61-65 ans)", 
                                value = 0, min = 0, step = 1),
                   
                   numericInput("nb_seniors_66_70", "Adultes (66-70 ans)", 
                                value = 0, min = 0, step = 1),
                   
                   numericInput("nb_seniors_71_80", "Adultes (71-80 ans)", 
                                value = 0, min = 0, step = 1)
               )
        ),
        
        # Panneau principal
        column(8,
               # Résultats
               div(class = "card",
                   div(class = "card-header",
                       h4("Résultats de Tarification", style = "margin: 0;")
                   ),
                   
                   div(class = "effectif-total",
                       textOutput("effectif_total")
                   ),
                   
                   div(class = "prime-display",
                       div(class = "prime-nette",
                           textOutput("prime_nette")
                       ),
                       div(class = "taxe-display",
                           textOutput("taxe_display")
                       ),
                       div(class = "commission-display",
                           textOutput("commission_display")
                       ),
                       div(class = "prime-ttc",
                           textOutput("prime_ttc")
                       )
                   ),
                   
                   downloadButton("download_devis", "Télécharger le Devis",
                                  class = "btn btn-success",
                                  style = "width: 100%; font-size: 16px;")
               ),
               
               div(class = "card",
                   div(class = "card-header",
                       h4("Zone Géographique", style = "margin: 0;")
                   ),
                   leafletOutput("map", height = 500)
               )
        )
      )
  )
)

server <- function(input, output, session) {
  
  # Effectif total réactif avec gestion des valeurs nulles
  effectif_total <- reactive({
    enfants <- if(is.null(input$nb_enfants) || is.na(input$nb_enfants)) 0 else input$nb_enfants
    adultes <- if(is.null(input$nb_adultes) || is.na(input$nb_adultes)) 0 else input$nb_adultes
    seniors_61_65 <- if(is.null(input$nb_seniors_61_65) || is.na(input$nb_seniors_61_65)) 0 else input$nb_seniors_61_65
    seniors_66_70 <- if(is.null(input$nb_seniors_66_70) || is.na(input$nb_seniors_66_70)) 0 else input$nb_seniors_66_70
    seniors_71_80 <- if(is.null(input$nb_seniors_71_80) || is.na(input$nb_seniors_71_80)) 0 else input$nb_seniors_71_80
    
    enfants + adultes + seniors_61_65 + seniors_66_70 + seniors_71_80
  })
  
  # Barème réactif basé sur l'effectif
  bareme_actuel <- reactive({
    effectif <- effectif_total()
    charger_bareme(effectif)
  })
  
  # Observer pour mettre à jour les choix des selectInput
  observe({
    bareme <- bareme_actuel()
    
    if (nrow(bareme) > 0) {
      # Mettre à jour les choix de territorialité
      territoires <- unique(bareme$TERRITORIALITE)
      current_territoire <- input$territoire
      
      if (!is.null(current_territoire) && current_territoire %in% territoires) {
        selected_territoire <- current_territoire
      } else {
        selected_territoire <- territoires[1]
      }
      
      updateSelectInput(session, "territoire",
                        choices = territoires,
                        selected = selected_territoire)
      
      # Mettre à jour les choix de taux
      taux_disponibles <- unique(bareme$TAUX)
      current_taux <- input$taux
      
      if (!is.null(current_taux) && current_taux %in% taux_disponibles) {
        selected_taux <- current_taux
      } else {
        selected_taux <- taux_disponibles[1]
      }
      
      updateSelectInput(session, "taux",
                        choices = taux_disponibles,
                        selected = selected_taux)
    }
  })
  
  # Données réactives
  selected_tarif <- reactive({
    req(input$territoire, input$taux)
    bareme <- bareme_actuel()
    filter(bareme, TERRITORIALITE == input$territoire, TAUX == input$taux)
  })
  
  reduction_taux <- reactive({
    calculer_reduction(effectif_total())
  })
  
  prime_brute <- reactive({
    tarif <- selected_tarif()
    if (nrow(tarif) == 0) return(0)
    
    # Gestion sécurisée des inputs avec valeurs par défaut
    nb_enfants <- if(is.null(input$nb_enfants) || is.na(input$nb_enfants)) 0 else input$nb_enfants
    nb_adultes <- if(is.null(input$nb_adultes) || is.na(input$nb_adultes)) 0 else input$nb_adultes
    nb_seniors_61_65 <- if(is.null(input$nb_seniors_61_65) || is.na(input$nb_seniors_61_65)) 0 else input$nb_seniors_61_65
    nb_seniors_66_70 <- if(is.null(input$nb_seniors_66_70) || is.na(input$nb_seniors_66_70)) 0 else input$nb_seniors_66_70
    nb_seniors_71_80 <- if(is.null(input$nb_seniors_71_80) || is.na(input$nb_seniors_71_80)) 0 else input$nb_seniors_71_80
    
    total <- 0
    if ("Prime_enfant" %in% names(tarif)) total <- total + tarif$Prime_enfant * nb_enfants
    if ("Prime_adulte" %in% names(tarif)) total <- total + tarif$Prime_adulte * nb_adultes
    if ("Prime_senior_61_65" %in% names(tarif)) total <- total + tarif$Prime_senior_61_65 * nb_seniors_61_65
    if ("Prime_senior_66_70" %in% names(tarif)) total <- total + tarif$Prime_senior_66_70 * nb_seniors_66_70
    if ("Prime_senior_71_80" %in% names(tarif)) total <- total + tarif$Prime_senior_71_80 * nb_seniors_71_80
    
    return(total)
  })
  
  prime_nette <- reactive({
    brute <- prime_brute()
    if (brute == 0) return(0)
    
    reduction <- reduction_taux()
    brute * (1 - reduction / 100)
  })
  
  # Calcul des commissions avec plafond
  commissions <- reactive({
    total_personnes <- effectif_total()
    if (total_personnes == 0) return(0)
    
    commission_brute <- total_personnes * 1500
    min(commission_brute, 300000)  # Plafond de 300 000 XOF
  })
  
  # Calcul des taxes (5% fixe)
  taxes <- reactive({
    nette <- prime_nette()
    if (nette == 0) return(0)
    
    (nette + commissions())*0.05
  })
  
  # Outputs
  output$effectif_total <- renderText({
    total <- effectif_total()
    paste("Effectif total:", format(total, big.mark = " "), "personnes")
  })
  
  output$prime_nette <- renderText({
    nette <- prime_nette()
    paste("Prime nette :", format(round(nette), big.mark = " ", scientific = FALSE), "XOF")
  })
  
  output$taxe_display <- renderText({
    taxe <- taxes()
    paste("Taxes :", format(round(taxe), big.mark = " ", scientific = FALSE), "XOF")
  })
  
  output$commission_display <- renderText({
    commission <- commissions()
    paste("Frais et accessoires :", format(round(commission), big.mark = " ", scientific = FALSE), "XOF")
  })
  
  output$prime_ttc <- renderText({
    nette <- prime_nette()
    taxe <- taxes()
    commission <- commissions()
    
    if (nette == 0) return("Prime TTC : 0 XOF")
    
    ttc <- nette + taxe + commission
    paste("Prime TTC :", format(round(ttc), big.mark = " ", scientific = FALSE), "XOF")
  })
  
  # Téléchargement du devis
  output$download_devis <- downloadHandler(
    filename = function() {
      prospect_name <- if(is.null(input$prospect) || input$prospect == "") "devis" else gsub("[^A-Za-z0-9]", "_", input$prospect)
      paste0("devis_", prospect_name, "_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      tarif <- selected_tarif()
      brute <- prime_brute()
      nette <- prime_nette()
      taxe <- taxes()
      commission <- commissions()
      ttc <- nette + taxe + commission
      total_personnes <- effectif_total()
      
      # Gestion sécurisée des inputs
      nb_enfants <- if(is.null(input$nb_enfants) || is.na(input$nb_enfants)) 0 else input$nb_enfants
      nb_adultes <- if(is.null(input$nb_adultes) || is.na(input$nb_adultes)) 0 else input$nb_adultes
      nb_seniors_61_65 <- if(is.null(input$nb_seniors_61_65) || is.na(input$nb_seniors_61_65)) 0 else input$nb_seniors_61_65
      nb_seniors_66_70 <- if(is.null(input$nb_seniors_66_70) || is.na(input$nb_seniors_66_70)) 0 else input$nb_seniors_66_70
      nb_seniors_71_80 <- if(is.null(input$nb_seniors_71_80) || is.na(input$nb_seniors_71_80)) 0 else input$nb_seniors_71_80
      
      # Récupération des primes unitaires
      prime_adulte <- if(nrow(tarif) > 0 && "Prime_adulte" %in% names(tarif)) tarif$Prime_adulte else 0
      prime_enfant <- if(nrow(tarif) > 0 && "Prime_enfant" %in% names(tarif)) tarif$Prime_enfant else 0
      prime_senior_61_65 <- if(nrow(tarif) > 0 && "Prime_senior_61_65" %in% names(tarif)) tarif$Prime_senior_61_65 else 0
      prime_senior_66_70 <- if(nrow(tarif) > 0 && "Prime_senior_66_70" %in% names(tarif)) tarif$Prime_senior_66_70 else 0
      prime_senior_71_80 <- if(nrow(tarif) > 0 && "Prime_senior_71_80" %in% names(tarif)) tarif$Prime_senior_71_80 else 0
      
      # Nom du prospect
      nom_prospect <- if(is.null(input$prospect) || input$prospect == "") "Prospect" else input$prospect
      
      # Création du workbook
      wb <- createWorkbook()
      addWorksheet(wb, "Devis")
      
      # Styles de formatage
      title_style <- createStyle(
        fontSize = 16,
        fontName = "Arial",
        fontColour = "#FFFFFF",
        fgFill = "#3498db",
        halign = "center",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#2980b9",
        borderStyle = "thick"
      )
      
      info_label_style <- createStyle(
        fontSize = 11,
        fontName = "Arial",
        fontColour = "#2c3e50",
        fgFill = "#ecf0f1",
        halign = "left",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#bdc3c7"
      )
      
      info_value_style <- createStyle(
        fontSize = 11,
        fontName = "Arial",
        halign = "left",
        valign = "center",
        border = "TopBottomLeftRight",
        borderColour = "#bdc3c7"
      )
      
      column_header_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        fontColour = "#FFFFFF",
        fgFill = "#3498db",
        halign = "center",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#2980b9"
      )
      
      data_style <- createStyle(
        fontSize = 11,
        fontName = "Arial",
        halign = "center",
        valign = "center",
        border = "TopBottomLeftRight",
        borderColour = "#bdc3c7"
      )
      
      total_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        fontColour = "#2c3e50",
        fgFill = "#f8f9fa",
        halign = "center",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#95a5a6",
        borderStyle = "medium"
      )
      
      recap_header_style <- createStyle(
        fontSize = 14,
        fontName = "Arial",
        fontColour = "#FFFFFF",
        fgFill = "#3498db",
        halign = "center",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#2980b9",
        borderStyle = "thick"
      )
      
      recap_label_style <- createStyle(
        fontSize = 11,
        fontName = "Arial",
        fontColour = "#2c3e50",
        fgFill = "#f8f9fa",
        halign = "left",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#bdc3c7"
      )
      
      recap_value_style <- createStyle(
        fontSize = 11,
        fontName = "Arial",
        halign = "right",
        valign = "center",
        border = "TopBottomLeftRight",
        borderColour = "#bdc3c7"
      )
      
      ttc_label_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        fontColour = "#FFFFFF",
        fgFill = "#27ae60",
        halign = "left",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#229954",
        borderStyle = "thick"
      )
      
      ttc_value_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        fontColour = "#FFFFFF",
        fgFill = "#27ae60",
        halign = "right",
        valign = "center",
        textDecoration = "bold",
        border = "TopBottomLeftRight",
        borderColour = "#229954",
        borderStyle = "thick"
      )
      
      # Logo (positionné à côté du tableau principal, colonnes F-H)
      tryCatch({
        if (file.exists("www/logos.png")) {
          insertImage(wb, "Devis", "www/logos.png", 
                      startCol = 6, startRow = 3, 
                      width = 3, height = 2.5)
        }
      }, error = function(e) {
        # Continue sans erreur si pas de logo
      })
      
      # Titre principal
      writeData(wb, "Devis", "DEVIS D'ASSURANCE SANTÉ COLLECTIVE", startCol = 1, startRow = 1)
      mergeCells(wb, "Devis", cols = 1:4, rows = 1)
      addStyle(wb, "Devis", title_style, rows = 1, cols = 1:4)
      
      # Informations du contrat
      current_row <- 3
      
      writeData(wb, "Devis", "Prospect:", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", nom_prospect, startCol = 2, startRow = current_row)
      writeData(wb, "Devis", "Date:", startCol = 3, startRow = current_row)
      writeData(wb, "Devis", format(Sys.Date(), "%d-%m-%Y"), startCol = 4, startRow = current_row)
      addStyle(wb, "Devis", info_label_style, rows = current_row, cols = c(1,3))
      addStyle(wb, "Devis", info_value_style, rows = current_row, cols = c(2,4))
      current_row <- current_row + 1
      
      writeData(wb, "Devis", "Territorialité:", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", input$territoire, startCol = 2, startRow = current_row)
      writeData(wb, "Devis", "Taux de prise en charge:", startCol = 3, startRow = current_row)
      writeData(wb, "Devis", paste0(input$taux), startCol = 4, startRow = current_row)
      addStyle(wb, "Devis", info_label_style, rows = current_row, cols = c(1,3))
      addStyle(wb, "Devis", info_value_style, rows = current_row, cols = c(2,4))
      current_row <- current_row + 2
      
      # Tableau principal - En-têtes
      writeData(wb, "Devis", "Catégorie", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", "Nombre", startCol = 2, startRow = current_row)
      writeData(wb, "Devis", "Prime / personne", startCol = 3, startRow = current_row)
      writeData(wb, "Devis", "Prime nette", startCol = 4, startRow = current_row)
      addStyle(wb, "Devis", column_header_style, rows = current_row, cols = 1:4)
      current_row <- current_row + 1
      
      # Données pour chaque catégorie
      if (nb_adultes > 0) {
        writeData(wb, "Devis", "Adulte", startCol = 1, startRow = current_row)
        writeData(wb, "Devis", nb_adultes, startCol = 2, startRow = current_row)
        writeData(wb, "Devis", format(prime_adulte, big.mark = " "), startCol = 3, startRow = current_row)
        writeData(wb, "Devis", format(round(prime_adulte * nb_adultes), big.mark = " "), startCol = 4, startRow = current_row)
        addStyle(wb, "Devis", data_style, rows = current_row, cols = 1:4)
        current_row <- current_row + 1
      }
      
      if (nb_enfants > 0) {
        writeData(wb, "Devis", "Enfant", startCol = 1, startRow = current_row)
        writeData(wb, "Devis", nb_enfants, startCol = 2, startRow = current_row)
        writeData(wb, "Devis", format(prime_enfant, big.mark = " "), startCol = 3, startRow = current_row)
        writeData(wb, "Devis", format(round(prime_enfant * nb_enfants), big.mark = " "), startCol = 4, startRow = current_row)
        addStyle(wb, "Devis", data_style, rows = current_row, cols = 1:4)
        current_row <- current_row + 1
      }
      
      if (nb_seniors_61_65 > 0) {
        writeData(wb, "Devis", "Adulte (61-65)", startCol = 1, startRow = current_row)
        writeData(wb, "Devis", nb_seniors_61_65, startCol = 2, startRow = current_row)
        writeData(wb, "Devis", format(prime_senior_61_65, big.mark = " "), startCol = 3, startRow = current_row)
        writeData(wb, "Devis", format(round(prime_senior_61_65 * nb_seniors_61_65), big.mark = " "), startCol = 4, startRow = current_row)
        addStyle(wb, "Devis", data_style, rows = current_row, cols = 1:4)
        current_row <- current_row + 1
      }
      
      if (nb_seniors_66_70 > 0) {
        writeData(wb, "Devis", "Adulte (66-70)", startCol = 1, startRow = current_row)
        writeData(wb, "Devis", nb_seniors_66_70, startCol = 2, startRow = current_row)
        writeData(wb, "Devis", format(prime_senior_66_70, big.mark = " "), startCol = 3, startRow = current_row)
        writeData(wb, "Devis", format(round(prime_senior_66_70 * nb_seniors_66_70), big.mark = " "), startCol = 4, startRow = current_row)
        addStyle(wb, "Devis", data_style, rows = current_row, cols = 1:4)
        current_row <- current_row + 1
      }
      
      if (nb_seniors_71_80 > 0) {
        writeData(wb, "Devis", "Adulte (71-80)", startCol = 1, startRow = current_row)
        writeData(wb, "Devis", nb_seniors_71_80, startCol = 2, startRow = current_row)
        writeData(wb, "Devis", format(prime_senior_71_80, big.mark = " "), startCol = 3, startRow = current_row)
        writeData(wb, "Devis", format(round(prime_senior_71_80 * nb_seniors_71_80), big.mark = " "), startCol = 4, startRow = current_row)
        addStyle(wb, "Devis", data_style, rows = current_row, cols = 1:4)
        current_row <- current_row + 1
      }
      
      # Ligne Total
      writeData(wb, "Devis", "TOTAL", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", total_personnes, startCol = 2, startRow = current_row)
      writeData(wb, "Devis", "", startCol = 3, startRow = current_row)
      writeData(wb, "Devis", paste(format(round(nette), big.mark = " "), "XOF"), startCol = 4, startRow = current_row)
      addStyle(wb, "Devis", total_style, rows = current_row, cols = 1:4)
      
      # Récapitulatif final
      current_row <- current_row + 3
      
      # En-tête du récapitulatif
      writeData(wb, "Devis", "RÉCAPITULATIF FINANCIER", startCol = 1, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:4, rows = current_row)
      addStyle(wb, "Devis", recap_header_style, rows = current_row, cols = 1:4)
      current_row <- current_row + 1
      
      # Prime nette globale
      writeData(wb, "Devis", "Prime nette globale", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", paste(format(round(nette), big.mark = " "), "XOF"), startCol = 4, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:3, rows = current_row)
      addStyle(wb, "Devis", recap_label_style, rows = current_row, cols = 1:3)
      addStyle(wb, "Devis", recap_value_style, rows = current_row, cols = 4)
      current_row <- current_row + 1
      
      # Frais et accessoires
      writeData(wb, "Devis", "Frais et Accessoires", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", paste(format(round(commission), big.mark = " "), "XOF"), startCol = 4, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:3, rows = current_row)
      addStyle(wb, "Devis", recap_label_style, rows = current_row, cols = 1:3)
      addStyle(wb, "Devis", recap_value_style, rows = current_row, cols = 4)
      current_row <- current_row + 1
      
      # Taxes
      writeData(wb, "Devis", "Taxes (5%)", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", paste(format(round(taxe), big.mark = " "), "XOF"), startCol = 4, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:3, rows = current_row)
      addStyle(wb, "Devis", recap_label_style, rows = current_row, cols = 1:3)
      addStyle(wb, "Devis", recap_value_style, rows = current_row, cols = 4)
      current_row <- current_row + 1
      
      # PRIME TTC (style spécial en vert)
      writeData(wb, "Devis", "PRIME TTC", startCol = 1, startRow = current_row)
      writeData(wb, "Devis", paste(format(round(ttc), big.mark = " "), "XOF"), startCol = 4, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:3, rows = current_row)
      addStyle(wb, "Devis", ttc_label_style, rows = current_row, cols = 1:3)
      addStyle(wb, "Devis", ttc_value_style, rows = current_row, cols = 4)
      
      # Pied de page
      current_row <- current_row + 3
      footer_text <- paste("Devis généré le", format(Sys.Date(), "%d/%m/%Y"), "- Valable 30 jours")
      writeData(wb, "Devis", footer_text, startCol = 1, startRow = current_row)
      mergeCells(wb, "Devis", cols = 1:4, rows = current_row)
      
      footer_style <- createStyle(
        fontSize = 10,
        fontName = "Arial",
        fontColour = "#7f8c8d",
        halign = "center",
        valign = "center",
        textDecoration = "italic"
      )
      addStyle(wb, "Devis", footer_style, rows = current_row, cols = 1:4)
      
      # Ajustement des largeurs de colonnes
      setColWidths(wb, "Devis", cols = 1:4, widths = c(20, 12, 18, 18))
      setColWidths(wb, "Devis", cols = 5:8, widths = c(2, 12, 12, 12)) # Espace pour le logo
      
      # Ajustement des hauteurs de lignes importantes
      setRowHeights(wb, "Devis", rows = 1, heights = 25)  # Titre principal
      
      # Finalisation et sauvegarde
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # Carte
  output$map <- renderLeaflet({
    if (is.null(input$territoire)) {
      leaflet() %>% 
        addTiles() %>%
        setView(lng = 0, lat = 0, zoom = 2)
    } else if (input$territoire == "Sénégal") {
      leaflet() %>%
        addTiles() %>%
        setView(lng = -14.5, lat = 14.5, zoom = 7.4) 
    } else if (input$territoire == "Monde entier") {
      leaflet() %>%
        addTiles() %>%
        setView(lng = 0, lat = 20, zoom = 2)
    } else if (input$territoire == "Afrique") {
      leaflet() %>%
        addTiles() %>%
        setView(lng = 20, lat = 0, zoom = 3)
    } else {
      leaflet() %>% 
        addTiles() %>%
        setView(lng = 0, lat = 0, zoom = 2)
    }
  })
}

shinyApp(ui, server)