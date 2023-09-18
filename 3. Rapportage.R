# Script Rapportage

# 0. Voorbereiding --------------------------------------------------------

# Environment leegmaken
rm(list = ls())

# Libraries laden
# Deze libraries moeten eenmalig worden geinstalleerd met install.packages()
library(haven) # Package Voor het laden van spss data bestanden
library(labelled) # Package om te werken met datalabels 
library(readxl) # Package om te werken met exceldocumenten
library(tidyverse) # Meer informatie over het tidyverse is te vinden op: https://www.tidyverse.org/
library(mschart) # Package voor het aanmaken van office grafieken
library(officer) # Package om Powerpoint (en andere MS Office) documenten aan te maken of te bewerken vanuit R


# 1. Nepdata genereren ----------------------------------------------------
data <- data.frame(GELUK = sample(x = c(NA, 0, 1), size = 10000, replace = T, prob = c(0.05, 0.15, 0.8)),
                   ERVGEZ = sample(x = c(NA, 0, 1), size = 10000, replace = T, prob = c(0.07, 0.2, 0.73)),
                   GESLACHT = sample(x = c('Jongen', 'Meisje'), size = 10000, replace = T))


# 2. Excelbestand met configuratie laden ----------------------------------
configuratie <- readxl::read_xlsx("1. Configuratie.xlsx")


# 3. Functies aanmaken om content te genereren en te plaatsen -------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
# Voor nu heb ik de content eenvoudig gehouden, maar voor elk 'type' kan een functie worden 
# aangemaakt om dat type content te genereren.
# Bij dit voorbeeld zijn er 3 verschillende type content (tekst, percentage en
# staafgrafiek). De content voor het type 'staafgrafiek' is nu nog niet logisch, de data
# moet nog worden bewerkt en de figuuropmaak moet nog worden toegevoegd.
content_genereren <- function(data, omschrijving, indicator, groepering, type) {
  
  if(type == 'tekst') {
  
    omschrijving
      
  } else if(type == 'percentage') {
    
    mean(data[[indicator]]*100, na.rm = T) %>%
      round(0) %>%
      paste0('%')
    
  } else if(type == 'staafgrafiek') {
    
    ms_barchart(data = data, x = groepering, y = indicator)
  
  }
  
}

# Functie aanmaken om de content te plaatsen. Eerst wordt de content gegenereert met de 
# content_genereren() functie. Daarna wordt de content op basis van de configuratie
# op de juiste slide en de juiste aanduiding geplaatst.
content_plaatsen <- function(data, template, omschrijving, indicator, groepering, type, index, label) {
  
  value <- content_genereren(data = data, omschrijving = omschrijving, indicator = indicator, groepering = groepering, type = type)
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
 
}


# 4. Rapportage maken -----------------------------------------------------

# Template laden
template <- read_pptx('2. Template.pptx')

# Rapportage maken, waarbij de functie content_plaatsen() wordt toegepast op
# de configuratie.
pwalk(.l = configuratie, .f = content_plaatsen, data = data, template = template)

# Rapportage opslaan als .pptx
print(template, 'Rapportage.pptx')

# Alle stappen onder stap 4 moeten samengevoegd worden tot 1 rapportage functie.
# Deze functie kan vervolgens worden toegepast om een x aantal rapportages aan te
# maken (bijvoorbeeld voor scholen of gemeenten).


# 5. Handmatig template vullen --------------------------------------------

# Onderstaande code kan gebruikt worden om handmatig een rapportage te vullen
# zodat je kan oefenen met deze manier van Powerpoints aanmaken
read_pptx('2. Template.pptx') %>%
  on_slide(index = 1) %>%
  ph_with(value = 'Vul hier zelf een value in', location = ph_location_label(ph_label = 'Titel')) %>%
  print('Oefenen met handmatig vullen.pptx')
