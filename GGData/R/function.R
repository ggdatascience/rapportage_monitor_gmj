rm(list = ls())
ls()

library(haven) # Package Voor het laden van spss data bestanden
library(labelled) # Package om te werken met datalabels
library(readxl) # Package om te werken met exceldocumenten
library(tidyverse) # Meer informatie over het tidyverse is te vinden op: https://www.tidyverse.org/
library(mschart) # Package voor het aanmaken van office grafieken
library(officer) # Package om Powerp

# Functie grafiekinstellingen bepalen ####
grafiekstijl <- function(x, grafiektitel = NULL, labelpositie = 'outEnd', labelkleur = 'black', ylimiet = 1) {
  x %>% chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = T, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur))
}

# Functie maken om cijfers te berekenen ####
bereken_cijfers <- function(data, indicator, omschrijving = NA, groepering = NA, uitsplitsing = NA, Nvar = 30, Ncel = 30) {
  
  result <- data %>%
    select(all_of(setdiff(c(indicator, groepering, uitsplitsing), NA))) %>%
    group_by(across(all_of(setdiff(c(indicator, groepering, uitsplitsing), NA)))) %>%
    tally() %>%
    drop_na() %>%
    {if(is.na(uitsplitsing) & is.na(groepering)) . 
      else group_by(., across(all_of(setdiff(c(groepering, uitsplitsing), NA))))} %>%
    mutate(ntot = sum(n),
           val = ifelse(ntot < Nvar | min(n) < Ncel | n == ntot, NA, n/ntot) %>% as.numeric()) %>%
    filter_at(vars(all_of(indicator)), function(x) x == 1) %>%
    ungroup() %>%
    mutate(across(!n & !ntot & !val, to_character),
           n_min = ntot - n) %>%
    rename(var = 1) %>%
    mutate(var = omschrijving)
  
  # Perform chi-squared test
  chi_squared_result <- chisq.test(result[, c("n", "n_min")])
  
  hoogste_waarde <- which.max(result$val)
  rij_met_hoogste_waarde <- result[hoogste_waarde, ]
  laagste_waarde <- which.min(result$val)
  rij_met_laagste_waarde <- result[laagste_waarde, ]
  
  # Extract 'omschrijving' for highest and lowest values
  hoogste_omschrijving <- rij_met_hoogste_waarde$var
  laagste_omschrijving <- rij_met_laagste_waarde$var
  hoogste_categorie <- rij_met_hoogste_waarde[2]
  laagste_categorie <- rij_met_laagste_waarde[2]
  
    # Add chi-squared test results to the result dataframe
  result$p_val <- chi_squared_result$p.value
  result$chi_val <- chi_squared_result$statistic
  result$df <- chi_squared_result$parameter
  
  # Create a text description string
  description_text <- paste(hoogste_categorie,"scoort hoger op", hoogste_omschrijving, "dan", laagste_categorie, "dit verschil is statistisch",ifelse(result$p_val[1]<= 0.05,'significant','niet significant'))
  
  
  # Create a list to store and return multiple results
  result_list <- list(result = result, description = description_text)
  
  return(result_list)
}

# Functie voor losse cijfers ####
type_percentage <- function(data, indicator, format = 'percentage', Nvar = 30, Ncel = 30) {
  data %>%
    bereken_cijfers(indicator = indicator, Nvar = Nvar, Ncel = Ncel) %>%
    mutate(., val = round(val*100)) %>%
    select(val) %>%
    unlist() %>%
    unname() %>%
    {if(is.na(.)) NA
      else if(format == 'percentage') paste0(., '%')
      else if(format == 'getal') .}
}

# Functie voor staafgrafiek ####
type_staafgrafiek <- function(data, indicator, omschrijving = NA, groepering = NA, uitsplitsing = NA, Nvar = 30, Ncel = 30) {
  data %>%
    bereken_cijfers(indicator = indicator, omschrijving = omschrijving, groepering = groepering, uitsplitsing = uitsplitsing, Nvar = Nvar, Ncel = Ncel) %>%
    rename(var = 1, groep = 2, uitsplitsing = 3) %>%
    ms_barchart(x = 'uitsplitsing', y = 'val', group = 'groep') %>%
    chart_data_fill(kleuren %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
    set_theme(bartheme) %>%
    grafiekstijl(grafiektitel = omschrijving)

}

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert ####
content_genereren <- function(data, omschrijving, indicator, groepering, type) {
  if(type == 'percentage') {
    type_percentage()
  } else if(type == 'staafgrafiek') {
    type_staafgrafiek()
  }
}

# Functie aanmaken om de content te plaatsen ####
# Eerst wordt de content gegenereert met de
# content_genereren() functie. Daarna wordt de content op basis van de configuratie
# op de juiste slide en de juiste aanduiding geplaatst.
content_plaatsen <- function(data, template, omschrijving, indicator, groepering, type, index, label) {
  value <- content_genereren(data = data, omschrijving = omschrijving, indicator = indicator, groepering = groepering, type = type)
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))

}

# Functie aanmaken om de rapportage te maken ####
rapportage_maken <- function(template, configuratie, data, rapport) {
  template <- read_pptx(template)
  pwalk(.l = configuratie, .f = content_plaatsen, data = data, template = template)
  print(template, rapport)
}

