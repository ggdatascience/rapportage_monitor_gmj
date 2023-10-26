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
                   GESLACHT = labelled(sample(x = c(0, 1), size = 10000, replace = T), c('Meisje' = 0, 'Jongen' = 1)), 
                   KLAS = labelled(sample(x = c(0, 1), size = 10000, replace = T), c('KLAS 2' = 0, 'KLAS 4' = 1)),
                   SCHOOL = labelled(sample(x = c(23001, 23002, 23003, 23004, 23005), size = 10000, replace = T),
                                     c('School 1' = 23001, 'School 2' = 23002, 'School 3' = 23003, 'School 4' = 23004, 'School 5' = 23005)),
                   OPLEIDING = labelled(sample(x = c(1, 2, 3), size = 10000, replace = T), c('vmbo' = 1, 'havo' = 2, 'vwo' = 3))) %>%
  mutate(opl_vmbo = ifelse(OPLEIDING == 1, 1, 0),
         opl_vmbohavo = ifelse(OPLEIDING == 3, 0, 1),
         opl_havovwo = ifelse(OPLEIDING == 1, 0, 1),
         opl_vwo = ifelse(OPLEIDING == 3, 1, 0),
         opl_totaal = 1,
         RAPPORT = to_character(SCHOOL))


# 2. Excelbestand met configuratie laden ----------------------------------
rapportconfiguratie <- readxl::read_xlsx("1. Configuratie.xlsx", sheet = 'Rapportconfiguratie')
slideconfiguratie <- readxl::read_xlsx("1. Configuratie.xlsx", sheet = 'Slideconfiguratie 1') %>%
  filter(werkt == 1) %>%
  select(-vraag, -werkt)

# 3. Grafiekinstellingen en opmaak ----------------------------------------

kleuren <- c('#e8525f', '#009898')

# kleuren %>% set_names('KLAS 2', 'KLAS 4')

lettertype <- "TT Arial"

# Thema voor staafgrafiek aanmaken
bartheme <- mschart_theme(
  main_title = fp_text(font.size = 12, bold = TRUE, font.family = lettertype, color = 'black'),
  axis_title_x = fp_text(font.size = 9, font.family = lettertype),
  axis_title_y = fp_text(font.size = 9, font.family = lettertype),
  axis_text_x = fp_text(font.size = 9, font.family = lettertype),
  axis_text_y = fp_text(font.size = 9, font.family = lettertype),
  legend_text = fp_text(font.size = 9, font.family = lettertype),
  grid_major_line = fp_border(style = "none"),
  axis_ticks_x =  fp_border(width = 1),
  axis_ticks_y =  fp_border(width = 1),
  legend_position = 'b'
)

# Grafiekinstellingen bepalen
grafiekstijl <- function(x, grafiektitel = NULL, labelpositie = 'outEnd', labelkleur = 'black', ylimiet = 1) {
  
  x %>% chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = T, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur)) 
  
}


# 4. Functie maken om cijfers te berekenen --------------------------------

## Update Arne 26-10: chi squared toegevoegd aan 'bereken_cijfers' functie
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
  
  # Add chi-squared test results to the result dataframe
  result$p_val <- chi_squared_result$p.value
  result$chi_val <- chi_squared_result$statistic
  result$df <- chi_squared_result$parameter
  
  return(result)
}

bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', groepering = 'KLAS')



# 5. Functies maken om content te genereren -------------------------------

# functie voor losse cijfers
type_percentage <- function(data, rapportnaam, indicator, format = 'percentage', Nvar = 30, Ncel = 30) {
  
  data %>%
    bereken_cijfers(rapportnaam = rapportnaam, indicator = indicator, Nvar = Nvar, Ncel = Ncel) %>%
    mutate(., val = round(val*100)) %>%
    select(val) %>%
    unlist() %>%
    unname() %>%
    {if(is.na(.)) NA 
      else if(format == 'percentage') paste0(., '%') 
      else if(format == 'getal') .}
}


# functie voor staafgrafiek
type_staafgrafiek <- function(data, rapportnaam, indicator, omschrijving = NA, groepering = NA, uitsplitsing = NA, Nvar = 30, Ncel = 30) {
  
  data %>%
    bereken_cijfers(indicator = indicator, rapportnaam = rapportnaam, omschrijving = omschrijving, groepering = groepering, uitsplitsing = uitsplitsing, Nvar = Nvar, Ncel = Ncel) %>%
    rename(var = 1, groep = 2, uitsplitsing = 3) %>%
    ms_barchart(x = 'uitsplitsing', y = 'val', group = 'groep') %>%
    chart_data_fill(kleuren %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
    set_theme(bartheme) %>%
    grafiekstijl(grafiektitel = omschrijving)

}

# Testen van type_grafiek() functie 
# type_staafgrafiek(data, rapportnaam = 'School 2', 'GELUK', omschrijving = 'Schoolbeleving', uitsplitsing = 'GESLACHT', groepering = 'KLAS') %>%
#   print(preview = T)


# 6. Functies aanmaken om content te genereren en te plaatsen -------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_genereren <- function(data, rapportnaam, omschrijving, indicator, uitsplitsing, groepering, vergelijking, type) {
  
  
  if(type == 'rapportnaam') {
  
    rapportnaam
      
  } else if(type == 'percentage') {
    
    type_percentage(data = data, rapportnaam = rapportnaam, indicator = 'GELUK')
  
  } else if(type == 'staafgrafiek') {
    
    type_staafgrafiek(data = data, rapportnaam = rapportnaam, indicator = 'GELUK', 
                      omschrijving = omschrijving, groepering = 'KLAS', uitsplitsing = 'GESLACHT')
    
  }
  
}

# Testen van content_generen() functie
# content_genereren(data = data, rapportnaam = 'School 2', omschrijving = 'bla', indicator = 'GELUK', groepering = 'GESLACHT', type = 'percentage')

# Functie aanmaken om de content te plaatsen. Eerst wordt de content gegenereert met de 
# content_genereren() functie. Daarna wordt de content op basis van de configuratie
# op de juiste slide en de juiste aanduiding geplaatst.
content_plaatsen <- function(data, template, rapportnaam, omschrijving, indicator, uitsplitsing, groepering, vergelijking, type, index, label) {
  
  value <- content_genereren(data = data, rapportnaam = rapportnaam, omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing,
                             groepering = groepering, vergelijking = vergelijking, type = type)
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
 
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(template, slideconfiguratie, data, rapportnaam) {
  
  template <- read_pptx(template)
  pwalk(.l = slideconfiguratie, .f = content_plaatsen, data = data, template = template, rapportnaam = rapportnaam)
  print(template, paste0('Rapport ', rapportnaam, '.pptx'))
  
}


# 8. Rapportage maken via functie -----------------------------------------

rapportage_maken(template = '2. Template.pptx',
                 slideconfiguratie = slideconfiguratie,
                 data = data,
                 rapportnaam = 'School 1')


# 9. Rapportages maken op basis van rapportconfiguratie -------------------

sapply(X = rapportconfiguratie$rapport, FUN = rapportage_maken, template = '2. Template.pptx',
       slideconfiguratie = slideconfiguratie, data = data)
