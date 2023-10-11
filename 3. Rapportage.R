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
         opl_totaal = 1)


# 2. Excelbestand met configuratie laden ----------------------------------
rapporten <- readxl::read_xlsx("1. Configuratie Figuren.xlsx", sheet = 'Rapportconfiguratie')
configuratie <- readxl::read_xlsx("1. Configuratie Figuren.xlsx", sheet = 'Slideconfiguratie 1')


# 3. Grafiekinstellingen en opmaak ----------------------------------------

kleuren <- c('#e8525f', '#009898')

kleuren %>% set_names('KLAS 2', 'KLAS 4')

val_labels(data$KLAS) %>% names()

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

bereken_cijfers <- function(data, indicator, omschrijving = NA, groepering = NA, uitsplitsing = NA, Nvar = 30, Ncel = 30) {
  
  data %>%
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
    mutate(across(!n & !ntot & !val, to_character)) %>%
    rename(var = 1) %>%
    mutate(var = omschrijving)
  
}

bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS')


# 5. Functies maken om content te genereren -------------------------------

# functie voor losse cijfers
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

type_percentage(data, 'GELUK',)

# functie voor staafgrafiek
type_staafgrafiek <- function(data, indicator, omschrijving = NA, groepering = NA, uitsplitsing = NA, Nvar = 30, Ncel = 30) {
  
  data %>%
    bereken_cijfers(indicator = indicator, omschrijving = omschrijving, groepering = groepering, uitsplitsing = uitsplitsing, Nvar = Nvar, Ncel = Ncel) %>%
    rename(var = 1, groep = 2, uitsplitsing = 3) %>%
    ms_barchart(x = 'uitsplitsing', y = 'val', group = 'groep') %>%
    chart_data_fill(kleuren %>% set_names(val_labels(data[[groepering]]) %>% names())) %>%
    set_theme(bartheme) %>%
    grafiekstijl(grafiektitel = omschrijving)

}

?ms_barchart


type_staafgrafiek(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS') %>%
  print(preview = T)


cijfers %>%
  {if(stapel) rename(., groep = 1) else .} %>%
  {if(gr | geen) mutate(., var = 'Totaal') else .} %>%
  ms_barchart(., 
              x = {if((uit | uitgr) | (stapel & uit)) uitsplitsing else if(gr | geen | stapel) 'var'}, 
              y = 'val', 
              group = {if(stapel) 'groep' else if(uitgr | gr) groep else if(uit | geen) NULL}) %>%
  {if(stapel) as_bar_stack(., dir = "horizontal") else .} %>%
  chart_data_fill(grafiekkleuren) %>%
  set_theme(bartheme) %>%
  chart_theme(legend_position = {if(uitgr | gr | stapel) 'b' else if(uit | geen) 'n'}) %>%
  grafiekstijl(grafiektitel = configuratie$omschrijving[j],
               labelpositie = {if(stapel) 'ctr' else 'outEnd'}, 
               labelkleur = {if(stapel) 'white' else 'black'})


# 6. Functies aanmaken om content te genereren en te plaatsen -------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_genereren <- function(data, omschrijving, indicator, groepering, type) {
  
  if(type == 'percentage') {
  
    type_percentage()
      
  } else if(type == 'staafgrafiek') {
    
    type_staafgrafiek()
  
  }
  
}  

content_genereren(data = data, omschrijving = 'bla', indicator = 'GELUK', groepering = 'GESLACHT', type = 'percentage')

# Functie aanmaken om de content te plaatsen. Eerst wordt de content gegenereert met de 
# content_genereren() functie. Daarna wordt de content op basis van de configuratie
# op de juiste slide en de juiste aanduiding geplaatst.
content_plaatsen <- function(data, template, omschrijving, indicator, groepering, type, index, label) {
  
  value <- content_genereren(data = data, omschrijving = omschrijving, indicator = indicator, groepering = groepering, type = type)
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
 
}


# 7. Rapportage maken -----------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(template, configuratie, data, rapport) {
  
  template <- read_pptx(template)
  pwalk(.l = configuratie, .f = content_plaatsen, data = data, template = template)
  print(template, rapport)
  
}


# 8. Rapportage maken via functie -----------------------------------------

rapportage_maken('2. Template.pptx',configuratie,data,'Overzicht figuren.pptx')
