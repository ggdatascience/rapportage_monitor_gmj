# Script Figurenboek

# 0. Voorbereiding --------------------------------------------------------

# Environment leegmaken
rm(list = ls())

# Libraries laden
# Deze libraries moeten eenmalig worden geinstalleerd met install.packages()
library(extrafont) # Package om het lettertype Century Gothic te gebruiken
library(haven) # Package Voor het laden van spss data bestanden
library(labelled) # Package om te werken met datalabels 
library(readxl) # Package om te werken met exceldocumenten
library(tidyverse) # Meer informatie over het tidyverse is te vinden op: https://www.tidyverse.org/
library(mschart) # Package voor het aanmaken van office grafieken
library(officer) # Package om Powerpoint (en andere MS Office) documenten aan te maken of te bewerken vanuit R
library(flextable) # Package om tabellen te maken

# Working directory bepalen
# Je kunt op meerdere manieren de werkmap bepalen waarin je documenten staan.
# In het paneel met het tabbled Files (meestal het paneel rechtsonder) zie je de map die R als werkmap beschouwd.
# Je kunt met de functie getwd() de padnaam zien van je werkmap.
#getwd()
#dir()

# De werkmap is meestal de map van waaruit je een R script opent. Is dit niet zo dan kun je met setwd() je werkmap bepalen.
#setwd('C:/padnaam')


# 1. Padnaam voor configuratie opgeven ------------------------------------

# Padnaam opgeven van de Excel configuratie. Staat dit bestand in je working directory dan hoef je alleen de bestandsnaam
# op te geven (inclusief de bestandsextensie)
padnaam_configuratie <- '1. Configuratie Figurenboek.xlsx'

# 2. Dataconfiguratie ----------------------------------------------------

# Data inladen
data <- read_sav( "C:/padnaam/databestand.sav")

# 3. Grafiekinstellingen en opmaak ----------------------------------------

# Kleurpalet aanmaken
kleuren <- c('#009898', '#91caca', '#b9b441', '#dcd9a6', '#e8525f', '#e9acb0')

# Lettertype importeren. In de console wordt gevraagd of je dit wil doen run 'y' in de console om te importeren.
extrafont::font_import(pattern = 'GOTHIC')

# Check om te zien of het lettertype goed is ingeladen (check of het lettertype Century Gothic ertussen staat)
windowsFonts()

# Lettertype bepalen
lettertype <- "Century Gothic"

# Thema voor staafgrafiek aanmaken
chart_theme <- mschart_theme(
  main_title = fp_text(font.size = 9, bold = TRUE, font.family = lettertype, color = 'black'),
  axis_title_x = fp_text(font.size = 9, font.family = lettertype),
  axis_title_y = fp_text(font.size = 9, font.family = lettertype),
  axis_text_x = fp_text(font.size = 8, font.family = lettertype),
  axis_text_y = fp_text(font.size = 9, font.family = lettertype),
  legend_text = fp_text(font.size = 9, font.family = lettertype),
  grid_major_line = fp_border(style = "none"),
  axis_ticks_x =  fp_border(width = 1),
  axis_ticks_y =  fp_border(width = 1),
  legend_position = 't'
)

# Grafiekinstellingen bepalen
grafiekstijl <- function(x, grafiektitel = NULL, datalabels = T, labelpositie = 'outEnd', labelkleur = 'black', ylimiet = 1) {
  
  x %>% chart_data_stroke('transparent') %>%
    chart_labels(title = grafiektitel, xlab = NULL, ylab = NULL) %>%
    chart_ax_x(major_tick_mark = 'none') %>%
    chart_ax_y(limit_min = 0, num_fmt = '0%', major_tick_mark = 'none', limit_max = ylimiet) %>%
    chart_data_labels(show_val = datalabels, num_fmt = '0%', position = labelpositie) %>%
    chart_labels_text(fp_text(font.size = 8, font.family = lettertype, color = labelkleur)) 
  
}

# 4. Functie maken om cijfers te berekenen --------------------------------

bereken_cijfers <- function(data,
                            basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                            omschrijving = NA, indicator, waarden = 1, valuelabel = NA, uitsplitsing = NA, groepering = NA, niveau, jaar = NA, 
                            var_jaar = 'Onderzoeksjaar', Nvar = 30, Ncel = 5, toetsen = F) {
  
  niveau_var <- {if(niveau == 'basis') basis else if(niveau == 'referentie') referentie}
  niveau_label <- {if(niveau == 'basis') basis_label else if(niveau == 'referentie') referentie_label}
  
  result <- data %>%
    filter(.data[[niveau_var]] == niveau_label & .data[[var_jaar]] %in% jaar) %>%
    select(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA))) %>%
    group_by(across(all_of(setdiff(c(var_jaar, indicator, uitsplitsing, groepering), NA))), .drop = T) %>%
    tally() %>%
    drop_na() %>%
    {if(is.na(uitsplitsing) & is.na(groepering)) .
      else group_by(., across(all_of(setdiff(c(uitsplitsing, groepering), NA))))} %>%
    mutate(ntot = sum(n),
           nmin = ntot-n,
           val = ifelse(ntot < Nvar | n < Ncel | nmin < Ncel | n == ntot, NA, n/ntot) %>% as.numeric(),
           indicator = indicator,
           omschrijving = omschrijving,
           niveau = {if(niveau == 'basis') 'Gemeente' else if(niveau == 'referentie') 'Totaal'}) %>%
    filter_at(vars(all_of(indicator)), function(x) x %in% waarden) %>%
    ungroup() %>%
    rename(jaar = 1, aslabel = 2) %>%
    {if(!is.na(valuelabel)) mutate(., aslabel = valuelabel) else .} %>%
    {if(!is.na(uitsplitsing)) rename(., uitsplitsing = .data[[uitsplitsing]]) else .} %>%
    {if(!is.na(groepering)) rename(., groepering = .data[[groepering]]) else .} %>%
    mutate(across(!n & !ntot & !nmin & !val & !jaar, to_character)) %>%
    select(indicator, omschrijving, jaar, niveau, aslabel, everything())
  
  # Perform chi-squared test
  
  if(toetsen == T){
    
    chi_squared_result <- chisq.test(result[, c("n", "nmin")])
    
    hoogste_waarde <- which.max(result$val)
    rij_met_hoogste_waarde <- result[hoogste_waarde, ]
    laagste_waarde <- which.min(result$val)
    rij_met_laagste_waarde <- result[laagste_waarde, ]
    
    # Extract 'omschrijving' for highest and lowest values
    hoogste_omschrijving <- rij_met_hoogste_waarde$omschrijving
    laagste_omschrijving <- rij_met_laagste_waarde$omschrijving
    hoogste_categorie <- rij_met_hoogste_waarde[2]
    laagste_categorie <- rij_met_laagste_waarde[2]
    
    # Add chi-squared test results to the result dataframe
    result$p_val <- chi_squared_result$p.value
    result$chi_val <- chi_squared_result$statistic
    result$df <- chi_squared_result$parameter
    
  }
  
  return(result)
  
  
}


# 5. Functies maken om content te genereren -------------------------------

# functie voor staafgrafiek
type_staafgrafiek <- function(data,
                              basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                              omschrijving = NA, indicator, waarden = 1, uitsplitsing = NA, groepering = NA, niveau, jaar) {
  
  if(niveau == 'basis' | niveau == 'referentie') {
    
    cijfers <- data %>%
      bereken_cijfers(basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                      indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering)
    
    if(!is.na(groepering)) {
     
    legendakleuren <- kleuren[1:length(unique(cijfers$groepering))] %>% set_names(unique(cijfers$groepering))
    
    }
    
    {if(!is.na(uitsplitsing) & !is.na(groepering))
      ms_barchart(cijfers, x = 'uitsplitsing', y = 'val', group = 'groepering') %>%
        chart_data_fill(legendakleuren) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'r') %>%
        grafiekstijl(grafiektitel = omschrijving)
      
      else if(!is.na(uitsplitsing) & is.na(groepering))
        ms_barchart(cijfers, x = 'uitsplitsing', y = 'val') %>%
        chart_data_fill(kleuren[1]) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'n') %>%
        grafiekstijl(grafiektitel = omschrijving,
                     ylimiet = ifelse((.$data %>% .$val %>% max(na.rm = T)) < 0.5, 0.5, 1))
      
      else if(is.na(uitsplitsing) & !is.na(groepering))
        ms_barchart(cijfers, x = 'omschrijving', y = 'val', group = 'groepering') %>%
        chart_data_fill(legendakleuren) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'r') %>%
        chart_ax_x(tick_label_pos = 'none') %>%
        grafiekstijl(grafiektitel = omschrijving)
      
      else
        ms_barchart(cijfers, x = 'omschrijving', y = 'val') %>%
        chart_data_fill(kleuren[1]) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'n') %>%
        chart_ax_x(tick_label_pos = 'none') %>%
        grafiekstijl(grafiektitel = omschrijving)
    }
    
  } else if(niveau == 'basis en referentie') {
    
    data.frame(niveau = c('basis', 'referentie'),
               basis = c(basis, NA),
               basis_label = c(basis_label, NA),
               referentie = c(NA, referentie),
               referentie_label = c(NA, referentie_label)) %>%
      pmap(bereken_cijfers, data = data, indicator = indicator, waarden = waarden, jaar = jaar, 
           omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
      reduce(bind_rows) %>%
      mutate(niveau = factor(niveau, levels = c('Gemeente', 'Totaal'), ordered = T)) %>%
      ms_barchart(x = 'aslabel', y = 'val', group = 'niveau') %>% # x is nu 'aslabel' en niet 'omschrijving' kan mogelijk problemen opleveren
      chart_data_fill(kleuren[c(1,5)] %>% set_names(c('Gemeente', 'Totaal'))) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      chart_ax_x(tick_label_pos = 'none') %>%
      chart_settings(overlap = -20) %>%
      grafiekstijl(grafiektitel = omschrijving)
    
  }
}

# functie voor combigrafiek
type_combi <- function(data,
                       basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                       omschrijving = NA, indicator, waarden = 1, uitsplitsing = NA, groepering = NA, niveau, jaar,
                       valuelabel = NA, horizontaal = F, selectie = F, selectie_n = NA) {
  
   cijfers <- data.frame(indicator = indicator, valuelabel = valuelabel) %>%
     pmap(bereken_cijfers, data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
         waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, uitsplitsing = uitsplitsing, groepering = groepering) %>%
     reduce(bind_rows) %>%
     mutate(aslabel = factor(aslabel, levels = unique(aslabel), ordered = T))
   
   if(!is.na(groepering)) {
     
     legendakleuren <- kleuren[1:length(unique(cijfers$groepering))] %>% set_names(unique(cijfers$groepering))
     
   }
   
   cijfers %>%
     {if(horizontaal == T & !is.na(groepering)) mutate(., groepering = fct_rev(groepering), aslabel = fct_rev(aslabel)) else .} %>%
     {if(selectie == T) arrange(., desc(val)) %>%
        top_n(n = selectie_n) else .} %>%
     {if(!is.na(groepering))
      ms_barchart(.,x = 'aslabel', y = 'val', group = 'groepering') %>%
        chart_data_fill(legendakleuren) %>%
        set_theme(chart_theme)

       else ms_barchart(.,x = 'aslabel', y = 'val') %>%
        chart_data_fill(kleuren[1]) %>%
        set_theme(chart_theme) %>%
        chart_theme(legend_position = 'n')} %>%
     grafiekstijl(grafiektitel = omschrijving, ylimiet = 1) %>%
     {if(horizontaal == T) chart_settings(., dir = "horizontal") else .}
}

# functie voor gestapelde staafgrafiek
type_staafgrafiek_gestapeld <- function(data,
                                        basis = NA, basis_label = NA, referentie = NA, referentie_label = NA,
                                        omschrijving = NA, indicator, waarden = NA, niveau, jaar,
                                        valuelabel = NA, horizontaal = F) {
  
  if(niveau == 'basis' | niveau == 'referentie') {
    
    cijfers <- bereken_cijfers(data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                               indicator = indicator, waarden = waarden, niveau = niveau, jaar = jaar, omschrijving = omschrijving, valuelabel = valuelabel) %>%
      mutate(aslabel = factor(aslabel, levels = unique(aslabel), ordered = T))
    
    legendakleuren <- kleuren[1:length(cijfers$aslabel)] %>% set_names(cijfers$aslabel)
    
    ms_barchart(cijfers, x = 'niveau', y = 'val', group = 'aslabel') %>% 
      {if(horizontaal == T) as_bar_stack(., dir = 'horizontal') else .} %>%
      chart_settings(grouping = "stacked", gap_width = 50, overlap = 100 ) %>%
      chart_data_fill(legendakleuren) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      grafiekstijl(grafiektitel = omschrijving,
                   ylimiet = 1,
                   labelpositie = 'ctr',
                   labelkleur = 'white')
    
  } else if(niveau == 'basis en referentie') {
    
    cijfers <- data.frame(niveau = c('basis', 'referentie'),
                          basis = c(basis, NA),
                          basis_label = c(basis_label, NA),
                          referentie = c(NA, referentie),
                          referentie_label = c(NA, referentie_label)) %>%
      pmap(bereken_cijfers, data = data, indicator = indicator, waarden = waarden, jaar = jaar,
           omschrijving = omschrijving) %>%
      reduce(bind_rows) %>%
      mutate(niveau = factor(niveau, levels = c('Gemeente', 'Totaal'), ordered = T),
             aslabel = factor(aslabel, levels = unique(aslabel), ordered = T))
    
    legendakleuren <- kleuren[1:length(unique(cijfers$aslabel))] %>% set_names(unique(cijfers$aslabel))
    
    ms_barchart(cijfers, x = 'niveau', y = 'val', group = 'aslabel') %>%
      {if(horizontaal == T) as_bar_stack(., dir = 'horizontal') else .} %>%
      chart_settings(grouping = "stacked", gap_width = 50, overlap = 100 ) %>%
      chart_data_fill(legendakleuren) %>%
      set_theme(chart_theme) %>%
      chart_theme(legend_position = 'b') %>%
      grafiekstijl(grafiektitel = omschrijving,
                   ylimiet = 1,
                   labelpositie = 'ctr',
                   labelkleur = 'white')
    
  }
}


# 6. Functie aanmaken om content te genereren en te plaatsen --------------

# Functie aanmaken die op basis van de kolom 'type' uit de configuratie content genereert.
content_plaatsen <- function(data, template, 
                             basis, basis_label, referentie, referentie_label, 
                             omschrijving, type, indicator, waarden, valuelabel, uitsplitsing, groepering, niveau, jaar, index, label) {
  
  if(type == 'rapportnaam') {
    
    value <- basis_label
    
  } else if(type == 'combi') {
    
    value <- type_combi(data = data,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel, referentie = referentie, referentie_label = referentie_label,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 

  } else if(type == 'combi liggend') {
    
    value <- type_combi(data = data, horizontaal = T,
                        basis = basis, basis_label = basis_label, valuelabel = valuelabel, referentie = referentie, referentie_label = referentie_label,
                        omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, waarden = waarden, niveau = niveau, jaar = jaar) 
    
  } else if(type == 'staafgrafiek') {
    
    value <- type_staafgrafiek(data = data,
                               basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                               omschrijving = omschrijving, indicator = indicator, uitsplitsing = uitsplitsing, groepering = groepering, 
                               waarden = waarden, niveau = niveau, jaar = jaar)
    
  } else if(type == 'staafgrafiek gestapeld') {
    
    value <- type_staafgrafiek_gestapeld(data = data, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label,
                                         niveau = niveau, jaar = jaar, waarden = waarden,
                                         indicator = indicator, omschrijving = omschrijving, horizontaal = T)
    
  }
  
  on_slide(template, index = index) %>%
    ph_with(value = value, location = ph_location_label(ph_label = label))
  
}


# 7. Rapportage functie  --------------------------------------------------

# Functie aanmaken om de rapportage te maken. 
rapportage_maken <- function(data, template, configuratie, basis, basis_label, referentie, referentie_label, slideconfiguratie, filteren = F) {
  
  template <- read_pptx(template)
  
  walk(configuratie[[slideconfiguratie]] %>% select(indeling, index) %>% unique() %>% pull(indeling), add_slide, x = template, master = 'Figurenboek')

  slideconfiguratie <- configuratie[[slideconfiguratie]] %>%
    {if(length(filteren) == 1 & !is.numeric(filteren)) . else slice(., -filteren)} %>%
    select(omschrijving, type, indicator, waarden, valuelabel, uitsplitsing, groepering, niveau, jaar, index, label) %>%
    mutate(across(c(indicator, waarden, valuelabel, jaar), ~ .x %>% str_split(';\\s*')))

  pwalk(.l = slideconfiguratie, .f = content_plaatsen, data = data, template = template, basis = basis, basis_label = basis_label, referentie = referentie, referentie_label = referentie_label)
  print(template, paste0('Figurenboek ', basis_label, '.pptx'))
  
}


# 8. Rapportages maken op basis van rapportconfiguratie -------------------

# Inladen van alle tabbladen van de Excel configuratie
configuratie <- padnaam_configuratie %>% 
  excel_sheets() %>% 
  set_names() %>% 
  map(read_excel, path = padnaam_configuratie)

# Maken van figurenboeken
pmap(.l = configuratie$Rapportconfiguratie,
     .f = rapportage_maken, data = data, configuratie = configuratie, template = '2. Template Figurenboek.pptx')
